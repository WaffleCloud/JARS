/*******************************
 * fullAbsenceRoute.gs (Same-Day Full Absence)
 * ✅ De-duped to rely on utils.gs for shared helpers:
 *   - getConfig_()
 *   - slackApi_(), json_(), debugLog_(), mark_()
 *   - ensureWorkerTriggerOnce_()
 *   - qPush_(), qShiftBatch_()
 *   - getMultiSelectValues_(), getPlainTextValue_(), getRadioValue_()
 *   - coverageOptionGroups_(), makeOption()
 *   - packMeta_(), unpackMeta_()
 *   - isBefore615(), formatIsoLocal_()
 *   - getSlackUserLabel_()  (if you have the SlackUsers mapping section)
 *
 * Fixes included:
 * 1) Removes duplicated helper functions that were shadowing utils.gs
 * 2) Fixes naming mismatch: uses isBefore615() (no underscore) from utils
 * 3) Adds explicit debug + SMOKE markers around sheet writes and row numbers
 * 4) Uses correct HTTP 200 on Slack “errors”
 *******************************/


/* ===========================
   VIEW SUBMISSION (FAST) => validate + enqueue
   =========================== */

function handleFullAbsenceSubmitFast_(cfg, payload) {
  try{

    const view = payload?.view || {}
    const state = view?.state?.values || {}
    
    const coverageValues = getMultiSelectValues_(state, 'coverage_block', 'coverage')
    const hasSubPlans = getRadioValue_(state, 'subplans_yn_block', 'subplans_yn')
    const subPlansLocation = getPlainTextValue_(state, 'subplans_location_block', 'subplans_location')
    
    const errors = {}
    
    if (!coverageValues || !coverageValues.length) errors.coverage_block = 'Required.'
    if (!hasSubPlans) errors.subplans_yn_block = 'Required.'
    if (!String(subPlansLocation || '').trim()) errors.subplans_location_block = 'Required. If none, type "none" or "n/a".'
    
    if (Object.keys(errors).length) {
      return json_(200, { response_action: 'errors', errors })
    }
    
    const now = new Date()
    const submittedBefore615 = isBefore615(now, cfg.TZ) === true
    
    qPush_(cfg.FULL_ABSENCE_SUBMISSION_QUEUE_KEY, {
      type: 'FULL_ABSENCE', // optional safety tag
      submittedAtIso: now.toISOString(),
      userId: payload?.user?.id || '',
      username: payload?.user?.username || payload?.user?.name || '',
      coverageValues: coverageValues || [],
      hasSubPlans: hasSubPlans, // YES / NO
      subPlansLocation: String(subPlansLocation || '').trim(),
      submittedBefore615
    }, 250, 60 * 30)
    
    return json_(200, { response_action: 'clear' })
  }catch(err){
        // IMPORTANT: do not debugLog_ here (sheets). Just clear.
    try { Logger.log('handleFullAbsenceSubmitFast_ err=' + String(err && err.stack || err)) } catch (e) {}
    return json_(200, { response_action: 'clear' })
  }
}



/* ===========================
   SUBMISSION WORKER
   Flush to sheet -> enqueue SEND_MESSAGES jobs
   =========================== */

function submissionWorker_() {
  if (isSlackHot_()) return
  const cfg = getConfig_()

  const batch = qShiftBatch_(cfg.FULL_ABSENCE_SUBMISSION_QUEUE_KEY, 50)
  if (!batch.length) return

  mark_('FULL_submissionWorker_start', { n: batch.length })

  try {
    const ss = SpreadsheetApp.openByUrl(cfg.SHEET_URL)
    const sh = ss.getSheetByName(cfg.FULL_ABSENCE_SHEET_NAME) || ss.insertSheet(cfg.FULL_ABSENCE_SHEET_NAME)

    // Ensure headers exist
    if (sh.getLastRow() === 0) {
      sh.appendRow([
        'Submitted At',
        'Slack User ID',
        'Slack Username',
        'Opened At (ISO)',
        'Coverage Needed',
        'Has Sub Plans (Y/N)',
        'Sub Plans Location',
        'Submitted Before 6:15am'
      ])
    }

    const startRow = sh.getLastRow() + 1

    const rows = batch.map(r => ([
      new Date(r.submittedAtIso),
      r.userId || '',
      r.username || '',
      r.openedAtIso || '',
      (r.coverageValues || []).join(', '),
      r.hasSubPlans || '',
      r.subPlansLocation || '',
      r.submittedBefore615 === true
    ]))

    sh.getRange(startRow, 1, rows.length, 8).setValues(rows)
    SpreadsheetApp.flush()

    // 🔥 This is the “receipt data” write — log it hard
    debugLog_(cfg, 'FULL_submissionWorker', `wrote rows=${rows.length} startRow=${startRow} sheet=${cfg.FULL_ABSENCE_SHEET_NAME}`)
    mark_('FULL_submissionWorker_wrote', { startRow, rows: rows.length })

    // enqueue jobs AFTER successful write
    batch.forEach((r, i) => {
      qPush_(cfg.FULL_ABSENCE_JOB_QUEUE_KEY, {
        kind: 'SEND_MESSAGES',
        receiptRow: startRow + i,
        submittedAtIso: r.submittedAtIso,
        userId: r.userId,
        username: r.username,
        coverageValues: r.coverageValues || [],
        hasSubPlans: r.hasSubPlans,
        subPlansLocation: r.subPlansLocation,
        submittedBefore615: r.submittedBefore615 === true
      }, 250, 60 * 30)
    })

    mark_('FULL_submissionWorker_enqueued_jobs', { n: batch.length })
  } catch (err) {
    const msg = String(err && err.stack || err)

    mark_('FULL_submissionWorker_ERROR', { err: msg })
    debugLog_(cfg, 'FULL_submissionWorker_ERROR', msg)

    // put the batch back so it can retry
    batch.forEach(item => {
      qPush_(cfg.FULL_ABSENCE_SUBMISSION_QUEUE_KEY, item, 250, 60 * 30)
    })

    // also push into debug queue for debugWorker_ to persist
    qPush_(cfg.DEBUG_QUEUE_KEY, {
      ts: new Date().toISOString(),
      source: 'submissionWorker_',
      message: 'EXCEPTION',
      data: { err: msg }
    }, 250, 60 * 30)
  }
}


/* ===========================
   JOB WORKER
   Slack API calls that can be delayed
   =========================== */

function jobWorker_() {
  if (isSlackHot_()) return
  const cfg = getConfig_()

  const jobs = qShiftBatch_(cfg.FULL_ABSENCE_JOB_QUEUE_KEY, 25)
  if (!jobs.length) return

  jobs.forEach(job => {
    try {
      if (job.kind === 'SEND_MESSAGES') sendNotifyAndReceipt_(job)
    } catch (err) {
      qPush_(cfg.DEBUG_QUEUE_KEY, {
        ts: new Date().toISOString(),
        source: 'jobWorker_',
        message: 'Job failed',
        data: { job, err: String(err && err.stack || err) }
      }, 250, 60 * 30)
    }
  })
}


/* ===========================
   SEND NOTIFY + DM RECEIPT (worker only)
   =========================== */

function sendNotifyAndReceipt_(job) {
  const cfg = getConfig_()

  // Prefer your SlackUsers mapping helper if present; fallback otherwise
  const who = (typeof getSlackUserLabel_ === 'function')
    ? getSlackUserLabel_(job.userId, job.username)
    : (job.username ? `@${job.username}` : (job.userId || 'Unknown'))

  const coverageText = (job.coverageValues && job.coverageValues.length)
    ? job.coverageValues.join(', ')
    : 'None'

  const ptoPolicyLine = job.submittedBefore615
    ? 'Submitted before 6:15am (will be processed as PTO if available)'
    : 'Submitted after 6:15am (HR time off violation; does not qualify for PTO)'

  // Notify channel
  if (cfg.FULL_ABSENCE_NOTIFY_CHANNEL) {
    const resp = slackApi_(cfg, 'chat.postMessage', {
      channel: cfg.FULL_ABSENCE_NOTIFY_CHANNEL,
      text: 'Same-Day Full Absence',
      blocks: [
        { type: 'section', text: { type: 'mrkdwn', text: '*Same-Day Full Absence*' } },
        {
          type: 'section',
          fields: [
            { type: 'mrkdwn', text: '*Who:*\n' + who },
            { type: 'mrkdwn', text: '*Receipt ID:*\nRow ' + job.receiptRow }
          ]
        },
        {
          type: 'section',
          fields: [
            { type: 'mrkdwn', text: '*Coverage Needed:*\n' + coverageText },
            { type: 'mrkdwn', text: '*Sub Plans:*\n' + (job.hasSubPlans || '') }
          ]
        },
        { type: 'section', text: { type: 'mrkdwn', text: '*Sub Plans Location:*\n' + (job.subPlansLocation || '') } },
        { type: 'context', elements: [{ type: 'mrkdwn', text: ptoPolicyLine }] }
      ]
    })

    qPush_(cfg.DEBUG_QUEUE_KEY, {
      ts: new Date().toISOString(),
      source: 'sendNotifyAndReceipt_',
      message: 'notify chat.postMessage',
      data: resp
    }, 250, 60 * 30)
  }

  // DM receipt
  if (job.userId) {
    const opened = slackApi_(cfg, 'conversations.open', { users: job.userId })
    const dmChannel = opened?.channel?.id || ''

    qPush_(cfg.DEBUG_QUEUE_KEY, {
      ts: new Date().toISOString(),
      source: 'sendNotifyAndReceipt_',
      message: 'dm conversations.open',
      data: opened
    }, 250, 60 * 30)

    if (dmChannel) {
      const dmResp = slackApi_(cfg, 'chat.postMessage', {
        channel: dmChannel,
        text: 'Same-Day Full Absence receipt',
        blocks: [
          { type: 'section', text: { type: 'mrkdwn', text: ':white_check_mark: *Same-Day Full Absence Submitted*' } },
          {
            type: 'section',
            fields: [
              { type: 'mrkdwn', text: '*Receipt ID:*\nRow ' + job.receiptRow },
              { type: 'mrkdwn', text: '*Submitted:*\n' + formatIsoLocal_(job.submittedAtIso, cfg.TZ) }
            ]
          },
          {
            type: 'section',
            fields: [
              { type: 'mrkdwn', text: '*Coverage Needed:*\n' + coverageText },
              { type: 'mrkdwn', text: '*Sub Plans:*\n' + (job.hasSubPlans || '') }
            ]
          },
          { type: 'section', text: { type: 'mrkdwn', text: '*Sub Plans Location:*\n' + (job.subPlansLocation || '') } },
          { type: 'context', elements: [{ type: 'mrkdwn', text: ptoPolicyLine }] }
        ]
      })

      qPush_(cfg.DEBUG_QUEUE_KEY, {
        ts: new Date().toISOString(),
        source: 'sendNotifyAndReceipt_',
        message: 'dm chat.postMessage',
        data: dmResp
      }, 250, 60 * 30)
    }
  }
}


/* ===========================
   VIEW BUILDER — one-step modal
   =========================== */

function buildFullAbsenceModal_(args) {
  const meta = {
    openedAtIso: args?.openedAtIso || new Date().toISOString(),
    userId: args?.userId || ''
  }

  const blocks = []

  blocks.push({ type: 'section', text: { type: 'mrkdwn', text: '*Same-Day Full Absence*' } })

  blocks.push({
    type: 'input',
    block_id: 'coverage_block',
    label: { type: 'plain_text', text: 'Coverage Needed:' },
    element: {
      type: 'multi_static_select',
      action_id: 'coverage',
      placeholder: { type: 'plain_text', text: 'Select coverage items…' },
      option_groups: coverageOptionGroups_()
    },
    optional: false
  })

  blocks.push({
    type: 'input',
    block_id: 'subplans_yn_block',
    label: { type: 'plain_text', text: 'Do you have sub plans?' },
    element: {
      type: 'radio_buttons',
      action_id: 'subplans_yn',
      options: [
        makeOption('Yes', 'YES'),
        makeOption('No', 'NO')
      ]
    },
    optional: false
  })

  blocks.push({
    type: 'input',
    block_id: 'subplans_location_block',
    label: { type: 'plain_text', text: 'Sub plans location (if none type "none" or "n/a")' },
    element: {
      type: 'plain_text_input',
      action_id: 'subplans_location',
      placeholder: { type: 'plain_text', text: 'Paste link or describe location…' }
    },
    optional: false
  })

  blocks.push({
    type: 'context',
    elements: [
      {
        type: 'mrkdwn',
        text: '*If submitted before 6:15am* this request will automatically be processed as PTO if PTO is available. *If after 6:15am* it will be considered a HR time off violation and not qualify for PTO.'
      }
    ]
  })

  return {
    type: 'modal',
    callback_id: 'same_day_full_absence',
    private_metadata: packMeta_(meta),
    title: { type: 'plain_text', text: 'Same-Day Full Absence' },
    close: { type: 'plain_text', text: 'Cancel' },
    submit: { type: 'plain_text', text: 'Submit' },
    blocks
  }
}
