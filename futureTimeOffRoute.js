// =====================
// VIEW SUBMISSION HANDLER
// =====================
function handleFutureViewSubmission_(cfg, payload) {
  
    const cb = payload?.view?.callback_id || ''

    if (cb === 'timeoff_step1') {
      const prevMeta = safeJsonParse_(payload?.view?.private_metadata) || {}
      const step1 = parseStep1State_(payload?.view?.state?.values || {})
      const errors = validateStep1_(step1)
      if (errors) return json_(200, { response_action: 'errors', errors })

      const apInfo = computeApEligibility_(cfg, step1.startDate, step1.isApRequest)

      const meta = {
        ...prevMeta,
        userId: payload?.user?.id || '',
        step1,
        apInfo
      }

      if (step1.isApRequest === 'yes' && apInfo.apEligible === true) {
        return json_(200, { response_action: 'update', view: buildApModal_({ cfg, privateMeta: meta }) })
      }

      return json_(200, { response_action: 'update', view: buildPtoModal_({ privateMeta: meta }) })
    }

    if (cb === 'timeoff_pto') {
      const meta = safeJsonParse_(payload?.view?.private_metadata) || {}
      const step1 = meta.step1 || {}
      const apInfo = meta.apInfo || {}
      const pto = parsePtoState_(payload?.view?.state?.values || {})

      const combined = {
        requestKind: 'PTO',
        slackUserId: payload?.user?.id || '',
        slackUserName: payload?.user?.username || payload?.user?.name || '',
        startDate: step1.startDate || '',
        startTime: step1.startTime || '',
        endDate: step1.endDate || '',
        endTime: step1.endTime || '',
        isApRequest: step1.isApRequest || 'no',
        apEligible: apInfo.apEligible ? 'yes' : 'no',
        apDaysUntilStart: String(apInfo.daysUntilStart ?? ''),
        offsiteAssignment: pto.offsiteAssignment || '',
        coverageNeeded: (pto.coverageNeeded || []).join(', '),
        hasSubPlans: pto.hasSubPlans || '',
        subPlansLink: pto.subPlansLink || '',
        notes: pto.notes || ''
      }

      const errors = validatePto_(combined, pto)
      if (errors) return json_(200, { response_action: 'errors', errors })

      enqueueFutureTimeOff_(cfg, payload, combined)
      return json_(200, { response_action: 'clear' })
    }

    if (cb === 'timeoff_ap') {
      const meta = safeJsonParse_(payload?.view?.private_metadata) || {}
      const step1 = meta.step1 || {}
      const apInfo = meta.apInfo || {}

      const combined = {
        requestKind: 'AP',
        slackUserId: payload?.user?.id || '',
        slackUserName: payload?.user?.username || payload?.user?.name || '',
        startDate: step1.startDate || '',
        startTime: step1.startTime || '',
        endDate: step1.endDate || '',
        endTime: step1.endTime || '',
        isApRequest: 'yes',
        apEligible: apInfo.apEligible ? 'yes' : 'no',
        apDaysUntilStart: String(apInfo.daysUntilStart ?? ''),
        apApprovalUrl: cfg.AP_APPROVAL_URL || ''
      }

      enqueueFutureTimeOff_(cfg, payload, combined)
      return json_(200, { response_action: 'clear' })
    }

    return json_(200, { response_action: 'clear' })

}



// =====================
// BACK BUTTON HANDLER
// =====================
function handleBackButton_(cfg, payload) {
  const meta = safeJsonParse_(payload.view?.private_metadata) || {}
  const fromCb = payload.view?.callback_id || ''

  // Preserve drafts from whichever modal we are on
  if (fromCb === 'timeoff_pto') meta.ptoDraft = parsePtoState_(payload.view?.state?.values || {})
  if (fromCb === 'timeoff_ap') meta.apDraft = parseApState_(payload.view?.state?.values || {})

  const step1View = buildTimeOffModalStep1_({ privateMeta: meta })

  slackApi_(cfg, 'views.update', {
    view_id: payload.view?.id,
    hash: payload.view?.hash,
    view: step1View
  })
}

// =====================
// OPTIONS
// =====================
function opt_(text, value) {
  return {
    text: { type: 'plain_text', text: text },
    value: String(value)
  }
}

function coverageOptionGroups_() {
  return [
    { label: { type: 'plain_text', text: 'N/A' }, options: [makeOption('No coverage needed', 'NA')] },
    {
      label: { type: 'plain_text', text: 'Classroom Coverage' },
      options: [
        makeOption('MS Period 1', 'MS_P1'),
        makeOption('MS Period 2', 'MS_P2'),
        makeOption('MS Period 3', 'MS_P3'),
        makeOption('MS Period 4', 'MS_P4'),
        makeOption('MS Period 5', 'MS_P5'),
        makeOption('HS Block 1', 'HS_B1'),
        makeOption('HS Block 2', 'HS_B2'),
        makeOption('HS Block 3', 'HS_B3'),
        makeOption('HS Block 4', 'HS_B4')
      ]
    },
    {
      label: { type: 'plain_text', text: 'Duty Coverage — AM' },
      options: [
        makeOption('AM Homeroom', 'DUTY_AM_HOMEROOM'),
        makeOption('AM Cafe Duty', 'DUTY_AM_CAFE'),
        makeOption('AM Hallway — Front', 'DUTY_AM_HALL_FRONT'),
        makeOption('AM Hallway — Back', 'DUTY_AM_HALL_BACK'),
        makeOption('AM Hallway — MS', 'DUTY_AM_HALL_MS'),
        makeOption('AM Hallway — HS', 'DUTY_AM_HALL_HS'),
        makeOption('AM Gym Door', 'DUTY_AM_GYM'),
        makeOption('AM Restroom — HS', 'DUTY_AM_RESTROOM_HS'),
        makeOption('AM Restroom — MS', 'DUTY_AM_RESTROOM_MS'),
        makeOption('AM Restroom — 8th Grade', 'DUTY_AM_RESTROOM_8TH'),
        makeOption('AM Bus Duty', 'DUTY_AM_BUS'),
        makeOption('AM Car Line / Walkers', 'DUTY_AM_CARLINE')
      ]
    },
    {
      label: { type: 'plain_text', text: 'Duty Coverage — Lunch' },
      options: [
        makeOption('Lunch — Cafe', 'DUTY_LUNCH_CAFE'),
        makeOption('Lunch — Recess', 'DUTY_LUNCH_RECESS'),
        makeOption('Lunch — Hall', 'DUTY_LUNCH_HALL')
      ]
    },
    {
      label: { type: 'plain_text', text: 'Duty Coverage — PM' },
      options: [
        makeOption('PM Hallway — Front', 'DUTY_PM_HALL_FRONT'),
        makeOption('PM Hallway — Back', 'DUTY_PM_HALL_BACK'),
        makeOption('PM Hallway — MS', 'DUTY_PM_HALL_MS'),
        makeOption('PM Hallway — HS', 'DUTY_PM_HALL_HS'),
        makeOption('PM Gym Door', 'DUTY_PM_GYM'),
        makeOption('PM Restroom — HS', 'DUTY_PM_RESTROOM_HS'),
        makeOption('PM Restroom — MS', 'DUTY_PM_RESTROOM_MS'),
        makeOption('PM Restroom — 8th Grade', 'DUTY_PM_RESTROOM_8TH'),
        makeOption('PM Bus Duty', 'DUTY_PM_BUS'),
        makeOption('PM Car Line / Walkers', 'DUTY_PM_CARLINE')
      ]
    }
  ]
}

function findOptionInGroups_(groups, value) {
  const v = String(value)
  for (let i = 0; i < groups.length; i++) {
    const opts = groups[i]?.options || []
    for (let j = 0; j < opts.length; j++) {
      if (String(opts[j].value) === v) return opts[j]
    }
  }
  return null
}

// =====================
// STEP 1 VIEW
// =====================
function buildTimeOffModalStep1_(opts) {
  const meta = opts?.privateMeta || {}
  const step1 = meta.step1 || {}


  const hours = buildHourOptions_(6, 18) // 06–18
  const mins = buildMinuteOptions_()  // 00–59

  return {
    type: 'modal',
    callback_id: 'timeoff_step1',
    title: { type: 'plain_text', text: 'Future Time Off Request' },
    submit: { type: 'plain_text', text: 'Next' },
    close: { type: 'plain_text', text: 'Cancel' },
    private_metadata: JSON.stringify(meta),
    blocks: [
      {
        type: 'input',
        block_id: 'start_date_blk',
        label: { type: 'plain_text', text: 'Start date' },
        element: {
          type: 'datepicker',
          action_id: 'start_date',
          placeholder: { type: 'plain_text', text: 'Select a date' },
          ...(step1.startDate ? { initial_date: step1.startDate } : {})
        }
      },
{
  type: 'input',
  block_id: 'start_hour_blk',
  label: { type: 'plain_text', text: 'Start hour' },
  element: {
    type: 'static_select',
    action_id: 'start_hour',
    placeholder: { type: 'plain_text', text: 'Hour' },
    options: hours,
    ...(step1.startHour ? { initial_option: findOption_(hours, step1.startHour) } : {})
  }
},
{
  type: 'input',
  block_id: 'start_min_blk',
  label: { type: 'plain_text', text: 'Start minute' },
  element: {
    type: 'static_select',
    action_id: 'start_min',
    placeholder: { type: 'plain_text', text: 'Minute' },
    options: mins,
    ...(step1.startMin ? { initial_option: findOption_(mins, step1.startMin) } : {})
  }
},

      {
        type: 'input',
        block_id: 'end_date_blk',
        label: { type: 'plain_text', text: 'End date' },
        element: {
          type: 'datepicker',
          action_id: 'end_date',
          placeholder: { type: 'plain_text', text: 'Select a date' },
          ...(step1.endDate ? { initial_date: step1.endDate } : {})
        }
      },

{
  type: 'input',
  block_id: 'end_hour_blk',
  label: { type: 'plain_text', text: 'End hour' },
  element: {
    type: 'static_select',
    action_id: 'end_hour',
    placeholder: { type: 'plain_text', text: 'Hour' },
    options: hours,
    ...(step1.endHour ? { initial_option: findOption_(hours, step1.endHour) } : {})
  }
},
{
  type: 'input',
  block_id: 'end_min_blk',
  label: { type: 'plain_text', text: 'End minute' },
  element: {
    type: 'static_select',
    action_id: 'end_min',
    placeholder: { type: 'plain_text', text: 'Minute' },
    options: mins,
    ...(step1.endMin ? { initial_option: findOption_(mins, step1.endMin) } : {})
  }
},

      { type: 'divider' },

      {
        type: 'input',
        block_id: 'ap_blk',
        label: { type: 'plain_text', text: 'Is this an AP request?' },
        element: {
          type: 'static_select',
          action_id: 'is_ap_request',
          placeholder: { type: 'plain_text', text: 'yes / no' },
          options: [opt_('Yes', 'yes'), opt_('No', 'no')],
          ...(step1.isApRequest
            ? { initial_option: step1.isApRequest === 'yes' ? opt_('Yes', 'yes') : opt_('No', 'no') }
            : {})
        }
      },

      {
        type: 'context',
        elements: [
          {
            type: 'mrkdwn',
            text: 'If you select *Yes*, your start date must be at least *5 days from today* to qualify for an AP request. Otherwise, it will be processed as PTO.'
          }
        ]
      }
    ]
  }
}

// =====================
// PTO MODAL
// =====================
function buildPtoModal_(opts) {
  const meta = opts?.privateMeta || {}
  const step1 = meta.step1 || {}
  const apInfo = meta.apInfo || {}
  const draft = meta.ptoDraft || {}

  const warning =
    (step1.isApRequest === 'yes' && apInfo.apEligible === false)
      ? [
          {
            type: 'section',
            text: {
              type: 'mrkdwn',
              text: '⚠️ *AP timing notice*\nYour start date is less than 5 days away, so this will be handled as a *PTO* request.'
            }
          },
          { type: 'divider' }
        ]
      : []

  const optByValue = (options, v) => options.find(o => o.value === String(v)) || null

  const offsiteOptions = [
    opt_('Yes', 'yes'),
    opt_('No', 'no')
  ]

  const subPlanOptions = [
    opt_('No, I do not have sub plans', 'no'),
    opt_('Yes, I have sub plans', 'yes'),
    opt_('N/A', 'na')
  ]

  return {
    type: 'modal',
    callback_id: 'timeoff_pto',
    title: { type: 'plain_text', text: 'Future Time Off Request' },
    submit: { type: 'plain_text', text: 'Submit' },
    close: { type: 'plain_text', text: 'Cancel' },
    private_metadata: JSON.stringify(meta),
    blocks: [
      ...warning,

      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text:
            `*Review:*\n• ${step1.startDate || ''} ${step1.startTime || ''} → ${step1.endDate || ''} ${step1.endTime || ''}\n` +
            `• AP request: *${step1.isApRequest || ''}*`
        }
      },

      {
        type: 'actions',
        block_id: 'nav_blk',
        elements: [
          {
            type: 'button',
            action_id: 'future_timeoff_back',
            text: { type: 'plain_text', text: 'Back' },
            value: 'back'
          }
        ]
      },

      { type: 'divider' },

      {
        type: 'input',
        block_id: 'offsite_blk',
        label: { type: 'plain_text', text: 'Is this an off-site assignment?' },
        element: {
          type: 'static_select',
          action_id: 'offsite',
          placeholder: { type: 'plain_text', text: 'Yes / No' },
          options: offsiteOptions,
          ...(draft.offsiteAssignment ? { initial_option: optByValue(offsiteOptions, draft.offsiteAssignment) } : {})
        }
      },

      {
        type: 'input',
        block_id: 'coverage_blk',
        label: { type: 'plain_text', text: 'Coverage needed' },
        element: {
          type: 'multi_static_select',
          action_id: 'coverage_needed',
          placeholder: { type: 'plain_text', text: 'Select coverage option(s)' },
          option_groups: coverageOptionGroups_(),

          ...(draft.coverageNeeded && Array.isArray(draft.coverageNeeded) && draft.coverageNeeded.length
            ? {
                initial_options: draft.coverageNeeded
                  .map(v => findOptionInGroups_(coverageOptionGroups_(), v))
                  .filter(Boolean)
              }
            : {})
        }
      },

      {
        type: 'input',
        block_id: 'subplans_blk',
        label: { type: 'plain_text', text: 'Do you have sub plans?' },
        element: {
          type: 'static_select',
          action_id: 'has_subplans',
          placeholder: { type: 'plain_text', text: 'Pick an option' },
          options: subPlanOptions,
          ...(draft.hasSubPlans ? { initial_option: optByValue(subPlanOptions, draft.hasSubPlans) } : {})
        }
      },

      {
        type: 'section',
        text: { type: 'mrkdwn', text: '*Sub plans link*\nPaste the link below. If non-instructional or not applicable, write `N/A`.' }
      },

      {
        type: 'input',
        block_id: 'subplans_link_blk',
        label: { type: 'plain_text', text: 'Sub plans link or location' },
        element: {
          type: 'plain_text_input',
          action_id: 'subplans_link',
          placeholder: { type: 'plain_text', text: 'Paste sub plans link (Drive / Doc) or N/A' },
          ...(draft.subPlansLink ? { initial_value: draft.subPlansLink } : {})
        }
      },

      {
        type: 'input',
        optional: true,
        block_id: 'notes_blk',
        label: { type: 'plain_text', text: 'Notes (optional)' },
        element: {
          type: 'plain_text_input',
          action_id: 'notes',
          placeholder: { type: 'plain_text', text: 'Anything else admin should know?' },
          multiline: true,
          ...(draft.notes ? { initial_value: draft.notes } : {})
        }
      }
    ]
  }
}

// =====================
// AP MODAL (your current stub)
// =====================
function buildApModal_(opts) {
   const cfg = opts?.cfg || getConfig_()
  const meta = opts?.privateMeta || {}
  const step1 = meta.step1 || {}
  const url = cfg.AP_APPROVAL_URL || '' 


  return {
    type: 'modal',
    callback_id: 'timeoff_ap',
    title: { type: 'plain_text', text: 'AP Request' },
    submit: { type: 'plain_text', text: 'I Understand' },
    close: { type: 'plain_text', text: 'Cancel' },
    private_metadata: JSON.stringify(meta),
    blocks: [
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text:
            `*AP Request qualifies*\n` +
            `• ${step1.startDate || ''} ${step1.startTime || ''} → ${step1.endDate || ''} ${step1.endTime || ''}\n\n` +
            `*Next step:*\nClick the link and fill out the paperwork for AP request approval.\n\n` +
            (url ? `👉 <${url}|Open AP approval paperwork>` : '⚠️ Missing AP approval link')
        }
      },
      { type: 'divider' },
      {
        type: 'actions',
        block_id: 'nav_blk',
        elements: [
          {
            type: 'button',
            action_id: 'future_timeoff_back',
            text: { type: 'plain_text', text: 'Back' },
            value: 'back'
          }
        ]
      }
    ]
  }
}


// =====================
// PARSERS + VALIDATION
// =====================
function parseStep1State_(values) {
  const pick = (blockId, actionId) => values?.[blockId]?.[actionId] || null

  const startHour = pick('start_hour_blk', 'start_hour')?.selected_option?.value || ''
  const startMin = pick('start_min_blk', 'start_min')?.selected_option?.value || ''

  const endHour = pick('end_hour_blk', 'end_hour')?.selected_option?.value || ''
  const endMin = pick('end_min_blk', 'end_min')?.selected_option?.value || ''

  return {
    startDate: pick('start_date_blk', 'start_date')?.selected_date || '',
    endDate: pick('end_date_blk', 'end_date')?.selected_date || '',

    // store pieces so we can repopulate dropdowns on Back
    startHour,
    startMin,
    endHour,
    endMin,

    // store canonical 24h strings for downstream logic/sheet
    startTime: startHour, startMin,
    endTime: endHour, endMin,

    isApRequest: pick('ap_blk', 'is_ap_request')?.selected_option?.value || ''
  }
}


function parsePtoState_(values) {
  const pick = (blockId, actionId) => values?.[blockId]?.[actionId] || null

  const coverageSelected = pick('coverage_blk', 'coverage_needed')?.selected_options || []
  const coverageValues = coverageSelected.map(o => o.value)

  return {
    offsiteAssignment: pick('offsite_blk', 'offsite')?.selected_option?.value || '',
    coverageNeeded: coverageValues,
    hasSubPlans: pick('subplans_blk', 'has_subplans')?.selected_option?.value || '',
    subPlansLink: pick('subplans_link_blk', 'subplans_link')?.value || '',
    notes: pick('notes_blk', 'notes')?.value || ''
  }
}



function validateStep1_(step1) {
  const errors = {}

  if (!step1.startDate) errors['start_date_blk'] = 'Please choose a start date'
  if (!step1.endDate) errors['end_date_blk'] = 'Please choose an end date'

  if (!step1.startHour) errors['start_hour_blk'] = 'Pick an hour'
  if (!step1.startMin) errors['start_min_blk'] = 'Pick minutes'


  if (!step1.endHour) errors['end_hour_blk'] = 'Pick an hour'
  if (!step1.endMin) errors['end_min_blk'] = 'Pick minutes'


  if (!step1.startTime) errors['start_hour_blk'] = 'Invalid start time'
  if (!step1.endTime) errors['end_hour_blk'] = 'Invalid end time'

  if (!step1.isApRequest) errors['ap_blk'] = 'Please select yes or no'

  return Object.keys(errors).length ? errors : null
}


function validatePto_(combined, ptoRaw) {
  const errors = {}

  if (!combined.offsiteAssignment) errors['offsite_blk'] = 'Please select Yes or No'

  const covArr = ptoRaw?.coverageNeeded || []
  if (!covArr || !covArr.length) errors['coverage_blk'] = 'Please select at least one coverage option'

  if (!combined.hasSubPlans) errors['subplans_blk'] = 'Please answer sub plans'
  if (!combined.subPlansLink) errors['subplans_link_blk'] = 'Please provide a link or N/A'

  return Object.keys(errors).length ? errors : null
}


function computeApEligibility_(cfg, startDateStr, isApRequest) {
  if (isApRequest !== 'yes') return { apEligible: true, daysUntilStart: null }
  if (!startDateStr) return { apEligible: false, daysUntilStart: null }

  // Force "today" to be America/Chicago calendar date (not server clock)
  const todayStr = Utilities.formatDate(new Date(), cfg.TZ, 'yyyy-MM-dd')

  const todayMs = ymdToUtcMs_(todayStr)
  const startMs = ymdToUtcMs_(startDateStr)

  const msPerDay = 24 * 60 * 60 * 1000
  const diffDays = Math.round((startMs - todayMs) / msPerDay)

  return {
    apEligible: diffDays >= 5,
    daysUntilStart: diffDays
  }
}

function ymdToUtcMs_(ymd) {
  const parts = String(ymd || '').split('-').map(n => Number(n))
  const y = parts[0]
  const m = parts[1]
  const d = parts[2]
  return Date.UTC(y, m - 1, d)
}


function safeJsonParse_(s) {
  try {
    return JSON.parse(s || '')
  } catch (e) {
    return null
  }
}

// =====================
// QUEUE (no Sheets writes in doPost)
// =====================
function enqueueFutureTimeOff_(cfg, payload, combined) {
  const props = PropertiesService.getScriptProperties()

  const item = {
    receivedAt: new Date().toISOString(),
    user: combined.slackUserId || payload?.user?.id || '',
    userName: combined.slackUserName || payload?.user?.username || payload?.user?.name || '',
    data: combined
  }

  const raw = props.getProperty(cfg.FUTURE_QUEUE_PROP_KEY)
  const arr = raw ? JSON.parse(raw) : []
  arr.push(item)
  props.setProperty(cfg.FUTURE_QUEUE_PROP_KEY, JSON.stringify(arr))
}


function flushFutureTimeOffQueueCentral_() {
  if (isSlackHot_()) return

  const cfg = getConfig_()
  const lock = LockService.getScriptLock()
  if (!lock.tryLock(1000)) return

  let items = []
  try {
    const props = PropertiesService.getScriptProperties()

    const raw = props.getProperty(cfg.FUTURE_QUEUE_PROP_KEY)
    items = raw ? JSON.parse(raw) : []
    if (!items.length) return

    // Clear queue first (so if we error mid-way, we can requeue cleanly)
    props.deleteProperty(cfg.FUTURE_QUEUE_PROP_KEY)

    const ss = SpreadsheetApp.openByUrl(cfg.SHEET_URL)
    const sh = ss.getSheetByName(cfg.FUTURE_SHEET_NAME) || ss.insertSheet(cfg.FUTURE_SHEET_NAME)
    const dbg = ss.getSheetByName(cfg.DEBUG_SHEET) || ss.insertSheet(cfg.DEBUG_SHEET)

    // headers
    if (sh.getLastRow() === 0) {
      sh.appendRow([
        'Timestamp ISO',
        'Request Kind',
        'Slack User ID',
        'Slack Username',
        'Slack Email',
        'Start Date',
        'Start Time',
        'End Date',
        'End Time',
        'Is AP Request',
        'AP Eligible (>=5 days)',
        'Days Until Start',
        'Off-site Assignment',
        'Coverage Needed',
        'Has Sub Plans',
        'Sub Plans Link',
        'Notes',
        'AP Approval URL'
      ])
    }

    const nowIso = () => new Date().toISOString()

    // cache emails within this flush
    const emailCache = {}

    const getSlackEmailSafe_ = (userId) => {
      const uid = String(userId || '').trim()
      if (!uid) return ''
      if (emailCache[uid] !== undefined) return emailCache[uid]

      try {
        const res = UrlFetchApp.fetch(
          'https://slack.com/api/users.info?user=' + encodeURIComponent(uid),
          {
            method: 'get',
            headers: { Authorization: 'Bearer ' + cfg.SLACK_BOT_TOKEN },
            muteHttpExceptions: true
          }
        )
        const data = JSON.parse(res.getContentText() || '{}')
        const email = (data && data.ok) ? (data.user?.profile?.email || '') : ''
        emailCache[uid] = email
        return email
      } catch (e) {
        emailCache[uid] = ''
        dbg.appendRow([new Date(), 'future_email_lookup', 'EXCEPTION', uid, String(e && e.stack || e)])
        return ''
      }
    }

    const formatAlertTextLocal_ = (it, d, rowNumber) => {
      const who = it.userName ? `@${it.userName}` : (it.user || 'Unknown user')
      const whenBits = [
        d.startDate ? `${d.startDate}${d.startTime ? ` ${d.startTime}` : ''}` : '',
        d.endDate ? `${d.endDate}${d.endTime ? ` ${d.endTime}` : ''}` : ''
      ].filter(Boolean)

      const lines = [
        d.requestKind === 'AP' ? '🚨 *New AP Request*' : '🚨 *New PTO Request*',
        `*From:* ${who}`,
        whenBits.length ? `*When:* ${whenBits.join(' → ')}` : null,
        `*AP request:* ${d.isApRequest || 'no'} (eligible: ${d.apEligible || 'n/a'})`,
        d.requestKind === 'PTO' ? `*Off-site assignment:* ${d.offsiteAssignment || ''}` : null,
        d.requestKind === 'PTO' ? (d.coverageNeeded ? `*Coverage:* ${d.coverageNeeded}` : null) : null,
        d.requestKind === 'PTO' ? (d.hasSubPlans ? `*Sub plans:* ${d.hasSubPlans}${d.subPlansLink ? ` (${d.subPlansLink})` : ''}` : null) : null,
        d.requestKind === 'AP' ? (d.apApprovalUrl ? `*AP approval:* ${d.apApprovalUrl}` : null) : null,
        d.notes ? `*Notes:* ${d.notes}` : null,
        `*Sheet row:* ${rowNumber}`
      ].filter(Boolean)

      return lines.join('\n')
    }

    const buildDmTextLocal_ = (it, d, rowNumber) => {
      const who = it.userName ? `@${it.userName}` : (it.user || 'Unknown')
      const whenBits = [
        d.startDate ? `${d.startDate}${d.startTime ? ` ${d.startTime}` : ''}` : '',
        d.endDate ? `${d.endDate}${d.endTime ? ` ${d.endTime}` : ''}` : ''
      ].filter(Boolean)

      const lines = [
        d.requestKind === 'AP' ? '✅ *AP Request Received*' : '✅ *PTO Request Received*',
        `*From:* ${who}`,
        whenBits.length ? `*When:* ${whenBits.join(' → ')}` : null,
        `*AP request:* ${d.isApRequest || 'no'} (eligible: ${d.apEligible || 'n/a'})`,
        d.requestKind === 'PTO' ? `*Off-site assignment:* ${d.offsiteAssignment || ''}` : null,
        d.requestKind === 'PTO' ? (d.coverageNeeded ? `*Coverage:* ${d.coverageNeeded}` : null) : null,
        d.requestKind === 'PTO' ? (d.hasSubPlans ? `*Sub plans:* ${d.hasSubPlans}${d.subPlansLink ? ` (${d.subPlansLink})` : ''}` : null) : null,
        d.requestKind === 'AP' ? (d.apApprovalUrl ? `*AP approval:* ${d.apApprovalUrl}` : null) : null,
        d.notes ? `*Notes:* ${d.notes}` : null,
        `*Receipt ID:* Row ${rowNumber}`
      ].filter(Boolean)

      return lines.join('\n')
    }

    const buildEmailSubjectLocal_ = (it, d) => {
      const who = it.userName ? `@${it.userName}` : (it.user || 'Unknown')
      const start = d.startDate || ''
      const end = d.endDate || ''
      const span = (start || end) ? ` ${start}${end ? ` → ${end}` : ''}` : ''
      return `Time Off Receipt: ${d.requestKind || 'Request'} ${who}${span}`
    }

    const buildEmailTextLocal_ = (it, d, userEmail, rowNumber) => {
      const lines = [
        'Time Off Request Receipt',
        '',
        `Receipt ID: Row ${rowNumber}`,
        `Request Kind: ${d.requestKind || ''}`,
        `Slack User: ${it.userName ? '@' + it.userName : (it.user || '')}`,
        `Email: ${userEmail || ''}`,
        '',
        `Start: ${(d.startDate || '')} ${(d.startTime || '')}`.trim(),
        `End: ${(d.endDate || '')} ${(d.endTime || '')}`.trim(),
        `AP request: ${d.isApRequest || ''} (eligible: ${d.apEligible || ''})`,
        d.apDaysUntilStart !== undefined ? `Days until start: ${d.apDaysUntilStart || ''}` : null,
        '',
        d.requestKind === 'PTO' ? `Off-site assignment: ${d.offsiteAssignment || ''}` : null,
        d.requestKind === 'PTO' ? `Coverage needed: ${d.coverageNeeded || ''}` : null,
        d.requestKind === 'PTO' ? `Has sub plans: ${d.hasSubPlans || ''}` : null,
        d.requestKind === 'PTO' ? `Sub plans link: ${d.subPlansLink || ''}` : null,
        d.notes ? `Notes: ${d.notes}` : null,
        d.requestKind === 'AP' && d.apApprovalUrl ? '' : null,
        d.requestKind === 'AP' && d.apApprovalUrl ? `AP approval link: ${d.apApprovalUrl}` : null
      ].filter(v => v !== null)

      return lines.join('\n')
    }

    const safeSendEmailLocal_ = (to, subject, textBody) => {
      try {
        const email = String(to || '').trim()
        if (!email) return
        const opts = {}
        if (cfg.EMAIL_ALIAS) opts.from = cfg.EMAIL_ALIAS
        MailApp.sendEmail(email, subject, textBody, opts)
      } catch (e) {
        dbg.appendRow([new Date(), 'future_email', 'FAIL', to, subject, String(e && e.stack || e)])
      }
    }

    // Build rows
    const startRow = sh.getLastRow() + 1
    const rows = items.map(it => {
      const d = it.data || {}
      const userId = String(it.user || '').trim()
      const email = getSlackEmailSafe_(userId)

      return [
        nowIso(),
        d.requestKind || '',
        userId,
        it.userName || '',
        email || '',
        d.startDate || '',
        d.startTime || '',
        d.endDate || '',
        d.endTime || '',
        d.isApRequest || '',
        d.apEligible || '',
        d.apDaysUntilStart || '',
        d.offsiteAssignment || '',
        d.coverageNeeded || '',
        d.hasSubPlans || '',
        d.subPlansLink || '',
        d.notes || '',
        d.apApprovalUrl || ''
      ]
    })

    sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows)

    // Notify/DM/email after sheet write so we have receipt row numbers
    items.forEach((it, idx) => {
      const d = it.data || {}
      const rowNumber = startRow + idx
      const userId = String(it.user || '').trim()
      const email = emailCache[userId] || ''

      // Notify channel
      try {
        const ch = cfg.FUTURE_NOTIFY_CHANNEL
        if (ch) {
          const text = formatAlertTextLocal_(it, d, rowNumber)
          const res = slackApi_(cfg, 'chat.postMessage', { channel: ch, text })
          dbg.appendRow([new Date(), 'future_notify', `ok=${res?.ok} err=${res?.error || ''}`, ch, rowNumber])
        }
      } catch (e) {
        dbg.appendRow([new Date(), 'future_notify', 'EXCEPTION', rowNumber, String(e && e.stack || e)])
      }

      // DM receipt
      try {
        const dmText = buildDmTextLocal_(it, d, rowNumber)
        dmUser_(cfg, userId, dmText)
        dbg.appendRow([new Date(), 'future_dm', 'SENT', userId, rowNumber])
      } catch (e) {
        dbg.appendRow([new Date(), 'future_dm', 'EXCEPTION', userId, rowNumber, String(e && e.stack || e)])
      }

      // Email receipt
      try {
        if (email) {
          const subject = buildEmailSubjectLocal_(it, d)
          const body = buildEmailTextLocal_(it, d, email, rowNumber)
          safeSendEmailLocal_(email, subject, body)

          if (cfg.ADMIN_EMAILS && Array.isArray(cfg.ADMIN_EMAILS) && cfg.ADMIN_EMAILS.length) {
            safeSendEmailLocal_(cfg.ADMIN_EMAILS.join(','), '[ADMIN COPY] ' + subject, body)
          }

          dbg.appendRow([new Date(), 'future_email', 'SENT', email, rowNumber])
        } else {
          dbg.appendRow([new Date(), 'future_email', 'SKIP no email', userId, rowNumber])
        }
      } catch (e) {
        dbg.appendRow([new Date(), 'future_email', 'EXCEPTION', userId, rowNumber, String(e && e.stack || e)])
      }
    })

    debugLog_(cfg, 'flushFutureTimeOffQueueCentral_', `Processed ${items.length} item(s)`)
  } catch (err) {
    // If flush fails, requeue everything so it retries next minute
    try {
      const cfg = getConfig_()
      const props = PropertiesService.getScriptProperties()
      const rawExisting = props.getProperty(cfg.FUTURE_QUEUE_PROP_KEY)
      const existing = rawExisting ? JSON.parse(rawExisting) : []
      // We can't access the local `items` here safely if parsing failed early,
      // so we only log the failure.
      debugLog_(cfg, 'flushFutureTimeOffQueueCentral__ERROR', String(err && err.stack || err))
      props.setProperty(cfg.FUTURE_QUEUE_PROP_KEY, JSON.stringify(existing))
    } catch (e) {}
  } finally {
    lock.releaseLock()
  }
}







