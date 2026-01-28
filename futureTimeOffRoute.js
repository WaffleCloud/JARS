// =====================
// VIEW SUBMISSION HANDLER
// =====================
function handleFutureViewSubmission_(cfg, payload) {
  const cb = payload.view?.callback_id || ''

  // Back button is a block_action, not view_submission
  // But if you ever handle it here, keep this stub:
  // if (cb === 'future_timeoff_back') ...

  if (cb === 'timeoff_step1') {
    const prevMeta = safeJsonParse_(payload.view?.private_metadata) || {}
    const step1 = parseStep1State_(payload.view?.state?.values || {})
    const errors = validateStep1_(step1)
    if (errors) return json_(200, { response_action: 'errors', errors })

    const apInfo = computeApEligibility_(cfg, step1.startDate, step1.isApRequest)

    const meta = {
      ...prevMeta,
      userId: payload.user?.id || '',
      step1,
      apInfo
    }

    if (step1.isApRequest === 'yes' && apInfo.apEligible === true) {
      return json_(200, { response_action: 'update', view: buildApModal_({ cfg, privateMeta: meta }) })
    }

    return json_(200, { response_action: 'update', view: buildPtoModal_({ privateMeta: meta }) })
  }

  if (cb === 'timeoff_pto') {
    const meta = safeJsonParse_(payload.view?.private_metadata) || {}
    const step1 = meta.step1 || {}
    const apInfo = meta.apInfo || {}
    const pto = parsePtoState_(payload.view?.state?.values || {})

    const combined = {
      requestKind: 'PTO',
      slackUserId: payload.user?.id || '',
      slackUserName: payload.user?.username || payload.user?.name || '',

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
    const meta = safeJsonParse_(payload.view?.private_metadata) || {}
    const step1 = meta.step1 || {}
    const apInfo = meta.apInfo || {}

    const combined = {
      requestKind: 'AP',
      slackUserId: payload.user?.id || '',
      slackUserName: payload.user?.username || payload.user?.name || '',

      startDate: step1.startDate || '',
      startTime: step1.startTime || '',
      endDate: step1.endDate || '',
      endTime: step1.endTime || '',

      isApRequest: 'yes',
      apEligible: apInfo.apEligible ? 'yes' : 'no',
      apDaysUntilStart: String(apInfo.daysUntilStart ?? ''),

      apApprovalUrl: cfg.AP_APPROVAL_URL || '' // only if you add this to getConfig_
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

function makeOption(text, value) {
  return {
    text: { type: 'plain_text', text },
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
    options: hourOptions_(),
    ...(step1.startHour ? { initial_option: findOption_(hourOptions_(), step1.startHour) } : {})
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
    options: minuteOptions_(),
    ...(step1.startMin ? { initial_option: findOption_(minuteOptions_(), step1.startMin) } : {})
  }
},
{
  type: 'input',
  block_id: 'start_ampm_blk',
  label: { type: 'plain_text', text: 'Start AM/PM' },
  element: {
    type: 'static_select',
    action_id: 'start_ampm',
    placeholder: { type: 'plain_text', text: 'AM / PM' },
    options: ampmOptions_(),
    ...(step1.startAmPm ? { initial_option: findOption_(ampmOptions_(), step1.startAmPm) } : {})
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
    options: hourOptions_(),
    ...(step1.endHour ? { initial_option: findOption_(hourOptions_(), step1.endHour) } : {})
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
    options: minuteOptions_(),
    ...(step1.endMin ? { initial_option: findOption_(minuteOptions_(), step1.endMin) } : {})
  }
},
{
  type: 'input',
  block_id: 'end_ampm_blk',
  label: { type: 'plain_text', text: 'End AM/PM' },
  element: {
    type: 'static_select',
    action_id: 'end_ampm',
    placeholder: { type: 'plain_text', text: 'AM / PM' },
    options: ampmOptions_(),
    ...(step1.endAmPm ? { initial_option: findOption_(ampmOptions_(), step1.endAmPm) } : {})
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
  const startAmPm = pick('start_ampm_blk', 'start_ampm')?.selected_option?.value || ''

  const endHour = pick('end_hour_blk', 'end_hour')?.selected_option?.value || ''
  const endMin = pick('end_min_blk', 'end_min')?.selected_option?.value || ''
  const endAmPm = pick('end_ampm_blk', 'end_ampm')?.selected_option?.value || ''

  return {
    startDate: pick('start_date_blk', 'start_date')?.selected_date || '',
    endDate: pick('end_date_blk', 'end_date')?.selected_date || '',

    // store pieces so we can repopulate dropdowns on Back
    startHour,
    startMin,
    startAmPm,
    endHour,
    endMin,
    endAmPm,

    // store canonical 24h strings for downstream logic/sheet
    startTime: to24h_(startHour, startMin, startAmPm),
    endTime: to24h_(endHour, endMin, endAmPm),

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

function parseApState_(values) {
  const sections = []

  for (let i = 1; i <= 6; i++) {
    const itemBlock = `ap_cov_item_${i}_blk`
    const itemAction = `ap_cov_item_${i}`
    const planBlock = `ap_cov_plan_${i}_blk`
    const planAction = `ap_cov_plan_${i}`

    const selected = values?.[itemBlock]?.[itemAction]?.selected_options || []
    const coverage = selected.map(o => o.value)
    const plan = values?.[planBlock]?.[planAction]?.value || ''

    sections.push({ index: i, coverage, plan })
  }

  return { sections }
}

function validateStep1_(step1) {
  const errors = {}

  if (!step1.startDate) errors['start_date_blk'] = 'Please choose a start date'
  if (!step1.endDate) errors['end_date_blk'] = 'Please choose an end date'

  if (!step1.startHour) errors['start_hour_blk'] = 'Pick an hour'
  if (!step1.startMin) errors['start_min_blk'] = 'Pick minutes'
  if (!step1.startAmPm) errors['start_ampm_blk'] = 'Pick AM or PM'

  if (!step1.endHour) errors['end_hour_blk'] = 'Pick an hour'
  if (!step1.endMin) errors['end_min_blk'] = 'Pick minutes'
  if (!step1.endAmPm) errors['end_ampm_blk'] = 'Pick AM or PM'

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

function validateAp_(apRaw) {
  const errors = {}
  const sections = apRaw?.sections || []
  const s1 = sections[0] || { coverage: [], plan: '' }

  if (!s1.coverage || !s1.coverage.length) errors['ap_cov_item_1_blk'] = 'Please select at least one coverage item'
  if (!String(s1.plan || '').trim()) errors['ap_cov_plan_1_blk'] = 'Please provide a coverage plan'

  for (let i = 2; i <= 6; i++) {
    const s = sections[i - 1] || { coverage: [], plan: '' }
    const hasAny = (s.coverage && s.coverage.length) || String(s.plan || '').trim()

    if (hasAny) {
      if (!s.coverage || !s.coverage.length) errors[`ap_cov_item_${i}_blk`] = 'Select at least one coverage item (or clear this section)'
      if (!String(s.plan || '').trim()) errors[`ap_cov_plan_${i}_blk`] = 'Provide a plan (or clear this section)'
    }
  }

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
  const item = {
    receivedAt: new Date().toISOString(),
    user: combined.slackUserId || payload.user?.id || '',
    userName: combined.slackUserName || payload.user?.username || payload.user?.name || '',
    data: combined
  }

  const props = PropertiesService.getScriptProperties()
  const raw = props.getProperty(cfg.FUTURE_QUEUE_PROP_KEY)
  const arr = raw ? JSON.parse(raw) : []
  arr.push(item)
  props.setProperty(cfg.FUTURE_QUEUE_PROP_KEY, JSON.stringify(arr))

  ensureFutureFlushTrigger_(cfg)
}

function ensureFutureFlushTrigger_(cfg) {
  const props = PropertiesService.getScriptProperties()
  const alreadySet = props.getProperty(cfg.FUTURE_FLUSH_TRIGGER_PROP_KEY)
  if (alreadySet === '1') return

  // delete any old triggers for this worker
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'flushFutureTimeOffQueue_') {
      ScriptApp.deleteTrigger(t)
    }
  })

  ScriptApp.newTrigger('flushFutureTimeOffQueue_')
    .timeBased()
    .after((cfg.FUTURE_FLUSH_TRIGGER_DELAY_SECONDS || 30) * 1000)
    .create()

  props.setProperty(cfg.FUTURE_FLUSH_TRIGGER_PROP_KEY, '1')
}



function parseSlackPayload_(e) {
  const raw = e?.parameter?.payload
  if (!raw) return null
  return JSON.parse(raw)
}

function getSlackEmailByUserId_(cfg, userId) {
  if (!userId) return ''

  const res = UrlFetchApp.fetch('https://slack.com/api/users.info?user=' + encodeURIComponent(userId), {
    method: 'get',
    headers: { Authorization: `Bearer ${cfg.SLACK_BOT_TOKEN}` },
    muteHttpExceptions: true
  })

  const data = JSON.parse(res.getContentText() || '{}')
  if (!data.ok) return ''

  return data.user?.profile?.email || ''
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name)
}

function getOrCreateDebugSheet_(cfg) {
  const ss = SpreadsheetApp.openById(cfg.SHEET_ID)
  return getOrCreateSheet_(ss, cfg.DEBUG_SHEET_NAME)
}

function getDebugSheet_(cfg) {
  const ss = SpreadsheetApp.openById(cfg.SHEET_ID)
  return getOrCreateSheet_(ss, cfg.DEBUG_SHEET_NAME)
}

// =====================
// Slack channel alert (dynamic channel)
// =====================
function sendSlackChannelAlert_(cfg, channelId, text) {
  const dbg = getOrCreateDebugSheet_(cfg)

  if (!channelId) {
    dbg.appendRow([new Date(), 'sendSlackChannelAlert_', 'ABORT missing NOTIFY_CHANNEL_ID'])
    return
  }

  const res = slackApi_(cfg, 'chat.postMessage', {
    channel: channelId,
    text
  })

  if (!res || !res.ok) {
    dbg.appendRow([new Date(), 'sendSlackChannelAlert_', 'FAIL', channelId, JSON.stringify(res || {})])
    return
  }

  dbg.appendRow([new Date(), 'sendSlackChannelAlert_', 'SENT', channelId, res.ts || ''])
}

function formatAlertText_(it, d, sheetRowNumber) {
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
    d.requestKind === 'AP' ? summarizeApCoverage_(d.apCoverageJson) : null,
    d.notes ? `*Notes:* ${d.notes}` : null,
    `*Sheet row:* ${sheetRowNumber}`
  ].filter(Boolean)

  return lines.join('\n')
}

function summarizeApCoverage_(apCoverageJson) {
  try {
    const sections = JSON.parse(apCoverageJson || '[]')
    const filled = sections
      .filter(s => (s.coverage && s.coverage.length) || String(s.plan || '').trim())
      .map(s => `• Set ${s.index}: ${((s.coverage || []).join(', ') || '—')} — ${String(s.plan || '').trim()}`)
    if (!filled.length) return '*Coverage:* (none provided)'
    return ['*Coverage:*', ...filled].join('\n')
  } catch (e) {
    return '*Coverage:* (unable to parse)'
  }
}

// =====================
// DM RECEIPT
// =====================
function sendReceiptDm_(it, d, sheetRowNumber) {
  const dbg = getOrCreateDebugSheet_(cfg)

  const userId = String(it.user || '').trim()
  if (!userId) {
    dbg.appendRow([new Date(), 'sendReceiptDm_', 'ABORT missing userId'])
    return
  }

  const text = buildReceiptDmText_(it, d, sheetRowNumber)

  const res = slackApi_(cfg, 'chat.postMessage', {
    channel: userId,
    text
  })

  if (!res || !res.ok) {
    dbg.appendRow([new Date(), 'sendReceiptDm_', 'FAIL', userId, JSON.stringify(res || {})])
    return
  }

  dbg.appendRow([new Date(), 'sendReceiptDm_', 'SENT', userId, res.ts || ''])
}

function buildReceiptDmText_(it, d, sheetRowNumber) {
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
    d.requestKind === 'AP' ? summarizeApCoverage_(d.apCoverageJson) : null,
    d.notes ? `*Notes:* ${d.notes}` : null,
    sheetRowNumber ? `*Receipt ID:* Row ${sheetRowNumber}` : null
  ].filter(Boolean)

  return lines.join('\n')
}

// =====================
// EMAIL RECEIPTS
// =====================
function sendSubmitReceiptEmails_(it, d, userEmailOverride) {
  const dbg = getDebugSheet_()

  const userId = String(it.user || '').trim()
  if (!userId) {
    dbg.appendRow([new Date(), 'sendSubmitReceiptEmails_', 'ABORT missing userId'])
    return
  }

  const userEmail = String(userEmailOverride || getSlackEmailByUserId_(cfg, userId) || '').trim()
  if (!userEmail) {
    dbg.appendRow([new Date(), 'sendSubmitReceiptEmails_', `ABORT no email for ${userId}`])
    return
  }

  const subject = buildReceiptSubject_(it, d)
  const htmlBody = buildReceiptHtml_(it, d, userEmail)
  const textBody = buildReceiptText_(it, d, userEmail)

  safeSendEmail_({ to: userEmail, subject, htmlBody, textBody }, dbg)

  if (cfg.ADMIN_EMAILS && cfg.ADMIN_EMAILS.length) {
    safeSendEmail_({
      to: cfg.ADMIN_EMAILS.join(','),
      subject: '[ADMIN COPY] ' + subject,
      htmlBody,
      textBody
    }, dbg)
  }

  dbg.appendRow([new Date(), 'sendSubmitReceiptEmails_', 'SENT', userEmail, (cfg.ADMIN_EMAILS || []).join(',')])
}

function buildReceiptSubject_(it, d) {
  const prefix = cfg.EMAIL_SUBJECT_PREFIX || 'Time Off Receipt'
  const who = it.userName ? `@${it.userName}` : (it.user || 'Unknown')
  const start = d.startDate || ''
  const end = d.endDate || ''
  const span = (start || end) ? ` ${start}${end ? ` → ${end}` : ''}` : ''
  return `${prefix}: ${d.requestKind || 'Request'} ${who}${span}`
}

function buildReceiptHtml_(it, d, userEmail) {
  const baseRows = [
    ['Request Kind', d.requestKind],
    ['Slack User', it.userName ? `@${it.userName}` : (it.user || '')],
    ['Email', userEmail],
    ['Start', `${d.startDate || ''} ${d.startTime || ''}`.trim()],
    ['End', `${d.endDate || ''} ${d.endTime || ''}`.trim()],
    ['Is AP Request', d.isApRequest],
    ['AP Eligible (>=5 days)', d.apEligible],
    ['Days Until Start', d.apDaysUntilStart]
  ]

  const ptoRows = [
    ['Off-site Assignment', d.offsiteAssignment],
    ['Coverage Needed', d.coverageNeeded],
    ['Has Sub Plans', d.hasSubPlans],
    ['Sub Plans Link', d.subPlansLink],
    ['Notes', d.notes]
  ]

  const apRows = [
    ['AP Coverage', summarizeApCoveragePlain_(d.apCoverageJson)]
  ]

  const rows = (d.requestKind === 'AP')
    ? baseRows.concat(apRows)
    : baseRows.concat(ptoRows)

  const filtered = rows.filter(r => String(r[1] || '').trim() !== '')

  const tr = filtered.map(r => {
    return '<tr>' +
      '<td style="border:1px solid #ddd;padding:8px;width:35%"><b>' + escapeHtml_(r[0]) + '</b></td>' +
      '<td style="border:1px solid #ddd;padding:8px">' + escapeHtml_(String(r[1] || '')) + '</td>' +
    '</tr>'
  }).join('')

  return [
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.4">',
    '<h2 style="margin:0 0 10px 0">Time Off Request Receipt</h2>',
    '<div style="color:#555;margin-bottom:12px">Automated copy of your Slack submission.</div>',
    '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%"><tbody>',
    tr,
    '</tbody></table>',
    '</div>'
  ].join('')
}

function buildReceiptText_(it, d, userEmail) {
  const lines = [
    'Time Off Request Receipt',
    `Request Kind: ${d.requestKind || ''}`,
    `Slack User: ${it.userName ? '@' + it.userName : (it.user || '')}`,
    `Email: ${userEmail}`,
    '',
    `Start: ${d.startDate || ''} ${d.startTime || ''}`.trim(),
    `End: ${d.endDate || ''} ${d.endTime || ''}`.trim(),
    `AP request: ${d.isApRequest || ''} (eligible: ${d.apEligible || ''})`
  ]

  if (d.requestKind === 'AP') {
    lines.push('')
    lines.push('AP Coverage:')
    lines.push(summarizeApCoveragePlain_(d.apCoverageJson))
  } else {
    lines.push('')
    lines.push(`Off-site Assignment: ${d.offsiteAssignment || ''}`)
    lines.push(`Coverage Needed: ${d.coverageNeeded || ''}`)
    lines.push(`Has Sub Plans: ${d.hasSubPlans || ''}`)
    lines.push(`Sub Plans Link: ${d.subPlansLink || ''}`)
    lines.push(`Notes: ${d.notes || ''}`)
  }

  return lines.join('\n')
}

function summarizeApCoveragePlain_(apCoverageJson) {
  try {
    const sections = JSON.parse(apCoverageJson || '[]')
    const filled = sections
      .filter(s => (s.coverage && s.coverage.length) || String(s.plan || '').trim())
      .map(s => `Set ${s.index}: ${(s.coverage || []).join(', ') || '—'} | ${String(s.plan || '').trim()}`)
    return filled.length ? filled.join('\n') : '(none provided)'
  } catch (e) {
    return '(unable to parse)'
  }
}

function safeSendEmail_(p, dbg) {
  try {
    const opts = { htmlBody: p.htmlBody || '' }
    if (cfg.FROM_ALIAS) opts.from = cfg.FROM_ALIAS
    MailApp.sendEmail(p.to, p.subject, p.textBody || stripHtml_(p.htmlBody || ''), opts)
  } catch (err) {
    dbg.appendRow([new Date(), 'safeSendEmail_', 'FAIL', p.to, p.subject, String(err && err.stack || err)])
  }
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
}

function stripHtml_(s) {
  return String(s).replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim()
}

// =====================
// OPTIONAL: Slack signature verification
// =====================
function verifySlackSignature_(e) {
  try {
    const headers = e?.headers || {}
    const timestamp = headers['X-Slack-Request-Timestamp'] || headers['x-slack-request-timestamp']
    const signature = headers['X-Slack-Signature'] || headers['x-slack-signature']
    if (!timestamp || !signature) return false

    const now = Math.floor(Date.now() / 1000)
    if (Math.abs(now - Number(timestamp)) > 60 * 5) return false

    const body = e.postData?.contents || ''
    const baseString = `v0:${timestamp}:${body}`
    const hash = Utilities.computeHmacSha256Signature(baseString, cfg.SLACK_SIGNING_SECRET)
    const hex = hash.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('')
    const expected = `v0=${hex}`
    return timingSafeEqual_(expected, signature)
  } catch (err) {
    return false
  }
}

function timingSafeEqual_(a, b) {
  if (a.length !== b.length) return false
  let out = 0
  for (let i = 0; i < a.length; i++) out |= a.charCodeAt(i) ^ b.charCodeAt(i)
  return out === 0
}

// =====================
// LEGACY / SEPARATE FORM SUBMIT WEBHOOK CODE
// (Leaving it here since you pasted it, but it is unrelated to the Slack modal flow)
// =====================
function getCfg_() {
  return {
    SHEET_NAME: 'Responses',
    SLACK_WEBHOOK_URL: PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL')
  }
}

function onFormSubmit(e) {
  const cfg = getCfg_()

  try {
    const sh = e.range.getSheet()
    if (cfg.SHEET_NAME && sh.getName() !== cfg.SHEET_NAME) return

    const row = e.range.getRow()
    const lastCol = sh.getLastColumn()

    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String)
    const values = sh.getRange(row, 1, 1, lastCol).getValues()[0]

    const rowObj = {}
    headers.forEach((h, i) => rowObj[h || `COL_${i + 1}`] = values[i])

    const title = '📥 New spreadsheet entry received'
    const lines = [
      `*Sheet:* ${sh.getName()}`,
      `*Row:* ${row}`,
      rowObj['Email Address'] ? `*Email:* ${rowObj['Email Address']}` : null,
      rowObj['Name'] ? `*Name:* ${rowObj['Name']}` : null
    ].filter(Boolean)

    postToSlack_(cfg.SLACK_WEBHOOK_URL, `${title}\n${lines.join('\n')}`)
  } catch (err) {
    console.error(err)
  }
}

function postToSlack_(webhookUrl, text) {
  if (!webhookUrl) throw new Error('Missing SLACK_WEBHOOK_URL')

  const payload = { text }

  UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  })
}


function flushFutureTimeOffQueue_() {
  const cfg = getConfig_()
  const props = PropertiesService.getScriptProperties()

  // allow a new trigger to be created next submission
  props.deleteProperty(cfg.FUTURE_FLUSH_TRIGGER_PROP_KEY)

  // drain queue
  const raw = props.getProperty(cfg.FUTURE_QUEUE_PROP_KEY)
  const items = raw ? JSON.parse(raw) : []
  if (!items.length) return

  // clear queue
  props.deleteProperty(cfg.FUTURE_QUEUE_PROP_KEY)

  // open spreadsheet + sheets
  const ss = SpreadsheetApp.openByUrl(cfg.SHEET_URL)
  const sh = ss.getSheetByName(cfg.FUTURE_SHEET_NAME) || ss.insertSheet(cfg.FUTURE_SHEET_NAME)
  const dbg = ss.getSheetByName(cfg.DEBUG_SHEET) || ss.insertSheet(cfg.DEBUG_SHEET)

  // ---------- local helpers ----------
  const nowIso = () => new Date().toISOString()

  const ensureHeader_ = () => {
    if (sh.getLastRow() > 0) return
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

  const getSlackEmailByUserId_ = (userId) => {
    if (!userId) return ''
    try {
      const res = UrlFetchApp.fetch(
        'https://slack.com/api/users.info?user=' + encodeURIComponent(userId),
        {
          method: 'get',
          headers: { Authorization: 'Bearer ' + cfg.SLACK_BOT_TOKEN },
          muteHttpExceptions: true
        }
      )
      const data = JSON.parse(res.getContentText() || '{}')
      if (!data.ok) return ''
      return data.user?.profile?.email || ''
    } catch (e) {
      dbg.appendRow([new Date(), 'getSlackEmailByUserId_', 'EXCEPTION', String(e && e.stack || e)])
      return ''
    }
  }

  const formatAlertText_ = (it, d, rowNumber) => {
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

  const buildDmReceiptText_ = (it, d, rowNumber) => {
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
      rowNumber ? `*Receipt ID:* Row ${rowNumber}` : null
    ].filter(Boolean)

    return lines.join('\n')
  }

  const safeSendEmail_ = (to, subject, textBody) => {
    try {
      if (!to) return
      const opts = {}
      if (cfg.EMAIL_ALIAS) opts.from = cfg.EMAIL_ALIAS
      MailApp.sendEmail(to, subject, textBody, opts)
    } catch (e) {
      dbg.appendRow([new Date(), 'safeSendEmail_', 'FAIL', to, subject, String(e && e.stack || e)])
    }
  }

  const buildEmailSubject_ = (it, d) => {
    const who = it.userName ? `@${it.userName}` : (it.user || 'Unknown')
    const start = d.startDate || ''
    const end = d.endDate || ''
    const span = (start || end) ? ` ${start}${end ? ` → ${end}` : ''}` : ''
    return `Time Off Receipt: ${d.requestKind || 'Request'} ${who}${span}`
  }

  const buildEmailText_ = (it, d, userEmail, rowNumber) => {
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
  // ---------- end helpers ----------

  ensureHeader_()

  // build rows first
  const startRow = sh.getLastRow() + 1
  const rows = []
  const emailCache = {}

  items.forEach(it => {
    const d = it.data || {}
    const userId = String(it.user || '').trim()

    if (!emailCache[userId]) emailCache[userId] = getSlackEmailByUserId_(userId)

    rows.push([
      nowIso(),
      d.requestKind || '',
      userId,
      it.userName || '',
      emailCache[userId] || '',
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
    ])
  })

  // write to sheet
  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows)

  // per-item notifications/receipts
  items.forEach((it, idx) => {
    const d = it.data || {}
    const rowNumber = startRow + idx
    const userId = String(it.user || '').trim()

    // channel alert
    try {
      const notifyChannelId = cfg.FUTURE_NOTIFY_CHANNEL
      if (notifyChannelId) {
        const text = formatAlertText_(it, d, rowNumber)
        const res = slackApi_(cfg, 'chat.postMessage', { channel: notifyChannelId, text })
        dbg.appendRow([new Date(), 'future_notify', `ok=${res?.ok} err=${res?.error || ''}`, notifyChannelId, rowNumber])
      } else {
        dbg.appendRow([new Date(), 'future_notify', 'SKIP missing FUTURE_NOTIFY_CHANNEL', rowNumber])
      }
    } catch (e) {
      dbg.appendRow([new Date(), 'future_notify', 'EXCEPTION', rowNumber, String(e && e.stack || e)])
    }

    // DM receipt
    try {
      const dmText = buildDmReceiptText_(it, d, rowNumber)
      dmUser_(cfg, userId, dmText)
      dbg.appendRow([new Date(), 'future_dm', 'SENT', userId, rowNumber])
    } catch (e) {
      dbg.appendRow([new Date(), 'future_dm', 'EXCEPTION', userId, rowNumber, String(e && e.stack || e)])
    }

    // email receipt
    try {
      const userEmail = emailCache[userId] || ''
      if (!userEmail) {
        dbg.appendRow([new Date(), 'future_email', 'SKIP no email', userId, rowNumber])
      } else {
        const subject = buildEmailSubject_(it, d)
        const body = buildEmailText_(it, d, userEmail, rowNumber)
        safeSendEmail_(userEmail, subject, body)

        if (cfg.ADMIN_EMAILS && Array.isArray(cfg.ADMIN_EMAILS) && cfg.ADMIN_EMAILS.length) {
          safeSendEmail_(cfg.ADMIN_EMAILS.join(','), '[ADMIN COPY] ' + subject, body)
        }

        dbg.appendRow([new Date(), 'future_email', 'SENT', userEmail, rowNumber])
      }
    } catch (e) {
      dbg.appendRow([new Date(), 'future_email', 'EXCEPTION', userId, rowNumber, String(e && e.stack || e)])
    }
  })

  debugLog_(cfg, 'flushFutureTimeOffQueue_', `Processed ${items.length} item(s)`)
}

