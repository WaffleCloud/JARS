// ===== VIEW SUBMISSIONS =====
function handleSameDayPartialAbsenceSubmit_(cfg, payload) {
  const cb = payload?.view?.callback_id || ''
  if (cb !== 'same_day_partial_absence') return json_(200, { response_action: 'clear' })

  const state = payload?.view?.state?.values || {}

  const parsed = parseSameDayPartialAbsenceState_(state)
  const errors = validateSameDayPartialAbsence_(cfg, parsed)

  if (Object.keys(errors).length) {
    return json_(200, { response_action: 'errors', errors })
  }

  const notifyChannel = pickPartialAbsenceNotifyChannel_(cfg, parsed.isLateArrival, parsed.isEarlyDeparture)

  enqueuePartialAbsence_(cfg, {
    submittedAt: new Date().toISOString(),
    user: payload.user,
    team: payload.team,
    date: parsed.date,
    startTime: parsed.startTime,
    endTime: parsed.endTime,
    coverage: parsed.coverageVals,
    isLateArrival: parsed.isLateArrival,
    isEarlyDeparture: parsed.isEarlyDeparture,
    hasSubPlans: parsed.hasSubPlans,
    subPlansLink: parsed.subPlansLink,
    notifyChannel
  })


  return json_(200, { response_action: 'clear' })
}

function parseSameDayPartialAbsenceState_(state) {
  const date = getDateValue_(state, 'date_block', 'absence_date')

  const startHour = getStaticSelectValue_(state, 'start_hour_block', 'start_hour')
  const startMin = getStaticSelectValue_(state, 'start_min_block', 'start_min')
  const endHour = getStaticSelectValue_(state, 'end_hour_block', 'end_hour')
  const endMin = getStaticSelectValue_(state, 'end_min_block', 'end_min')

  const coverageVals = getMultiSelectValues_(state, 'coverage_block', 'coverage_items') || []

  const isEarlyDepartureVal = getStaticSelectValue_(state, 'early_departure_block', 'early_departure')
  const hasSubPlansVal = getStaticSelectValue_(state, 'has_subplans_block', 'has_subplans')
  const subPlansLinkRaw = getPlainTextValue_(state, 'subplans_link_block', 'subplans_link') || ''

  const startTime = (startHour && startMin) ? `${startHour}:${startMin}` : ''
  const endTime = (endHour && endMin) ? `${endHour}:${endMin}` : ''

  return {
    date,
    startHour,
    startMin,
    endHour,
    endMin,
    startTime,
    endTime,
    coverageVals,
    isEarlyDepartureVal,
    hasSubPlansVal,
    subPlansLink: subPlansLinkRaw.trim()
  }
}

function validateSameDayPartialAbsence_(cfg, p) {
  const errors = {}

  if (!p.date) errors.date_block = 'Please choose a date.'

  if (!p.startHour) errors.start_hour_block = 'Please choose a start hour.'
  if (!p.startMin) errors.start_min_block = 'Please choose a start minute.'
  if (!p.endHour) errors.end_hour_block = 'Please choose an end hour.'
  if (!p.endMin) errors.end_min_block = 'Please choose an end minute.'

  if (!p.coverageVals.length) errors.coverage_block = 'Please select coverage needed.'
  if (!p.isEarlyDepartureVal) errors.early_departure_block = 'Please choose Yes or No.'
  if (!p.hasSubPlansVal) errors.has_subplans_block = 'Please choose Yes or No.'
  if (!p.subPlansLink) errors.subplans_link_block = 'Please enter a link/location or N/A.'

  if (p.startTime && p.endTime) {
    const startMinTotal = hhmmToMinutes_(p.startTime)
    const endMinTotal = hhmmToMinutes_(p.endTime)
    if (endMinTotal <= startMinTotal) errors.end_hour_block = 'End time must be after start time.'
  }

  // derived flags
  p.isEarlyDeparture = (p.isEarlyDepartureVal === 'YES')

  const lateThresholdMin = hhmmToMinutes_(cfg.LATE_THRESHOLD_HHMM)
  p.isLateArrival = p.startTime ? (hhmmToMinutes_(p.startTime) > lateThresholdMin) : false

  p.hasSubPlans = (p.hasSubPlansVal === 'YES')

  return errors
}


function pickPartialAbsenceNotifyChannel_(cfg, isLateArrival, isEarlyDeparture) {
  // Late arrival wins, even if also early departure
  if (isLateArrival) return cfg.LATE_NOTIFY_CHANNEL || ''

  // Only early departure
  if (isEarlyDeparture) return cfg.EARLY_NOTIFY_CHANNEL || ''

  // If neither flag is set, fallback (or blank)
  return cfg.EARLY_NOTIFY_CHANNEL || ''
}


// ===== MODAL BUILDER =====
function buildSameDayPartialAbsenceModal_(cfg, payload) {
  const today = Utilities.formatDate(new Date(), cfg.TZ, 'yyyy-MM-dd')

  const hourOptions = buildHourOptions_(6, 18) // 06–18
  const minuteOptions = buildMinuteOptions_()  // 00–59

  const startHourDefault = '07'
  const startMinDefault = '15'
  const endHourDefault = '03'
  const endMinDefault = '45'

  return {
    type: 'modal',
    callback_id: 'same_day_partial_absence',
    title: { type: 'plain_text', text: 'Partial Absence' },
    submit: { type: 'plain_text', text: 'Submit' },
    close: { type: 'plain_text', text: 'Cancel' },
    blocks: [
      {
        type: 'section',
        text: { type: 'mrkdwn', text: '*Same Day Partial Absence*\n(Late arrival / early departure)' }
      },
      {
        type: 'input',
        block_id: 'date_block',
        label: { type: 'plain_text', text: 'Date' },
        element: {
          type: 'datepicker',
          action_id: 'absence_date',
          initial_date: today,
          placeholder: { type: 'plain_text', text: 'Select a date' }
        }
      },

      // Start time (Hour)
      {
        type: 'input',
        block_id: 'start_hour_block',
        label: { type: 'plain_text', text: 'Start time — hour' },
        element: {
          type: 'static_select',
          action_id: 'start_hour',
          placeholder: { type: 'plain_text', text: 'Select hour' },
          options: hourOptions,
          initial_option: findOptionByValue_(hourOptions, startHourDefault)
        }
      },
      // Start time (Minute)
      {
        type: 'input',
        block_id: 'start_min_block',
        label: { type: 'plain_text', text: 'Start time — minute' },
        element: {
          type: 'static_select',
          action_id: 'start_min',
          placeholder: { type: 'plain_text', text: 'Select minute' },
          options: minuteOptions,
          initial_option: findOptionByValue_(minuteOptions, startMinDefault)
        }
      },

      // End time (Hour)
      {
        type: 'input',
        block_id: 'end_hour_block',
        label: { type: 'plain_text', text: 'End time — hour' },
        element: {
          type: 'static_select',
          action_id: 'end_hour',
          placeholder: { type: 'plain_text', text: 'Select hour' },
          options: hourOptions,
          initial_option: findOptionByValue_(hourOptions, endHourDefault)
        }
      },
      // End time (Minute)
      {
        type: 'input',
        block_id: 'end_min_block',
        label: { type: 'plain_text', text: 'End time — minute' },
        element: {
          type: 'static_select',
          action_id: 'end_min',
          placeholder: { type: 'plain_text', text: 'Select minute' },
          options: minuteOptions,
          initial_option: findOptionByValue_(minuteOptions, endMinDefault)
        }
      },

      {
        type: 'input',
        block_id: 'coverage_block',
        label: { type: 'plain_text', text: 'Coverage needed?' },
        element: {
          type: 'multi_static_select',
          action_id: 'coverage_items',
          placeholder: { type: 'plain_text', text: 'Select coverage items' },
          option_groups: coverageOptionGroups_()
        }
      },

      // New: Early Departure?
      {
        type: 'input',
        block_id: 'early_departure_block',
        label: { type: 'plain_text', text: 'Will you be departing earlier than your scheduled departure time?' },
        element: {
          type: 'static_select',
          action_id: 'early_departure',
          placeholder: { type: 'plain_text', text: 'Select one' },
          options: [
            { text: { type: 'plain_text', text: 'Yes' }, value: 'YES' },
            { text: { type: 'plain_text', text: 'No' }, value: 'NO' }
          ]
        }
      },

      {
        type: 'input',
        block_id: 'has_subplans_block',
        label: { type: 'plain_text', text: 'Do you have sub plans?' },
        element: {
          type: 'static_select',
          action_id: 'has_subplans',
          placeholder: { type: 'plain_text', text: 'Select one' },
          options: [
            { text: { type: 'plain_text', text: 'Yes' }, value: 'YES' },
            { text: { type: 'plain_text', text: 'No' }, value: 'NO' }
          ]
        }
      },
      {
        type: 'input',
        block_id: 'subplans_link_block',
        label: { type: 'plain_text', text: 'Link/Location to sub plans (if none or n/a write n/a)' },
        element: {
          type: 'plain_text_input',
          action_id: 'subplans_link',
          multiline: false,
          placeholder: { type: 'plain_text', text: 'Paste link or type N/A' }
        }
      }
    ]
  }
}

function findOptionByValue_(options, value) {
  for (const opt of options || []) {
    if (opt && opt.value === value) return opt
  }
  return null
}

function buildHourOptions_(startHour, endHour) {
  const options = []
  for (let h = startHour; h <= endHour; h++) {
    const hh = String(h).padStart(2, '0') // "06"
    const label = formatTimeForDisplay_(`${hh}:00`).replace(':00', '') // "6 AM", "7 PM"
    options.push({
      text: { type: 'plain_text', text: label },
      value: hh
    })
  }
  return options
}

function buildMinuteOptions_() {
  const options = []
  for (let m = 0; m <= 59; m++) {
    const mm = String(m).padStart(2, '0')
    options.push({
      text: { type: 'plain_text', text: mm },
      value: mm
    })
  }
  return options
}



// ===== QUEUE + WORKER =====
function enqueuePartialAbsence_(cfg, item) {
  const cache = CacheService.getScriptCache()
  const lock = LockService.getScriptLock()

  // Never wait long on the Slack view_submission hot path
  const got = lock.tryLock(50)

  const key = cfg.PARTIAL_ABSENCE_QUEUE_KEY
  const fallbackKey = cfg.PARTIAL_ABSENCE_QUEUE_KEY + '_FALLBACK'

  try {
    const targetKey = got ? key : fallbackKey

    const raw = cache.get(targetKey)
    const arr = raw ? safeJsonParse_(raw) : []
    const safeArr = Array.isArray(arr) ? arr : []

    safeArr.push(item)

    // Keep TTL long enough for worker to catch it
    cache.put(targetKey, JSON.stringify(safeArr), 21600)
  } finally {
    if (got) lock.releaseLock()
  }
}






function partialAbsenceWorker_() {
  if (isSlackHot_()) return

  const cfg = getConfig_()
  const cache = CacheService.getScriptCache()

  const key = cfg.PARTIAL_ABSENCE_QUEUE_KEY
  const fallbackKey = cfg.PARTIAL_ABSENCE_QUEUE_KEY + '_FALLBACK'

  // Drain both queues up front
  const rawMain = cache.get(key)
  const rawFallback = cache.get(fallbackKey)

  const mainArr = rawMain ? safeJsonParse_(rawMain) : []
  const fallbackArr = rawFallback ? safeJsonParse_(rawFallback) : []

  const batch = []
    .concat(Array.isArray(mainArr) ? mainArr : [])
    .concat(Array.isArray(fallbackArr) ? fallbackArr : [])

  if (!batch.length) return

  cache.remove(key)
  cache.remove(fallbackKey)

  const failed = []

  for (const item of batch) {
    try {
      let rowNumber = ''
      if (cfg.SHEET_URL) rowNumber = appendPartialAbsenceToSheet_(cfg, item)

      const ch = item?.notifyChannel || ''
      if (ch) postPartialAbsenceToSlack_(cfg, item, ch)

      sendPartialAbsenceReceiptDm_(cfg, item, rowNumber)
    } catch (err) {
      failed.push(item)
      try {
        Logger.log('partialAbsenceWorker_ failed: ' + String(err && err.stack || err))
      } catch (e) {}
    }
  }

  if (failed.length) {
    cache.put(key, JSON.stringify(failed), 21600)
  }
}



// ===== SHEET OUTPUT =====
function appendPartialAbsenceToSheet_(cfg, item) {
  const ss = SpreadsheetApp.openByUrl(cfg.SHEET_URL)
  const sh = ss.getSheetByName(cfg.PARTIAL_ABSENCE_SHEET_NAME) || ss.insertSheet(cfg.PARTIAL_ABSENCE_SHEET_NAME)

  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'Timestamp',
      'User ID',
      'User Name',
      'Date',
      'Start Time',
      'End Time',
      'Late Arrival?',
      'Early Departure?',
      'Coverage',
      'Has Sub Plans?',
      'Sub Plans Link/Location',
      'Notify Channel'
    ])
  }

  const coverageText = (item.coverage && item.coverage.length)
    ? item.coverage.join(', ')
    : ''

  sh.appendRow([
    new Date(),
    item.user && item.user.id ? item.user.id : '',
    (item.user && (item.user.username || item.user.name)) ? (item.user.username || item.user.name) : '',
    item.date || '',
    item.startTime || '',
    item.endTime || '',
    item.isLateArrival ? 'Yes' : 'No',
    item.isEarlyDeparture ? 'Yes' : 'No',
    coverageText,
    item.hasSubPlans ? 'Yes' : 'No',
    item.subPlansLink || '',
    item.notifyChannel || ''
  ])

  return sh.getLastRow()
}

// ===== SLACK NOTIFICATION =====
function postPartialAbsenceToSlack_(cfg, item, channelId) {
  const who = (item.user && (item.user.username || item.user.name)) ? (item.user.username || item.user.name) : 'Someone'

  const rangeText = `${formatTimeForDisplay_(item.startTime)} – ${formatTimeForDisplay_(item.endTime)}`
  const covText = (item.coverage && item.coverage.length) ? item.coverage.join(', ') : 'N/A'

  const tags = []
  if (item.isLateArrival) tags.push(':alarm_clock: *Late arrival*')
  if (item.isEarlyDeparture) tags.push(':door: *Early departure*')

  const tagLine = tags.length ? `${tags.join('  ')}\n` : ''

  const text =
    `:hourglass_flowing_sand: *Same Day Partial Absence*\n` +
    tagLine +
    `• *Who:* ${who}\n` +
    `• *Date:* ${item.date}\n` +
    `• *Time:* ${rangeText}\n` +
    `• *Coverage:* ${covText}\n` +
    `• *Sub plans:* ${item.hasSubPlans ? 'Yes' : 'No'}\n` +
    `• *Plans link/location:* ${item.subPlansLink || 'N/A'}`

  const res = slackApi_(cfg, 'chat.postMessage', {
    channel: channelId,
    text
  })

  // debugLog_(cfg, 'postPartialAbsenceToSlack_', `channel=${channelId} res=${JSON.stringify(res || {})}`)
}

// ===== DM RECEIPT =====
function sendPartialAbsenceReceiptDm_(cfg, item, sheetRowNumber) {
  const userId = item?.user?.id || ''
  if (!userId) {
    // debugLog_(cfg, 'sendPartialAbsenceReceiptDm_', 'ABORT missing userId')
    return
  }

  const text = buildPartialAbsenceReceiptDmText_(item, sheetRowNumber)

  const open = slackApi_(cfg, 'conversations.open', { users: userId })
  // debugLog_(cfg, 'sendPartialAbsenceReceiptDm_', 'OPEN ' + JSON.stringify(open || {}))

  if (!open || !open.ok) return

  const channelId = open?.channel?.id || ''
  if (!channelId) {
    // debugLog_(cfg, 'sendPartialAbsenceReceiptDm_', 'OPEN_NO_CHANNEL_ID')
    return
  }

  const post = slackApi_(cfg, 'chat.postMessage', { channel: channelId, text })
  // debugLog_(cfg, 'sendPartialAbsenceReceiptDm_', 'POST ' + JSON.stringify(post || {}))
}

function buildPartialAbsenceReceiptDmText_(item, sheetRowNumber) {
  const who = (item.user && (item.user.username || item.user.name)) ? (item.user.username || item.user.name) : 'You'

  const rangeText = `${formatTimeForDisplay_(item.startTime)} – ${formatTimeForDisplay_(item.endTime)}`
  const covText = (item.coverage && item.coverage.length) ? item.coverage.join(', ') : 'N/A'

  const tags = []
  if (item.isLateArrival) tags.push('Late arrival')
  if (item.isEarlyDeparture) tags.push('Early departure')

  const lines = [
    '✅ *Partial Absence Receipt*',
    `• *Who:* ${who}`,
    `• *Date:* ${item.date}`,
    `• *Time:* ${rangeText}`,
    tags.length ? `• *Type:* ${tags.join(' + ')}` : '',
    `• *Coverage:* ${covText}`,
    `• *Sub plans:* ${item.hasSubPlans ? 'Yes' : 'No'}`,
    `• *Plans link/location:* ${item.subPlansLink || 'N/A'}`
  ].filter(Boolean)

  if (sheetRowNumber) lines.push(`• *Receipt ID:* Row ${sheetRowNumber}`)

  lines.push('\nIf anything looks wrong, reply in your coach channel and we’ll fix it.')

  return lines.join('\n')
}
