// ===== VIEW SUBMISSIONS =====
function handleSameDayPartialAbsenceSubmit_(cfg, payload) {
  const view = payload.view || {}
  const callbackId = view.callback_id

  if (callbackId !== 'same_day_partial_absence') {
    return json_(200, { response_action: 'clear' })
  }

  const state = view.state && view.state.values ? view.state.values : {}

  const date = getDateValue_(state, 'date_block', 'absence_date')

  const startHour = getStaticSelectValue_(state, 'start_hour_block', 'start_hour') // "07"
  const startMin = getStaticSelectValue_(state, 'start_min_block', 'start_min') // "15"
  const endHour = getStaticSelectValue_(state, 'end_hour_block', 'end_hour')     // "14"
  const endMin = getStaticSelectValue_(state, 'end_min_block', 'end_min')        // "05"

  const coverageVals = getMultiSelectValues_(state, 'coverage_block', 'coverage_items')
  const isEarlyDepartureVal = getStaticSelectValue_(state, 'early_departure_block', 'early_departure') // YES|NO
  const hasSubPlans = getStaticSelectValue_(state, 'has_subplans_block', 'has_subplans') // YES|NO
  const subPlansLink = getPlainTextValue_(state, 'subplans_link_block', 'subplans_link')

  const errors = {}

  if (!date) errors['date_block'] = 'Please choose a date.'

  if (!startHour) errors['start_hour_block'] = 'Please choose a start hour.'
  if (!startMin) errors['start_min_block'] = 'Please choose a start minute.'
  if (!endHour) errors['end_hour_block'] = 'Please choose an end hour.'
  if (!endMin) errors['end_min_block'] = 'Please choose an end minute.'

  if (!coverageVals || !coverageVals.length) errors['coverage_block'] = 'Please select coverage needed.'
  if (!isEarlyDepartureVal) errors['early_departure_block'] = 'Please choose Yes or No.'
  if (!hasSubPlans) errors['has_subplans_block'] = 'Please choose Yes or No.'
  if (!subPlansLink || !subPlansLink.trim()) errors['subplans_link_block'] = 'Please enter a link/location or N/A.'

  const startTime = (startHour && startMin) ? `${startHour}:${startMin}` : ''
  const endTime = (endHour && endMin) ? `${endHour}:${endMin}` : ''

  if (startTime && endTime) {
    const startMinTotal = hhmmToMinutes_(startTime)
    const endMinTotal = hhmmToMinutes_(endTime)
    if (endMinTotal <= startMinTotal) errors['end_hour_block'] = 'End time must be after start time.'
  }

  const isEarlyDeparture = isEarlyDepartureVal === 'YES'

  
  const lateThresholdMin = hhmmToMinutes_(cfg.LATE_THRESHOLD_HHMM)
  const isLateArrival = startTime ? (hhmmToMinutes_(startTime) > lateThresholdMin) : false


  if (Object.keys(errors).length) {
    return json_(500, {
      response_action: 'errors',
      errors
    })
  }

  const notifyChannel = pickPartialAbsenceNotifyChannel_(cfg, isLateArrival)

  enqueuePartialAbsence_(cfg, {
    submittedAt: new Date().toISOString(),
    user: payload.user,
    team: payload.team,
    date,
    startTime,
    endTime,
    coverage: coverageVals,
    isLateArrival,
    isEarlyDeparture,
    hasSubPlans: hasSubPlans === 'YES',
    subPlansLink: subPlansLink.trim(),
    notifyChannel
  })

  return json_(200, { response_action: 'clear' })
}

function pickPartialAbsenceNotifyChannel_(cfg, isLateArrival, isEarlyDeparture) {
    return isLateArrival
    ? (cfg.LATE_NOTIFY_CHANNEL || '')
    : (cfg.EARLY_NOTIFY_CHANNEL || '') // your normal partial-day channel
}

// ===== MODAL BUILDER =====
function buildSameDayPartialAbsenceModal_(cfg, payload) {
  const today = Utilities.formatDate(new Date(), cfg.TZ, 'yyyy-MM-dd')

  const hourOptions = buildHourOptions_(6, 18) // 06–18
  const minuteOptions = buildMinuteOptions_()  // 00–59

  const startHourDefault = '07'
  const startMinDefault = '15'

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
          options: hourOptions
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
          options: minuteOptions
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
  const raw = cache.get(cfg.PARTIAL_ABSENCE_QUEUE_KEY)
  const arr = raw ? JSON.parse(raw) : []
  arr.push(item)
  cache.put(cfg.PARTIAL_ABSENCE_QUEUE_KEY, JSON.stringify(arr), 21600)
  ensureWorkerTriggerOnce_('partialAbsenceWorker_', cfg.PARTIAL_ABSENCE_WORKER_FLAG, 1)
}



function partialAbsenceWorker_() {
  const cfg = getConfig_()
  const cache = CacheService.getScriptCache()
  const raw = cache.get(cfg.PARTIAL_ABSENCE_QUEUE_KEY)
  const arr = raw ? JSON.parse(raw) : []
  if (!arr.length) return

  cache.remove(cfg.PARTIAL_ABSENCE_QUEUE_KEY)

  for (const item of arr) {
    try {
      let rowNumber = ''

      if (cfg.SHEET_URL) {
        rowNumber = appendPartialAbsenceToSheet_(cfg, item)
      }

      const ch = item?.notifyChannel || ''
      if (ch) {
        postPartialAbsenceToSlack_(cfg, item, ch)
      }

      sendPartialAbsenceReceiptDm_(cfg, item, rowNumber)
    } catch (err) {
      debugLog_(cfg, 'partialAbsenceWorker_', String(err && err.stack || err))
    }
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

  debugLog_(cfg, 'postPartialAbsenceToSlack_', `channel=${channelId} res=${JSON.stringify(res || {})}`)
}

// ===== DM RECEIPT =====
function sendPartialAbsenceReceiptDm_(cfg, item, sheetRowNumber) {
  const userId = item?.user?.id || ''
  if (!userId) {
    debugLog_(cfg, 'sendPartialAbsenceReceiptDm_', 'ABORT missing userId')
    return
  }

  const text = buildPartialAbsenceReceiptDmText_(item, sheetRowNumber)

  const open = slackApi_(cfg, 'conversations.open', { users: userId })
  debugLog_(cfg, 'sendPartialAbsenceReceiptDm_', 'OPEN ' + JSON.stringify(open || {}))

  if (!open || !open.ok) return

  const channelId = open?.channel?.id || ''
  if (!channelId) {
    debugLog_(cfg, 'sendPartialAbsenceReceiptDm_', 'OPEN_NO_CHANNEL_ID')
    return
  }

  const post = slackApi_(cfg, 'chat.postMessage', { channel: channelId, text })
  debugLog_(cfg, 'sendPartialAbsenceReceiptDm_', 'POST ' + JSON.stringify(post || {}))
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
