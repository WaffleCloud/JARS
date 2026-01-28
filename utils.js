
function getConfig_() {
  return {
    TZ: 'America/Chicago',

    SLACK_BOT_TOKEN: PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN'),

    SHEET_URL: 'https://docs.google.com/spreadsheets/d/1ugqee5z754SaeWA6BX6yAT6iFxNsl7YTZeQdadkyOzA/edit?gid=0#gid=0',
    TEN_MIN_SHEET_NAME: '10 min or Less',
    PARTIAL_ABSENCE_SHEET_NAME: 'Partial Day Absence',
    FULL_ABSENCE_SHEET_NAME: 'Full Day Absence',
    FUTURE_SHEET_NAME: 'Future Time Off',
    DEBUG_SHEET: 'SlackDebug',
    EMAIL_ALIAS: 'twade@jeffersonrise.org',

    TEN_NOTIFY_CHANNEL: 'C08K43BP64B',
    LATE_NOTIFY_CHANNEL: 'C08K43BP64B',
    EARLY_NOTIFY_CHANNEL: 'C08K43BP64B',
    FULL_ABSENCE_NOTIFY_CHANNEL: 'C08K43BP64B',
    FUTURE_NOTIFY_CHANNEL: 'C08K43BP64B',

    BUTTON_CHANNEL: 'C08K43BP64B',

    // Cache queues (fast + simple)
    TEN_QUEUE_KEY: 'QUEUE_TEN_MIN_LATE_V1',
    LATE_QUEUE_KEY: 'QUEUE_LATE_ARRIVAL_V1',
    EARLY_QUEUE_KEY: 'QUEUE_EARLY_DEPARTURE_V1',
    CALLOUT_QUEUE_KEY: 'QUEUE_SAME_DAY_CALLOUT_V1',
    TEN_WORKER_FLAG: 'FLAG_TEN_WORKER_SET_V1',
    LATE_WORKER_FLAG: 'FLAG_LATE_WORKER_SET_V1',
    EARLY_WORKER_FLAG: 'FLAG_EARLY_WORKER_SET_V1',
    CALLOUT_WORKER_FLAG: 'FLAG_CALLOUT_WORKER_SET_V1',


    // Future Time Off uses Properties queue + delayed trigger (safer for bigger payloads)
    FUTURE_QUEUE_PROP_KEY: 'QUEUE_FUTURE_TIMEOFF_V1',
    FUTURE_FLUSH_TRIGGER_PROP_KEY: 'QUEUE_FUTURE_TIMEOFF_TRIGGER_SET_V1',
    FUTURE_FLUSH_TRIGGER_DELAY_SECONDS: 30,

    FULL_ABSENCE_SUBMISSION_QUEUE_KEY: 'QUEUE_FULL_ABSENCE_SUBMIT_V1',
    FULL_ABSENCE_JOB_QUEUE_KEY: 'QUEUE_FULL_ABSENCE_JOBS_V1',
    DEBUG_QUEUE_KEY: 'QUEUE_DEBUG_V1',
    PARTIAL_ABSENCE_QUEUE_KEY: 'QUEUE_PARTIAL_ABSENCE_V1',
    LATE_THRESHOLD_HHMM: '07:26',
    PARTIAL_ABSENCE_WORKER_FLAG: 'FLAG_PARTIAL_ABSENCE_WORKER_SET_V1',
    AP_APPROVAL_URL: 'https://....',


  }
}

//===================
//MESSAGE & BUTTON LAUNCHER
//===================

function postAbsenceButtons() {
  const cfg = getConfig_()
  const url = 'https://slack.com/api/chat.postMessage'

  const payload = {
    channel: cfg.BUTTON_CHANNEL,
    text: 'Absence Reporting Options.',
    blocks: [
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: 'Select the type of absence you would like to report.'
        }
      },
      {
        type: 'actions',
        block_id: 'absence_launcher_actions',
        elements: [
          {
            type: 'button',
            action_id: 'ten_min_late_button',
            text: { type: 'plain_text', text: 'I will be ≤ 10 minutes late.' },
            style: 'primary',
            value: 'ten_min_late'
          },
          {
            type: 'button',
            action_id: 'partial_absence_button',
            text: { type: 'plain_text', text: 'Partial Same Day Absence' },
            style: 'primary',
            value: 'partial_absence'
          },
          {
            type: 'button',
            action_id: 'same_day_full_absence',
            text: { type: 'plain_text', text: 'Full Same Day Absence.' },
            style: 'primary',
            value: 'same_day_callout'
          },
          {
            type: 'button',
            action_id: 'future_time_off_button',
            text: { type: 'plain_text', text: 'Future Time off request.' },
            style: 'primary',
            value: 'future_time_off'
          }
        ]
      }
    ]
  }

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + cfg.SLACK_BOT_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  })

  const body = res.getContentText() || '{}'
  const out = JSON.parse(body)

  if (!out.ok) {
    // Slack usually includes "response_metadata.messages" explaining which block failed
    throw new Error('Slack API error: ' + out.error + ' | HTTP ' + res.getResponseCode() + ' | ' + body)
  }
}


// function postAbsenceButtons() {
//   const cfg = getConfig_()
//   const url = 'https://slack.com/api/chat.postMessage'

//   const payload = {
//     channel: cfg.BUTTON_CHANNEL,
//     text: 'Absence Reporting Options.',
//     blocks: [
//       {
//         type: 'section',
//         text: {
//           type: 'mrkdwn',
//           text: 'Select the type of absence you would like to report.'
//         }
//       },
//       {
//         type: 'actions',
//         block_id: 'ten_min_late_launcher_block',
//         elements: [
//           {
//             type: 'button',
//             action_id: 'ten_min_late_button',
//             text: { type: 'plain_text', text: 'I will be ≤ 10 minutes late.' },
//             style: 'primary',
//             value: 'ten_min_late'
//           },
//           {
//             type: 'button',
//             action_id: 'partial_absence_button',
//             text: { type: 'plain_text', text: 'I will be absent part of the day today (Late Arrival/Early Departure/ Partial Day)' },
//             style: 'primary',
//             value: 'late_Arrival'
//           },
//           {
//             type: 'button',
//             action_id: 'same_day_full_absence',
//             text: { type: 'plain_text', text: 'I will be absent for the full day today.' },
//             style: 'primary',
//             value: 'same_day_callout'
//           },
//           {
//             type: 'button',
//             action_id: 'future_time_off_button',
//             text: { type: 'plain_text', text: 'I am requesting time off for a future date.' },
//             style: 'primary',
//             value: 'future_time_off'
//           },
//         ]
//       }
//     ]
//   }

//   const res = UrlFetchApp.fetch(url, {
//     method: 'post',
//     contentType: 'application/json; charset=utf-8',
//     headers: { Authorization: 'Bearer ' + cfg.SLACK_BOT_TOKEN },
//     payload: JSON.stringify(payload),
//     muteHttpExceptions: true
//   })

//   const out = JSON.parse(res.getContentText() || '{}')
//   if (!out.ok) throw new Error('Slack API error: ' + out.error)
// }


// radio helper (since your other helpers are select/multi/plain)
function getRadioValue_(stateValues, blockId, actionId) {
  const block = stateValues[blockId]
  if (!block || !block[actionId]) return ''
  const sel = block[actionId].selected_option
  return sel && sel.value ? sel.value : ''
}

// =====================
// SLACK API + HELPERS
// =====================

function slackApi_(cfg, method, payload) {
  const url = 'https://slack.com/api/' + method

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + cfg.SLACK_BOT_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  })

  return JSON.parse(res.getContentText() || '{}')
}

function json_(code, obj) {
  const out = ContentService
    .createTextOutput(typeof obj === 'string' ? obj : JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)

  if (typeof out.setResponseCode === 'function') out.setResponseCode(code)
  return out
}

function safeJsonParse_(s) {
  try {
    return JSON.parse(s || '')
  } catch (e) {
    return null
  }
}

// Cache queue helpers (for ten/late/early/callout)
function cacheQRead_(key) {
  const raw = CacheService.getScriptCache().get(key)
  const arr = raw ? safeJsonParse_(raw) : []
  return Array.isArray(arr) ? arr : []
}

function cacheQWrite_(key, arr, ttlSeconds) {
  CacheService.getScriptCache().put(key, JSON.stringify(arr), ttlSeconds)
}

function cacheQPush_(key, item, ttlSeconds) {
  const arr = cacheQRead_(key)
  arr.push(item)
  cacheQWrite_(key, arr, ttlSeconds)
}

function cacheQDrain_(key) {
  const cache = CacheService.getScriptCache()
  const raw = cache.get(key)
  const arr = raw ? safeJsonParse_(raw) : []
  cache.remove(key)
  return Array.isArray(arr) ? arr : []
}

// Generic DM helper
function dmUser_(cfg, userId, text) {
  if (!userId) return

  const open = slackApi_(cfg, 'conversations.open', { users: userId })
  if (!open || !open.ok) return

  const channelId = open?.channel?.id || ''
  if (!channelId) return

  slackApi_(cfg, 'chat.postMessage', { channel: channelId, text })
}

//===================
// DEBUG LOGGER
//===================
function debugLog_(cfg, source, message) {
  try {
    Logger.log(`[${source}] ${message}`)

    if (!cfg?.SHEET_URL) return

    const ss = SpreadsheetApp.openByUrl(cfg.SHEET_URL)
    const sh = ss.getSheetByName(cfg.DEBUG_SHEET) || ss.insertSheet(cfg.DEBUG_SHEET)
    sh.appendRow([new Date(), source, message])
  } catch (err) {
    Logger.log(`[debugLog_ ERROR] ${String(err && err.stack || err)}`)
  }
}


// ===== TIME HELPERS =====
function formatTimeForDisplay_(hhmm) {
  const parts = String(hhmm).split(':')
  let h = Number(parts[0] || 0)
  const m = parts[1] || '00'
  const ampm = h >= 12 ? 'PM' : 'AM'
  h = h % 12
  if (h === 0) h = 12
  return `${h}:${m} ${ampm}`
}

function hhmmToMinutes_(hhmm) {
  const parts = String(hhmm).split(':')
  const h = Number(parts[0] || 0)
  const m = Number(parts[1] || 0)
  return h * 60 + m
}

// ===== TIME OPTIONS =====
function buildTimeOptions_(startHHMM, endHHMM, stepMinutes) {
  const start = hhmmToMinutes_(startHHMM)
  const end = hhmmToMinutes_(endHHMM)

  const options = []
  for (let t = start; t <= end; t += stepMinutes) {
    const value = minutesToHHMM_(t) // 24h "HH:mm"
    const label = formatTimeForDisplay_(value) // "h:mm AM/PM"
    options.push({
      text: { type: 'plain_text', text: label },
      value
    })
  }

  return options
}

function minutesToHHMM_(minutes) {
  const h = Math.floor(minutes / 60)
  const m = minutes % 60
  return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0')
}

// ===== COVERAGE OPTIONS (unchanged) =====
function makeOption(text, value) {
  return {
    text: { type: 'plain_text', text },
    value
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


// =========================
// SLACK VIEW STATE HELPERS
// =========================
function getStaticSelectValue_(stateValues, blockId, actionId) {
  const block = stateValues?.[blockId]
  const node = block?.[actionId]
  const sel = node?.selected_option
  return sel?.value || ''
}

function getMultiSelectValues_(stateValues, blockId, actionId) {
  const block = stateValues?.[blockId]
  const node = block?.[actionId]
  const sels = node?.selected_options || []
  return sels.map(o => o?.value).filter(Boolean)
}

function getPlainTextValue_(stateValues, blockId, actionId) {
  const block = stateValues?.[blockId]
  const node = block?.[actionId]
  return node?.value || ''
}

function getDateValue_(stateValues, blockId, actionId) {
  const block = stateValues?.[blockId]
  const node = block?.[actionId]
  return node?.selected_date || ''
}


/* ===========================
   HOT LOG (Logger only)
   =========================== */

function hotLog_(source, message, data) {
  const cfg = getConfig_()
  if (!cfg.DEBUG_MODE) return
  try {
    Logger.log(source + ' | ' + message + ' | ' + JSON.stringify(data || {}))
  } catch (e) {}
}


// =========================
// TRIGGER FLAG OPTIMIZATION
// Hot path: avoid scanning triggers every submission
// =========================
function ensureWorkerTriggerOnce_(handlerName, propKey, everyMinutes) {
  const props = PropertiesService.getScriptProperties()
  const flag = props.getProperty(propKey)
  if (flag === '1') return

  // Double-check in case flag got cleared but trigger exists
  const triggers = ScriptApp.getProjectTriggers()
  const exists = triggers.some(t => t.getHandlerFunction && t.getHandlerFunction() === handlerName)
  if (exists) {
    props.setProperty(propKey, '1')
    return
  }

  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .everyMinutes(everyMinutes || 1)
    .create()

  props.setProperty(propKey, '1')
}

// Optional: call this at the top of each worker so if someone deletes triggers manually,
// the next submission can recreate cleanly.
function markWorkerTriggerPresent_(propKey) {
  PropertiesService.getScriptProperties().setProperty(propKey, '1')
}


/* ===========================
   TIME HELPERS (6:15 cutoff)
   =========================== */

function isBefore615(dateObj, timezone) {
  const hour = parseInt(Utilities.formatDate(dateObj, timezone, 'H'), 10)
  const minute = parseInt(Utilities.formatDate(dateObj, timezone, 'm'), 10)

  if (hour < 6) return true
  if (hour > 6) return false
  return minute <= 15
}

function formatIsoLocal_(iso, timezone) {
  try {
    const d = new Date(iso)
    return Utilities.formatDate(d, timezone, 'EEE M/d h:mm a')
  } catch (e) {
    return iso || ''
  }
}


/* ===========================
   QUEUES (CacheService)
   =========================== */

function qRead_(key) {
  const raw = CacheService.getScriptCache().get(key)
  if (!raw) return []
  const arr = safeJsonParse(raw)
  return Array.isArray(arr) ? arr : []
}

function qWrite_(key, arr, ttlSeconds) {
  CacheService.getScriptCache().put(key, JSON.stringify(arr), ttlSeconds)
}

function qPush_(key, item, maxLen, ttlSeconds) {
  const arr = qRead_(key)
  arr.push(item)
  const trimmed = arr.length > maxLen ? arr.slice(arr.length - maxLen) : arr
  qWrite_(key, trimmed, ttlSeconds)
}

function qShiftBatch_(key, maxItems) {
  const arr = qRead_(key)
  if (!arr.length) return []
  const batch = arr.slice(0, maxItems)
  const remainder = arr.slice(batch.length)
  if (remainder.length) {
    qWrite_(key, remainder, 60 * 30)
  } else {
    CacheService.getScriptCache().remove(key)
  }
  return batch
}

/* ===========================
   WORKERS SETUP
   =========================== */

function setupWorkers() {
  const triggers = ScriptApp.getProjectTriggers()
  triggers.forEach(t => {
    const fn = t.getHandlerFunction()
    if (fn === 'jobWorker_' || fn === 'submissionWorker_' || fn === 'debugWorker_') {
      ScriptApp.deleteTrigger(t)
    }
  })

  ScriptApp.newTrigger('jobWorker_').timeBased().everyMinutes(1).create()
  ScriptApp.newTrigger('submissionWorker_').timeBased().everyMinutes(1).create()
  ScriptApp.newTrigger('debugWorker_').timeBased().everyMinutes(1).create()

  hotLog_('setupWorkers', 'Created workers: jobWorker_, submissionWorker_, debugWorker_', {})
}

// ===================
//SMOKE TEST FUNCTIONS
// ==================
function mark_(label, data) {
  try {
    const cfg = getConfig_()
    const ss = SpreadsheetApp.openByUrl(cfg.SHEET_URL)
    const sh = ss.getSheetByName('SMOKE') || ss.insertSheet('SMOKE')
    sh.appendRow([new Date().toISOString(), label, '', JSON.stringify(data || {})])
  } catch (e) {}
}


function hourOptions_() {
  const out = []
  for (let h = 1; h <= 12; h++) {
    out.push(makeOption(String(h), String(h)))
  }
  return out
}

function minuteOptions_() {
  const out = []
  for (let m = 0; m <= 59; m++) {
    const mm = String(m).padStart(2, '0')
    out.push(makeOption(mm, mm))
  }
  return out
}

function ampmOptions_() {
  return [makeOption('AM', 'AM'), makeOption('PM', 'PM')]
}

function findOption_(options, value) {
  return (options || []).find(o => String(o.value) === String(value)) || null
}

function to24h_(hour12Str, minuteStr, ampmStr) {
  const h12 = Number(hour12Str)
  const mm = String(minuteStr || '00').padStart(2, '0')
  const ap = String(ampmStr || '').toUpperCase()

  if (!h12 || h12 < 1 || h12 > 12) return ''
  if (!/^\d{2}$/.test(mm)) return ''
  if (ap !== 'AM' && ap !== 'PM') return ''

  let h = h12 % 12
  if (ap === 'PM') h += 12

  return String(h).padStart(2, '0') + ':' + mm
}

function format12h_(hour12Str, minuteStr, ampmStr) {
  const h = String(hour12Str || '').trim()
  const mm = String(minuteStr || '').padStart(2, '0')
  const ap = String(ampmStr || '').toUpperCase()
  if (!h || !mm || (ap !== 'AM' && ap !== 'PM')) return ''
  return `${h}:${mm} ${ap}`
}


function safeJsonParse(str) {
  try { return str ? JSON.parse(str) : null } catch (e) { return null }
}



/* ===========================
   META PACKING
   =========================== */

function packMeta_(meta) {
  return JSON.stringify({ meta })
}

function unpackMeta_(privateMetadata) {
  const obj = safeJsonParse(privateMetadata) || {}
  return obj.meta || {}
}