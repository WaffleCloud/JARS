// =====================
// 10minuteRoute.gs (READY TO PASTE)
// Ten-minutes-or-less late heads-up
// =====================


// =====================
// MODAL
// =====================
function buildTenMinLateConfirmModal_(payload) {
  return {
    type: 'modal',
    callback_id: 'ten_min_late_confirm',
    title: { type: 'plain_text', text: 'Late Heads-up' },
    submit: { type: 'plain_text', text: 'Submit' },
    close: { type: 'plain_text', text: 'Cancel' },
    blocks: [
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text:
            'You’re about to notify the team that you will be *10 minutes or less late*.\n\n' +
            'Click *Submit* to send, or *Cancel* to abort.'
        }
      }
    ]
  }
}


// =====================
// VIEW SUBMISSION HANDLER (HOT PATH)
// =====================
function handleTenMinLateSubmit_(cfg, payload) {
  // keep it tiny + safe
  const userId = payload?.user?.id || ''
  const username = payload?.user?.username || payload?.user?.name || ''

  // Minimal job payload
  const job = {
    ts: new Date().toISOString(),
    userId,
    username
  }

  // enqueue must be FAST
  enqueueTenMinLateJob_(cfg, job)

  // ensure worker exists (optional if you rely on masterWorker trigger)
  // ensureWorkerTriggerOnce_('masterWorker', 'FLAG_MASTER_WORKER_SET_V1', 1)

  return json_(200, { response_action: 'clear' })
}


// =====================
// QUEUE (CacheService) — fast per-job keys
// =====================
// cfg.TEN_QUEUE_KEY is now the INDEX list key
function enqueueTenMinLateJob_(cfg, job) {
  const cache = CacheService.getScriptCache()

  // A unique key per job
  const jobKey = `${cfg.TEN_QUEUE_KEY}:JOB:${Date.now()}:${Math.random().toString(16).slice(2)}`
  cache.put(jobKey, JSON.stringify(job), 21600)

  // Maintain a small index list of job keys
  const rawIndex = cache.get(cfg.TEN_QUEUE_KEY)
  const index = rawIndex ? (safeJsonParse_(rawIndex) || []) : []
  index.push(jobKey)

  // guardrail: keep index from growing forever
  const trimmed = index.length > 200 ? index.slice(index.length - 200) : index

  cache.put(cfg.TEN_QUEUE_KEY, JSON.stringify(trimmed), 21600)
}



// =====================
// VIEW SUBMISSION HANDLER (HOT PATH)
// =====================
function handleTenMinLateSubmit_(cfg, payload) {
  // keep it tiny + safe
  const userId = payload?.user?.id || ''
  const username = payload?.user?.username || payload?.user?.name || ''

  // Minimal job payload
  const job = {
    ts: new Date().toISOString(),
    userId,
    username
  }

  // enqueue must be FAST
  enqueueTenMinLateJob_(cfg, job)

  // ensure worker exists (optional if you rely on masterWorker trigger)
  // ensureWorkerTriggerOnce_('masterWorker', 'FLAG_MASTER_WORKER_SET_V1', 1)

  return json_(200, { response_action: 'clear' })
}


// =====================
// QUEUE (CacheService) — fast per-job keys
// =====================
// cfg.TEN_QUEUE_KEY is now the INDEX list key
function enqueueTenMinLateJob_(cfg, job) {
  const cache = CacheService.getScriptCache()

  // A unique key per job
  const jobKey = `${cfg.TEN_QUEUE_KEY}:JOB:${Date.now()}:${Math.random().toString(16).slice(2)}`
  cache.put(jobKey, JSON.stringify(job), 21600)

  // Maintain a small index list of job keys
  const rawIndex = cache.get(cfg.TEN_QUEUE_KEY)
  const index = rawIndex ? (safeJsonParse_(rawIndex) || []) : []
  index.push(jobKey)

  // guardrail: keep index from growing forever
  const trimmed = index.length > 200 ? index.slice(index.length - 200) : index

  cache.put(cfg.TEN_QUEUE_KEY, JSON.stringify(trimmed), 21600)
}



// =====================
// WORKER
// =====================
function tenMinLateWorker() {
  if (isSlackHot_()) return
  const cfg = getConfig_()

  const cache = CacheService.getScriptCache()

  // Drain index
  const rawIndex = cache.get(cfg.TEN_QUEUE_KEY)
  const index = rawIndex ? (safeJsonParse_(rawIndex) || []) : []
  if (!index.length) return

  // clear index first to avoid double-processing if worker overlaps
  cache.remove(cfg.TEN_QUEUE_KEY)

  for (const jobKey of index) {
    const rawJob = cache.get(jobKey)
    cache.remove(jobKey) // remove immediately (at-least-once behavior)

    if (!rawJob) continue
    const job = safeJsonParse_(rawJob)
    if (!job) continue

    try {
      // Rebuild the old "item" shape only inside the worker
      const item = {
        type: 'TEN_MIN_LATE',
        ts: job.ts,
        user: {
          id: job.userId || '',
          username: job.username || '',
          name: job.username || ''
        }
      }

      let rowNumber = ''

      if (cfg.SHEET_URL) {
        rowNumber = appendTenMinLateToSheet_(cfg, item)
      }

      let meta = { name: '', email: '' }
      if (cfg.SHEET_URL && rowNumber) {
        meta = getNameEmailAndSentColsFromRow_(cfg, rowNumber)
      }

      const who = meta.name || item?.user?.username || item?.user?.name || 'Employee'

      if (cfg.TEN_NOTIFY_CHANNEL) {
        postTenMinLateToSlack_(cfg, item, who)
      }

      sendTenMinLateReceiptDm_(cfg, item, rowNumber, who)
    } catch (err) {
      debugLog_(cfg, 'tenMinLateWorker', String(err && err.stack || err))
    }
  }

  debugLog_(cfg, 'tenMinLateWorker', 'DONE processed=' + index.length)
}




// =====================
// SHEET OUTPUT
// =====================
function appendTenMinLateToSheet_(cfg, item) {
  const ss = SpreadsheetApp.openByUrl(cfg.SHEET_URL)
  const sh = ss.getSheetByName(cfg.TEN_MIN_SHEET_NAME) || ss.insertSheet(cfg.TEN_MIN_SHEET_NAME)

  // Ensure headers exist
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'User ID', 'User Name', 'Type', 'Real Name', 'Email', 'Receipt ID', 'Email Sent?', 'Email Sent At'])
  }

  // Find first empty row in column B (User ID column)
  const startRow = 2
  const col = 2 // B
  const last = sh.getLastRow()
  const numRows = Math.max(last - startRow + 1, 1)

  const colValues = sh.getRange(startRow, col, numRows, 1).getValues().map(r => r[0])
  let targetRow = startRow + colValues.findIndex(v => v === '' || v === null)

  // If none empty found, append after last
  if (targetRow < startRow) targetRow = last + 1

  const u = item && item.user ? item.user : {}

  // Write only A–D; formulas in E/F populate from User ID (B)
  sh.getRange(targetRow, 1, 1, 4).setValues([[
    new Date(),
    u.id || '',
    u.username || u.name || '',
    (item && item.type) ? item.type : 'TEN_MIN_LATE'
  ]])

  return targetRow
}


// Read name/email from SAME row by HEADER names: "Real Name", "Email"
// function getNameAndEmailFromResponsesRow_(cfg, rowNumber) {
//   if (!cfg.SHEET_URL || !rowNumber) return { name: '', email: '' }

//   const ss = SpreadsheetApp.openByUrl(cfg.SHEET_URL)
//   const sh = ss.getSheetByName(cfg.TEN_MIN_SHEET_NAME)
//   if (!sh) return { name: '', email: '' }

//   const lastCol = sh.getLastColumn()
//   if (lastCol < 1) return { name: '', email: '' }

//   const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
//   const nameCol = headers.indexOf('Real Name') + 1
//   const emailCol = headers.indexOf('Email') + 1

//   if (nameCol < 1 || emailCol < 1) {
//     debugLog_(cfg, 'getNameAndEmailFromResponsesRow_', 'Missing headers. nameCol=' + nameCol + ' emailCol=' + emailCol)
//     return { name: '', email: '' }
//   }

//   SpreadsheetApp.flush()

//   let name = ''
//   let email = ''

//   for (let i = 0; i < 5; i++) {
//     name = String(sh.getRange(rowNumber, nameCol).getValue() || '').trim()
//     email = String(sh.getRange(rowNumber, emailCol).getValue() || '').trim()

//     if (name || email) break
//     Utilities.sleep(250)
//     SpreadsheetApp.flush()
//   }

//   return { name: name, email: email }
// }


// =====================
// SLACK NOTIFY + DM RECEIPT
// =====================
function postTenMinLateToSlack_(cfg, item, displayName) {
  const u = item && item.user ? item.user : {}
  const fallback = u.username || u.name || 'Employee'
  const who = (displayName && String(displayName).trim()) ? String(displayName).trim() : fallback

  const text =
    ':alarm_clock:\n' +
    '*Late Heads-up (≤ 10 minutes)*\n' +
    '• *Who:* ' + who

  const res = slackApi_(cfg, 'chat.postMessage', {
    channel: cfg.TEN_NOTIFY_CHANNEL,
    text: text
  })

  debugLog_(cfg, 'postTenMinLateToSlack_', 'channel=' + cfg.TEN_NOTIFY_CHANNEL + ' res=' + JSON.stringify(res || {}))
}

function sendTenMinLateReceiptDm_(cfg, item, sheetRowNumber, displayName) {
  const u = item && item.user ? item.user : {}
  const userId = u.id || ''

  if (!userId) {
    debugLog_(cfg, 'sendTenMinLateReceiptDm_', 'ABORT missing userId')
    return
  }

  const text = buildTenMinLateReceiptDmText_(item, sheetRowNumber, displayName)

  const open = slackApi_(cfg, 'conversations.open', { users: userId })
  debugLog_(cfg, 'sendTenMinLateReceiptDm_', 'OPEN ' + JSON.stringify(open || {}))
  if (!open || !open.ok) return

  const channelId = open.channel && open.channel.id ? open.channel.id : ''
  if (!channelId) {
    debugLog_(cfg, 'sendTenMinLateReceiptDm_', 'OPEN_NO_CHANNEL_ID')
    return
  }

  const post = slackApi_(cfg, 'chat.postMessage', { channel: channelId, text: text })
  debugLog_(cfg, 'sendTenMinLateReceiptDm_', 'POST ' + JSON.stringify(post || {}))
}

function buildTenMinLateReceiptDmText_(item, sheetRowNumber, displayName) {
  const u = item && item.user ? item.user : {}
  const fallback = u.username || u.name || 'You'
  const who = (displayName && String(displayName).trim()) ? String(displayName).trim() : fallback

  const lines = [
    '✅ *10 minutes or less late heads-up sent*',
    '• *Who:* ' + who
  ]

  if (sheetRowNumber) lines.push('• *Receipt ID:* Row ' + sheetRowNumber)
  lines.push('\nIf anything looks wrong, message in your coach channel.')

  return lines.join('\n')
}


// =====================
// EMAIL RECEIPT TO USER
// =====================
// function sendTenMinLateEmailReceipt_(cfg, toEmail, item, rowNumber, displayName) {
//   const email = String(toEmail || '').trim()
//   if (!email) return

//   const u = item && item.user ? item.user : {}
//   const fallback = u.username || u.name || 'You'
//   const who = (displayName && String(displayName).trim()) ? String(displayName).trim() : fallback

//   const subject = (cfg.EMAIL_SUBJECT_PREFIX)
//     ? (cfg.EMAIL_SUBJECT_PREFIX + ' — Late Heads-up')
//     : 'Late Heads-up Receipt'

//   const body =
//     'Late Heads-up Sent\n\n' +
//     'Who: ' + who + '\n' +
//     'ETA: 10 minutes or less\n' +
//     (rowNumber ? ('Receipt ID: Row ' + rowNumber + '\n') : '') +
//     '\nIf anything looks wrong, message in your coach channel.'

//   try {
//     // NOTE: MailApp does NOT support {from: alias}. Use GmailApp if you need alias sending.
//     MailApp.sendEmail(email, subject, body)
//     debugLog_(cfg, 'sendTenMinLateEmailReceipt_', 'SENT to=' + email)
//   } catch (err) {
//     debugLog_(cfg, 'sendTenMinLateEmailReceipt_', 'ERROR to=' + email + ' err=' + String(err && err.stack || err))
//   }
// }

function sweepTenMinEmails() {
  if (isSlackHot_()) return
  const cfg = getConfig_()
  const ss = SpreadsheetApp.openByUrl(cfg.SHEET_URL)
  const sh = ss.getSheetByName(cfg.TEN_MIN_SHEET_NAME)
  if (!sh) return

  const lastRow = sh.getLastRow()
  const lastCol = sh.getLastColumn()
  if (lastRow < 2) return

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim())
  const nameCol = headers.indexOf('Real Name') + 1
  const emailCol = headers.indexOf('Email') + 1
  const sentCol = headers.indexOf('Email Sent?') + 1

  if (emailCol < 1 || sentCol < 1) {
    debugLog_(cfg, 'sweepTenMinEmails', 'Missing Email or Email Sent? headers')
    return
  }

  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues()

  for (let i = 0; i < values.length; i++) {
    const rowIdx = i + 2
    const email = String(values[i][emailCol - 1] || '').trim()
    const sent = String(values[i][sentCol - 1] || '').trim().toLowerCase()

    if (!email) continue
    if (sent === 'yes' || sent === 'true' || sent === 'sent') continue

    const who = nameCol > 0 ? String(values[i][nameCol - 1] || '').trim() : ''

    // You may not have the original item payload here, so send a generic receipt
    try {
      MailApp.sendEmail(
        email,
        (cfg.EMAIL_SUBJECT_PREFIX ? cfg.EMAIL_SUBJECT_PREFIX + ' — Late Heads-up' : 'Late Heads-up Receipt'),
        'Late Heads-up Sent\n\n' +
        (who ? ('Who: ' + who + '\n') : '') +
        'ETA: 10 minutes or less\n' +
        'Receipt ID: Row ' + rowIdx + '\n'
      )

      sh.getRange(rowIdx, sentCol).setValue('YES')
      const sentAtCol = headers.indexOf('Email Sent At') + 1
      if (sentAtCol > 0) sh.getRange(rowIdx, sentAtCol).setValue(new Date())
    } catch (err) {
      debugLog_(cfg, 'sweepTenMinEmails_', 'Row ' + rowIdx + ' email failed: ' + String(err && err.message || err))
    }
  }
}
