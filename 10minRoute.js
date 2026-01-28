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
            `You’re about to notify the team that you will be *10 minutes or less late*.\n\n` +
            `Click *Submit* to send, or *Cancel* to abort.`
        }
      }
    ]
  }
}


// =====================
// QUEUE + WORKER
// =====================
function enqueueTenMinLate_(cfg, item) {
  const cache = CacheService.getScriptCache()
  const raw = cache.get(cfg.TEN_QUEUE_KEY)
  const arr = raw ? JSON.parse(raw) : []

  arr.push(item)

  cache.put(cfg.TEN_QUEUE_KEY, JSON.stringify(arr), 21600) // 6 hours
  

    ensureWorkerTriggerOnce_(
    'tenMinLateWorker',
    cfg.TEN_WORKER_FLAG,
    1
  )

}



function tenMinLateWorker() {
  const cfg = getConfig_()
  debugLog_(cfg, 'tenMinLateWorker_', 'START')
  const cache = CacheService.getScriptCache()
  const raw = cache.get(cfg.TEN_QUEUE_KEY)
  const arr = raw ? JSON.parse(raw) : []
  if (!arr.length) return

  cache.remove(cfg.TEN_QUEUE_KEY)

  for (const item of arr) {
    try {
      let rowNumber = ''
      let meta = { name: '', email: '' }

      if (cfg.SHEET_URL) {
        rowNumber = appendTenMinLateToSheet_(cfg, item)
        meta = getNameAndEmailFromResponsesRow_(cfg, rowNumber)
      }

      if (cfg.TEN_NOTIFY_CHANNEL) {
        postTenMinLateToSlack_(cfg, item, meta.name)
      }

      sendTenMinLateReceiptDm_(cfg, item, rowNumber, meta.name)

      if (meta.email) {
        sendTenMinLateEmailReceipt_(cfg, meta.email, item, rowNumber, meta.name)
      }
    } catch (err) {
      //debugLog_(cfg, 'tenMinLateWorker', String(err && err.stack || err))
    }
  }
}



// =====================
// SHEET OUTPUT
// =====================
function appendTenMinLateToSheet_(cfg, item) {
  const ss = SpreadsheetApp.openByUrl(cfg.SHEET_URL)
  const sh = ss.getSheetByName(cfg.TEN_MIN_SHEET_NAME) || ss.insertSheet(cfg.TEN_MIN_SHEET_NAME)

  // Ensure headers exist (A–F)
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'User ID', 'User Name', 'Type', 'Real Name', 'Email'])
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

  // Write only A–D; formulas in E/F can populate based on B (User ID)
  sh.getRange(targetRow, 1, 1, 4).setValues([[
    new Date(),
    item?.user?.id || '',
    item?.user?.username || item?.user?.name || '',
    item?.type || 'TEN_MIN_LATE'
  ]])

  return targetRow
}


function getNameAndEmailFromResponsesRow_(cfg, rowNumber) {
  if (!cfg.SHEET_URL || !rowNumber) return { name: '', email: '' }

  const ss = SpreadsheetApp.openByUrl(cfg.SHEET_URL)
  const sh = ss.getSheetByName(cfg.TEN_MIN_SHEET_NAME)
  if (!sh) return { name: '', email: '' }

  // Give formulas a moment to populate (if E/F are formula-driven)
  SpreadsheetApp.flush()

  // Columns: E=5, F=6
  // Read E:F from the row
  let name = ''
  let email = ''

  // Try a couple times in case formulas recalc slightly after flush
  for (let i = 0; i < 3; i++) {
    const vals = sh.getRange(rowNumber, 5, 1, 2).getValues()[0] // [E, F]
    name = String(vals[0] || '').trim()
    email = String(vals[1] || '').trim()

    if (name || email) break
    Utilities.sleep(250)
    SpreadsheetApp.flush()
  }

  return { name, email }
}


function handleTenMinLateSubmit_(cfg, payload) {
  const user = payload.user || {}

  const item = {
    type: 'TEN_MIN_LATE',
    ts: new Date().toISOString(),
    user: {
      id: user.id || '',
      username: user.username || '',
      name: user.name || ''
    }
  }

  enqueueTenMinLate_(cfg, item)

  // closes modal immediately
  return json_(200, { response_action: 'clear' })
}


// =====================
// SLACK NOTIFY + DM
// =====================
function postTenMinLateToSlack_(cfg, item, displayName) {
  const fallback = item?.user?.username || item?.user?.name || 'Employee'
  const who = (displayName && String(displayName).trim()) ? String(displayName).trim() : fallback

  const text =
    `:alarm_clock:\n` +
    `• *Who:* ${who}\n`

  const res = slackApi_(cfg, 'chat.postMessage', {
    channel: cfg.TEN_NOTIFY_CHANNEL,
    text
  })

  debugLog_(cfg, 'postTenMinLateToSlack_', `channel=${cfg.TEN_NOTIFY_CHANNEL} res=${JSON.stringify(res || {})}`)
}

function sendTenMinLateReceiptDm_(cfg, item, sheetRowNumber, displayName) {
  const userId = item?.user?.id || ''
  if (!userId) {
    debugLog_(cfg, 'sendTenMinLateReceiptDm_', 'ABORT missing userId')
    return
  }

  const text = buildTenMinLateReceiptDmText_(item, sheetRowNumber, displayName)

  const open = slackApi_(cfg, 'conversations.open', { users: userId })
  debugLog_(cfg, 'sendTenMinLateReceiptDm_', 'OPEN ' + JSON.stringify(open || {}))
  if (!open || !open.ok) return

  const channelId = open?.channel?.id || ''
  if (!channelId) {
    debugLog_(cfg, 'sendTenMinLateReceiptDm_', 'OPEN_NO_CHANNEL_ID')
    return
  }

  const post = slackApi_(cfg, 'chat.postMessage', { channel: channelId, text })
  debugLog_(cfg, 'sendTenMinLateReceiptDm_', 'POST ' + JSON.stringify(post || {}))
}

function buildTenMinLateReceiptDmText_(item, sheetRowNumber, displayName) {
   const fallback = item?.user?.username || item?.user?.name || 'You'
   const who = (displayName && String(displayName).trim()) ? String(displayName).trim() : fallback

  const lines = [
    '✅ *10 minutes or less late heads-up sent*',
    `• *Who:* ${who}`,
  ]

  if (sheetRowNumber) lines.push(`• *Receipt ID:* Row ${sheetRowNumber}`)
  lines.push('\nIf anything looks wrong, message in your coach channel.')

  return lines.join('\n')
}



// =====================
// EMAIL RECEIPT TO USER
// =====================

// function sendTenMinLateEmailReceipt_(cfg, toEmail, item, rowNumber, displayName) {
//   const email = String(toEmail || '').trim()
//   if (!email) return

//   const fallback = item?.user?.username || item?.user?.name || 'You'
//   const who = (displayName && String(displayName).trim()) ? String(displayName).trim() : fallback

//   const subject = `${cfg.EMAIL_SUBJECT_PREFIX || 'Late Heads-up Receipt'}`
//   const body =
//     `Late Heads-up Sent\n\n` +
//     `Who: ${who}\n` +
//     `ETA: 10 minutes or less\n` +
//     (rowNumber ? `Receipt ID: Row ${rowNumber}\n` : '') +
//     `\nIf anything looks wrong, reply in your coach channel.`

//   try {
//     MailApp.sendEmail({
//       to: email,
//       subject,
//       body
//     })
//     debugLog_(cfg, 'sendTenMinLateEmailReceipt_', `SENT to=${email}`)
//   } catch (err) {
//     debugLog_(cfg, 'sendTenMinLateEmailReceipt_', `ERROR to=${email} err=${String(err && err.stack || err)}`)
//   }
// }




// =====================
// EMAIL RECEIPTS
// =====================
// function sendSubmitReceiptEmails_(it, d, userEmailOverride) {
//   const dbg = getDebugSheet_()

//   const userId = String(it.user || '').trim()
//   if (!userId) {
//     dbg.appendRow([new Date(), 'sendSubmitReceiptEmails_', 'ABORT missing userId'])
//     return
//   }

//   const userEmail = String(userEmailOverride || getSlackEmailByUserId_(userId) || '').trim()
//   if (!userEmail) {
//     dbg.appendRow([new Date(), 'sendSubmitReceiptEmails_', `ABORT no email for ${userId}`])
//     return
//   }

//   const subject = buildReceiptSubject_(it, d)
//   const htmlBody = buildReceiptHtml_(it, d, userEmail)
//   const textBody = buildReceiptText_(it, d, userEmail)

//   safeSendEmail_({ to: userEmail, subject, htmlBody, textBody }, dbg)

//   if (CFG.ADMIN_EMAILS && CFG.ADMIN_EMAILS.length) {
//     safeSendEmail_({
//       to: CFG.ADMIN_EMAILS.join(','),
//       subject: '[ADMIN COPY] ' + subject,
//       htmlBody,
//       textBody
//     }, dbg)
//   }

//   dbg.appendRow([new Date(), 'sendSubmitReceiptEmails_', 'SENT', userEmail, (CFG.ADMIN_EMAILS || []).join(',')])
// }


