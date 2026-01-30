

//===================
//DO POST
//===================

function doGet(e) {
  return ContentService.createTextOutput('ok')
}


function doPost(e) {
  const cfg = getConfig_()
  
  try {
    const payloadStr = (e && e.parameter && e.parameter.payload) ? e.parameter.payload : ''
    if (!payloadStr) return ContentService.createTextOutput('ok')
      
      const payload = JSON.parse(payloadStr)

    // IMPORTANT: ACK block_actions as fast as possible (no sheet writes)
    if (payload.type === 'block_actions') {
      markSlackHot_(3)
      handleBlockActions_(cfg, payload)
      return ContentService.createTextOutput('ok') // <- changed from ''
    }

    if (payload.type === 'view_submission') {
      markSlackHot_(3)
      return handleViewSubmission_(cfg, payload)
    }

    return ContentService.createTextOutput('ok')
  } catch (err) {
    // This is not hot-path (only on errors), okay to log to sheet
    debugLog_(cfg, 'doPost', String(err && err.stack || err))
    return ContentService.createTextOutput('ok')
  }
}


//===================
//HANDLE BLOCK ACTIONS
//===================



function handleBlockActions_(cfg, payload) {
  const triggerId = payload.trigger_id
  const actionId = payload.actions && payload.actions[0] && payload.actions[0].action_id

  // FAST logging only (Logger) - do NOT write to Sheets here
  const fastLog_ = (msg) => {
    try { Logger.log(msg) } catch (e) {}
  }

  const openModal_ = (view, label) => {
    try {
      const res = slackApi_(cfg, 'views.open', { trigger_id: triggerId, view })
      fastLog_(`${label} views.open ok=${res?.ok} error=${res?.error || ''}`)
      return res
    } catch (err) {
      fastLog_(`${label} views.open EXCEPTION ${String(err && err.stack || err)}`)
      return null
    }
  }

  if (actionId === 'ten_min_late_button') {
    openModal_(buildTenMinLateConfirmModal_(payload), 'ten_min_late_button')
    return ContentService.createTextOutput('ok')
  }

  if (actionId === 'partial_absence_button') {
    openModal_(buildSameDayPartialAbsenceModal_(cfg, payload), 'partial_absence_button')
    return ContentService.createTextOutput('ok')
  }

  if (actionId === 'same_day_full_absence') {
    openModal_(buildFullAbsenceModal_({
      openedAtIso: new Date().toISOString(),
      userId: payload.user?.id || ''
    }), 'same_day_full_absence')
    return ContentService.createTextOutput('ok')
  }

  if (actionId === 'future_time_off_button') {
    openModal_(buildTimeOffModalStep1_({ cfg, privateMeta: { userId: payload.user?.id || '' } }), 'future_time_off_button')
    return ContentService.createTextOutput('ok')
  }

  // NOTE: back button is usually a block_action too — keep it fast
  if (actionId === 'future_timeoff_back') {
    try {
      handleBackButton_(cfg, payload)
    } catch (err) {
      fastLog_(`future_timeoff_back EXCEPTION ${String(err && err.stack || err)}`)
    }
    return ContentService.createTextOutput('ok')
  }

  return ContentService.createTextOutput('ok')
}



// =====================
// VIEW SUBMISSIONS ROUTER
// =====================
function handleViewSubmission_(cfg, payload) {
  const t0 = Date.now()
  const cb = payload?.view?.callback_id || ''

  try {
    let res
    if (cb === 'ten_min_late_confirm') res = handleTenMinLateSubmit_(cfg, payload)
    else if (cb === 'same_day_partial_absence') res = handleSameDayPartialAbsenceSubmit_(cfg, payload)
    else if (cb === 'same_day_full_absence') res = handleFullAbsenceSubmitFast_(cfg, payload)
    else if (cb === 'timeoff_step1' || cb === 'timeoff_pto' || cb === 'timeoff_ap') res = handleFutureViewSubmission_(cfg, payload)
    else res = json_(200, { response_action: 'clear' })

    Logger.log('view_submission cb=' + cb + ' ms=' + (Date.now() - t0))
    return res
  } catch (err) {
    Logger.log('view_submission ERROR cb=' + cb + ' ms=' + (Date.now() - t0) + ' err=' + String(err && err.stack || err))
    return json_(200, { response_action: 'clear' })
  }
}





