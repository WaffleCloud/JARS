

//===================
//DO POST
//===================



function doPost(e) {
  const cfg = getConfig_()

  try {
    const payloadStr = (e && e.parameter && e.parameter.payload) ? e.parameter.payload : ''
    if (!payloadStr) return ContentService.createTextOutput('ok')

    const payload = JSON.parse(payloadStr)

    if (payload.type === 'block_actions') {
      handleBlockActions_(cfg, payload)
      return ContentService.createTextOutput('')
    }

    if (payload.type === 'view_submission') {
      return handleViewSubmission_(cfg, payload)
    }

    return ContentService.createTextOutput('ok')
  } catch (err) {
    debugLog_(cfg, 'doPost', String(err && err.stack || err))
    // Always return 200 so Slack doesn't retry
    return ContentService.createTextOutput('ok')
  }
}

//===================
//HANDLE BLOCK ACTIONS
//===================



function handleBlockActions_(cfg, payload) {
  const triggerId = payload.trigger_id
  const actionId = payload.actions && payload.actions[0] && payload.actions[0].action_id

  // Helpful tracing when Slack says a modal "doesn't open"
  // (Most often: invalid_auth, trigger_expired, or missing scope)
  const openModal_ = (view, label) => {
    try {
      const res = slackApi_(cfg, 'views.open', { trigger_id: triggerId, view })
      debugLog_(cfg, 'handleBlockActions_', `${label} views.open ok=${res?.ok} error=${res?.error || ''}`)
      return res
    } catch (err) {
      debugLog_(cfg, 'handleBlockActions_', `${label} views.open EXCEPTION ${String(err && err.stack || err)}`)
      return null
    }
  }

  if (actionId === 'ten_min_late_button') {
    openModal_(buildTenMinLateConfirmModal_(payload), 'ten_min_late_button')
    return ContentService.createTextOutput('')
  }

  if (actionId === 'partial_absence_button') {
    openModal_(buildSameDayPartialAbsenceModal_(cfg, payload), 'partial_absence_button')
    return ContentService.createTextOutput('')
  }

  if (actionId === 'same_day_full_absence') {
    openModal_(buildFullAbsenceModal_({
    openedAtIso: new Date().toISOString(),
    userId: payload.user?.id || ''
    }), 'same_day_full_absence')
    return ContentService.createTextOutput('')
  }

  if (actionId === 'future_time_off_button') {
    openModal_(buildTimeOffModalStep1_({ cfg, privateMeta: { userId: payload.user?.id || '' } }), 'future_time_off_button')
    return ContentService.createTextOutput('')
  }

  if (actionId === 'future_timeoff_back') {
  handleBackButton_(cfg, payload)
  return ContentService.createTextOutput('')
}


  return ContentService.createTextOutput('')
}


// =====================
// VIEW SUBMISSIONS ROUTER
// =====================
function handleViewSubmission_(cfgOrPayload, maybePayload) {
  // Support both call styles:
  // 1) handleViewSubmission_(cfg, payload)
  // 2) handleViewSubmission_(payload)
  const cfg = (maybePayload ? cfgOrPayload : getConfig_())
  const payload = (maybePayload ? maybePayload : cfgOrPayload) || {}

  const view = payload.view || {}
  const cb = view.callback_id || ''

  try {
    // TEN MIN confirm
    if (cb === 'ten_min_late_confirm') {
      return handleTenMinLateSubmit_(cfg, payload)
    }

    // Partial Absence
    if (cb === 'same_day_partial_absence') {
      return handleSameDayPartialAbsenceSubmit_(cfg, payload)
    }


    // Full Absence
    if (cb === 'same_day_full_absence') {
      return handleFullAbsenceSubmitFast_(cfg, payload)
    }

    // Future Time Off
   if (cb === 'timeoff_step1' || cb === 'timeoff_pto' || cb === 'timeoff_ap') {
    return handleFutureViewSubmission_(cfg, payload)
}


    // Unknown callback: just close/clear
    return json_(200, { response_action: 'clear' })
  } catch (err) {
    // Never let Slack hang due to an exception
    debugLog_(cfg, 'handleViewSubmission_', `ERROR cb=${cb} ${String(err && err.stack || err)}`)
    return json_(200, { response_action: 'clear' })
  }
}




