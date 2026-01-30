doPost and blockActions needs to return within 3 seconds or slack will time out
    minimize items in this path (no spreadsheet calls, no extra function calls just views.open)

SET UP
run setupTenMinWorkerTrigger

triggers 
set up sweepTenMinEmails to run between every 10-60 minutes