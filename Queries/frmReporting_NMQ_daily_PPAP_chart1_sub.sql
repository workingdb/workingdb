SELECT PI.*, PI.recordId
FROM frmReporting_NMQ_daily_PPAP_chart1_sub_1 AS PI
WHERE PPAPsubmit IS NULL AND PPAPdue < Date() + 30;

