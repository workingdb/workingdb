SELECT 'Due' AS PPAP_Date, Count(PPAPdue) as dateCount FROM frmReporting_NMQ_daily_PPAP_chart1_sub
UNION ALL
SELECT 'Submission' AS PPAP_Date, Count(PPAPsubmit) as dateCount FROM frmReporting_NMQ_daily_PPAP_chart1_sub
UNION ALL SELECT 'Approval' AS PPAP_Date, Count(PPAPapproval) as dateCount FROM frmReporting_NMQ_daily_PPAP_chart1_sub;

