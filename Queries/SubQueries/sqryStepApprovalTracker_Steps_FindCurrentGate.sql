SELECT Min(plannedDate) AS MinOfplannedDate, Min(recordId) AS MinOfrecordId, partNumber
FROM tblPartGates
WHERE actualDate Is Null
GROUP BY partNumber;

