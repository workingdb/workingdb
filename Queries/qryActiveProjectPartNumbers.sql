SELECT tblPartSteps.partNumber AS PN
FROM tblPartSteps
WHERE (((tblPartSteps.closeDate) Is Null))
GROUP BY tblPartSteps.partNumber;

