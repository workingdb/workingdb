SELECT Min(tblPartSteps.dueDate) AS due, tblPartSteps.partNumber
FROM sqryStepApprovalTracker_Steps_FindCurrentGate INNER JOIN tblPartSteps ON sqryStepApprovalTracker_Steps_FindCurrentGate.MinOfrecordId = tblPartSteps.partGateId
GROUP BY tblPartSteps.partNumber, tblPartSteps.closeDate
HAVING (((Min(tblPartSteps.dueDate)) Is Not Null) AND ((tblPartSteps.closeDate) Is Null))
ORDER BY Min(tblPartSteps.dueDate);

