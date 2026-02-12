SELECT sqryStepApprovalTracker_Steps_FindCurrentGate.partNumber, Max(tblPartSteps.indexOrder) AS [index], sqryStepApprovalTracker_Steps_FindCurrentGate.MinOfplannedDate AS gatePlannedDate, Max(tblPartSteps.dueDate) AS gateLastPillar, Max(sqryFindNextPillar.due) AS soonestPillar, tblPartSteps.partGateId
FROM sqryStepApprovalTracker_Steps_FindCurrentGate INNER JOIN (tblPartSteps INNER JOIN sqryFindNextPillar ON (tblPartSteps.dueDate = sqryFindNextPillar.due) AND (sqryFindNextPillar.partNumber = tblPartSteps.partNumber)) ON sqryStepApprovalTracker_Steps_FindCurrentGate.MinOfrecordId = tblPartSteps.partGateId
GROUP BY sqryStepApprovalTracker_Steps_FindCurrentGate.partNumber, sqryStepApprovalTracker_Steps_FindCurrentGate.MinOfplannedDate, tblPartSteps.partGateId, tblPartSteps.closeDate
HAVING (((tblPartSteps.closeDate) Is Null));

