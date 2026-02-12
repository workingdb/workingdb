SELECT tblPartSteps.partProjectId, tblPartSteps.status, Count(tblPartSteps.recordId) AS CountOfrecordId
FROM tblPartSteps INNER JOIN sqryStepApprovalTracker_Steps_FindCurrentGate ON tblPartSteps.partGateId = sqryStepApprovalTracker_Steps_FindCurrentGate.MinOfrecordId
GROUP BY tblPartSteps.partProjectId, tblPartSteps.status;

