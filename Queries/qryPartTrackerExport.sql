SELECT tblPartProject.partNumber, tblPartGates.gateTitle, tblPartSteps.stepType, tblPartSteps.responsible, tblPartSteps.status, tblPartSteps.dueDate
FROM tblPartSteps RIGHT JOIN (tblPartGates RIGHT JOIN tblPartProject ON tblPartGates.projectId = tblPartProject.recordId) ON tblPartSteps.partGateId = tblPartGates.recordId
GROUP BY tblPartProject.partNumber, tblPartGates.gateTitle, tblPartSteps.stepType, tblPartSteps.responsible, tblPartSteps.status, tblPartSteps.dueDate
HAVING (((tblPartProject.partNumber)="29393"))
ORDER BY tblPartSteps.dueDate;

