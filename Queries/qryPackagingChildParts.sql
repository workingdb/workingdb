SELECT tblPartProjectPartNumbers.childPartNumber AS PN
FROM tblPartProjectPartNumbers INNER JOIN tblPartSteps ON tblPartProjectPartNumbers.projectId = tblPartSteps.partProjectId
WHERE (((tblPartSteps.closeDate) Is Null))
GROUP BY tblPartProjectPartNumbers.childPartNumber, tblPartProjectPartNumbers.childPartNumberType
HAVING (((tblPartProjectPartNumbers.childPartNumberType)=2));

