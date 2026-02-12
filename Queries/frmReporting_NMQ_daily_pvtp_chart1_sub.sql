SELECT tblPartTesting.*, tblDropDownsSP.testtype, tblDropDownsSP_1.testclassification, tblDropDownsSP_2.teststatus, tblPartProject.recordId
FROM ((((tblPartTesting LEFT JOIN tblDropDownsSP ON tblPartTesting.testType = tblDropDownsSP.recordid) LEFT JOIN tblDropDownsSP AS tblDropDownsSP_1 ON tblPartTesting.testClassification = tblDropDownsSP_1.recordid) LEFT JOIN tblDropDownsSP AS tblDropDownsSP_2 ON tblPartTesting.testStatus = tblDropDownsSP_2.recordid) LEFT JOIN tblPartInfo ON tblPartTesting.partNumber = tblPartInfo.partNumber) INNER JOIN tblPartProject ON tblPartTesting.projectId = tblPartProject.recordId
WHERE tblPartTesting.plannedStart < Date() AND tblPartTesting.actualStart is null AND tblPartTesting.testStatus < 3;

