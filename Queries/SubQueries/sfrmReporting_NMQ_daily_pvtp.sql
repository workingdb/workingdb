SELECT tblPartTesting.recordId, tblPartProject.partNumber, tblDropDownsSP.testtype, tblPartTesting.plannedStart, tblPartTesting.plannedEnd, rnotes.reportNote, rnotes.recordId AS noteId, rnotes.updatedBy, rnotes.updatedDate, tblPartInfo.developingLocation, tblPartInfo.unitId, rnotes.dataTag0, tblPartTesting.actualStart, tblPartTesting.actualEnd, tblPartTesting.pass, tblDropDownsSP_2.teststatus, tblDropDownsSP_1.testclassification, tblPartProject.recordId, tblPartTesting.testStatus, tblUnits.unitName
FROM (tblPartProject LEFT JOIN (((((tblPartTesting LEFT JOIN (SELECT
                recordId,
                refId,
                reportNote,
                updatedBy,
                updatedDate,
                dataTag0
            from
                tblReporting_notes
            where dataTag0 = 'nmq_morning_trials'
        )  AS rnotes ON tblPartTesting.recordId = rnotes.refId) LEFT JOIN tblPartInfo ON tblPartTesting.partNumber = tblPartInfo.partNumber) LEFT JOIN tblDropDownsSP ON tblPartTesting.testType = tblDropDownsSP.recordid) LEFT JOIN tblDropDownsSP AS tblDropDownsSP_1 ON tblPartTesting.testClassification = tblDropDownsSP_1.recordid) LEFT JOIN tblDropDownsSP AS tblDropDownsSP_2 ON tblPartTesting.testStatus = tblDropDownsSP_2.recordid) ON tblPartProject.recordId = tblPartTesting.projectId) LEFT JOIN tblUnits ON tblPartInfo.unitId = tblUnits.recordID
WHERE (((tblPartProject.projectStatus)=1))
ORDER BY tblPartTesting.plannedStart DESC;

