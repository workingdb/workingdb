SELECT tblPartTrials.recordId, tblPartTrials.partNumber, tblDropDownsSP.trialStatus, tblPartTrials.trialDate, tblDropDownsSP_1.trialResult, rnotes.reportNote, rnotes.recordId AS noteId, rnotes.updatedBy, rnotes.updatedDate, tblPartInfo.developingLocation, tblPartInfo.unitId, rnotes.dataTag0, tblUnits.unitName
FROM ((tblDropDownsSP RIGHT JOIN ((tblPartTrials LEFT JOIN (SELECT
                recordId,
                refId,
                reportNote,
                updatedBy,
                updatedDate,
                dataTag0
            from
                tblReporting_notes
            where dataTag0 = 'nmq_morning_trials'
        )  AS rnotes ON tblPartTrials.recordId = rnotes.refId) LEFT JOIN tblPartInfo ON tblPartTrials.partInfoId = tblPartInfo.recordId) ON tblDropDownsSP.recordid = tblPartTrials.trialStatus) LEFT JOIN tblDropDownsSP AS tblDropDownsSP_1 ON tblPartTrials.trialResult = tblDropDownsSP_1.recordid) LEFT JOIN tblUnits ON tblPartInfo.unitId = tblUnits.recordID
ORDER BY tblPartTrials.trialDate DESC;

