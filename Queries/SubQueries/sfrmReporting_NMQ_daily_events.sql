SELECT tblProgramEvents.ID, tblProgramEvents.eventTitle, tblProgramEvents.eventDate, rnotes.reportNote, rnotes.recordId AS noteId, rnotes.updatedBy, rnotes.updatedDate, rnotes.dataTag0, tblProgramEvents.dataSubmitted, tblProgramEvents.dataSubmittedDate, tblDropDownsSP.eventtype, tblDropDownsSP_1.programeventscorrelatedgate, tblPrograms.manufacturer, tblPrograms.OEM, tblPrograms.modelYear, tblPrograms.modelName, tblPrograms.modelCode, tblPrograms.SOPdate, tblProgramEvents.programId
FROM (((SELECT
                recordId,
                refId,
                reportNote,
                updatedBy,
                updatedDate,
                dataTag0
            from
                tblReporting_notes
            where dataTag0 = 'nmq_morning_events'
        )  AS rnotes RIGHT JOIN (tblPrograms RIGHT JOIN tblProgramEvents ON tblPrograms.ID = tblProgramEvents.programId) ON rnotes.refId = tblProgramEvents.ID) LEFT JOIN tblDropDownsSP ON tblProgramEvents.eventType = tblDropDownsSP.recordid) LEFT JOIN tblDropDownsSP AS tblDropDownsSP_1 ON tblProgramEvents.correlatedGate = tblDropDownsSP_1.recordid
ORDER BY tblProgramEvents.eventDate DESC;

