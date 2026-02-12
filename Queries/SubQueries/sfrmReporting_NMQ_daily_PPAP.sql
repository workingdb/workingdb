SELECT tblPartInfo.recordId, tblPartInfo.partNumber, tblPartInfo.PPAPdue, tblPartInfo.PPAPsubmit, tblPartInfo.PPAPapproval, rnotes.reportNote, rnotes.recordId AS noteId, rnotes.updatedBy, rnotes.updatedDate, tblPartInfo.developingLocation, tblPartInfo.unitId, rnotes.dataTag0, tblPrograms.modelCode, tblUnits.unitName
FROM ((tblPartInfo LEFT JOIN (SELECT
                recordId,
                refId,
                reportNote,
                updatedBy,
                updatedDate,
                dataTag0
            from
                tblReporting_notes
            where dataTag0 = 'nmq_morning_ppap'
        )  AS rnotes ON tblPartInfo.recordId = rnotes.refId) LEFT JOIN tblPrograms ON tblPartInfo.programId = tblPrograms.ID) LEFT JOIN tblUnits ON tblPartInfo.unitId = tblUnits.recordID
WHERE (((tblPartInfo.PPAPdue)>#6/1/2025#) AND ((tblPartInfo.[partNumber]) In (SELECT partNumber FROM tblPartProject WHERE projectStatus = 1)));

