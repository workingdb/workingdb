SELECT tblPartInfo.partNumber, tblPartInfo.description, tblDropDownsSP.partType, tblPrograms.modelCode, tblPrograms.OEM, tblPrograms.modelYear, tblPartInfo.customerPN, tblPartInfo.SOPdate, tblPartOutsourceInfo.outsourceCost, tblPartOutsourceInfo.outsourceVendor, tblPartInfo.monthlyVolume, tblPartInfo.developingLocation, rnotes.reportNote, rnotes.updatedBy, rnotes.updatedDate, rnotes.dataTag0, rnotes.recordId AS noteId
FROM (((tblPartInfo LEFT JOIN tblDropDownsSP ON tblPartInfo.partType = tblDropDownsSP.recordid) LEFT JOIN tblPrograms ON tblPartInfo.programId = tblPrograms.ID) LEFT JOIN tblPartOutsourceInfo ON tblPartInfo.outsourceInfoId = tblPartOutsourceInfo.recordId) LEFT JOIN (SELECT
                recordId,
                refId,
                reportNote,
                updatedBy,
                updatedDate,
                dataTag0
            from
                tblReporting_notes
            where
                dataTag0 = 'pe_partinfo_outsource'
        )  AS rnotes ON tblPartInfo.recordId = rnotes.refId
WHERE tblPartInfo.partNumber IN (SELECT partNumber FROM tblPartProject WHERE projectStatus = 1);

