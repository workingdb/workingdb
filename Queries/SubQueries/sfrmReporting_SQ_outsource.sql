SELECT tblPartInfo.partNumber, tblPartInfo.description, tblDropDownsSP.partType, tblPartInfo.SOPdate, tblPartOutsourceInfo.outsourceVendor, rnotes.reportNote, rnotes.recordId AS noteId, rnotes.updatedBy, rnotes.updateddate, rnotes.dataTag0
FROM ((tblPartInfo LEFT JOIN tblPartOutsourceInfo ON tblPartInfo.outsourceInfoId = tblPartOutsourceInfo.recordId) LEFT JOIN (SELECT
            recordId,
            refId,
            reportNote,
            updatedBy,
            updatedDate,
            dataTag0
        from
            tblReporting_notes
        where
            dataTag0 = 'sq_outsource'
    )  AS rnotes ON tblPartInfo.recordId = rnotes.refId) LEFT JOIN tblDropDownsSP ON tblPartInfo.partType = tblDropDownsSP.recordid
WHERE (((tblPartInfo.partNumber) In (SELECT
            partNumber
        FROM
            tblPartProject
        WHERE
            projectStatus = 1
    )));

