SELECT Count(tblPartAttachmentsSP.ID) AS CountOfID, tblPartAttachmentsSP.partStepId
FROM tblPartAttachmentsSP
GROUP BY tblPartAttachmentsSP.partStepId;

