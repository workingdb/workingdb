SELECT tblPartProject.partNumber AS primaryPN, tblPartProjectPartNumbers.childPartNumber AS relatedPN, tblDropDownsSP.childPartNumberType AS TYPE
FROM tblDropDownsSP INNER JOIN (tblPartProjectPartNumbers INNER JOIN tblPartProject ON tblPartProjectPartNumbers.projectId = tblPartProject.recordId) ON tblDropDownsSP.recordid = tblPartProjectPartNumbers.childPartNumberType;

