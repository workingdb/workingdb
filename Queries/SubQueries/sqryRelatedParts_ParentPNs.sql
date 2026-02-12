SELECT tblPartProjectPartNumbers.childPartNumber AS primaryPN, tblPartProject.partNumber AS relatedPN, tblDropDownsSP.childPartNumberType AS TYPE
FROM tblDropDownsSP INNER JOIN (tblPartProjectPartNumbers INNER JOIN tblPartProject ON tblPartProjectPartNumbers.projectId = tblPartProject.recordId) ON tblDropDownsSP.recordid = tblPartProjectPartNumbers.childPartNumberType;

