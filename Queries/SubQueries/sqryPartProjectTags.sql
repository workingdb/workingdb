SELECT PP.recordId, DD_Status.partProjectStatus, DD_Type.partProjectType
FROM ((tblPartProject AS PP INNER JOIN tblPartProjectTemplate AS PT ON PP.projectTemplateId = PT.recordId) INNER JOIN tblDropDownsSP AS DD_Status ON PP.projectStatus = DD_Status.recordid) INNER JOIN tblDropDownsSP AS DD_Type ON PT.templateType = DD_Type.recordid;

