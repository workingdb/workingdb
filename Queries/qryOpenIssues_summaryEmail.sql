SELECT tblPartIssues.recordId, tblPartIssues.partNumber, tblDropDownsSP.issueType, tblDropDownsSP_1.issueSource, tblPartIssues.closeDate, tblPartIssues.dueDate, tblPartIssues.inCharge
FROM (tblPartIssues LEFT JOIN tblDropDownsSP ON tblPartIssues.issueType = tblDropDownsSP.recordID) LEFT JOIN tblDropDownsSP AS tblDropDownsSP_1 ON tblPartIssues.issueSource = tblDropDownsSP_1.recordID
WHERE (((tblPartIssues.closeDate) Is Null));

