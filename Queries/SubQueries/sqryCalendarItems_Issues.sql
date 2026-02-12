SELECT tblPartIssues.recordId AS ID, "Part Issue" AS TYPE, tblPartIssues.partNumber AS partNumber, tblDropDownsSP.issueType AS [Action], tblPartIssues.dueDate AS due, tblPartIssues.inCharge AS Person
FROM tblPartIssues LEFT JOIN tblDropDownsSP ON tblPartIssues.issueType = tblDropDownsSP.recordid
GROUP BY tblPartIssues.recordId, "Part Issue", tblPartIssues.partNumber, tblDropDownsSP.issueType, tblPartIssues.dueDate, tblPartIssues.inCharge, tblPartIssues.closeDate
HAVING (
        ((tblPartIssues.dueDate) IS NOT NULL)
        AND ((tblPartIssues.inCharge) IS NOT NULL)
        AND ((tblPartIssues.closeDate) IS NULL)
    );

