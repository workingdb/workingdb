SELECT tblPartTeam.partNumber, tblPartTeam.person
FROM tblPartTeam INNER JOIN tblPermissions ON tblPartTeam.person = tblPermissions.User
WHERE (((tblPermissions.Dept)="Project") AND ((tblPermissions.Level)="Engineer"));

