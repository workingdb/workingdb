SELECT tblPermissions.User, Sum(IIf(IsNull([dbo_tblTimeTrackChild]![TimeTrack_Work_Hours]),0,[dbo_tblTimeTrackChild]![TimeTrack_Work_Hours])) AS Hours, dbo_tblTimeTrackChild.TimeTrack_Work_Date, 6 AS Target
FROM dbo_tblTimeTrackChild INNER JOIN tblPermissions ON dbo_tblTimeTrackChild.Associate_ID = tblPermissions.ID
WHERE (((dbo_tblTimeTrackChild.TimeTrack_Work_Date)>=Date()-7))
GROUP BY tblPermissions.User, dbo_tblTimeTrackChild.TimeTrack_Work_Date;

