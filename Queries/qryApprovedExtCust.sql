SELECT dbo_tblDRS.Control_Number, dbo_tblDRS.Assignee, dbo_tblDRS.Completed_Date, IIf(IIf(IsNull([Completed_Date]),Date(),[Completed_Date])<=IIf(IsNull([Adjusted_Due_Date]),[Due_Date],[Adjusted_Due_Date]),"On Time","Late") AS Judgment, dbo_tblDRS.Adjusted_Due_Date
FROM dbo_tblDRS
WHERE (((dbo_tblDRS.Completed_Date) Is Not Null) AND ((dbo_tblDRS.Request_Type)=1) AND ((dbo_tblDRS.Approval_Status)=2)) OR (((dbo_tblDRS.Completed_Date) Is Not Null) AND ((dbo_tblDRS.Request_Type)=2) AND ((dbo_tblDRS.Approval_Status)=2)) OR (((dbo_tblDRS.Completed_Date) Is Not Null) AND ((dbo_tblDRS.Request_Type)=3) AND ((dbo_tblDRS.Approval_Status)=2));

