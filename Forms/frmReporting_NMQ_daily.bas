Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function PPAP_applyFilter(filterCond As String, filterName As String)

If Nz(Me.PPAP_fltUnit) <> "" Then
    filterCond = Split(filterCond, " AND unitId")(0) 'remove unit filter if present
    filterCond = filterCond & " AND unitId = " & Me.PPAP_fltUnit
    
    filterName = Split(filterName, " [")(0) 'remove unit from caption if present
    filterName = filterName & " [" & Me.PPAP_fltUnit.column(1) & "]"
End If

Me.sfrmReporting_NMQ_daily_PPAP.Form.filter = filterCond
Me.sfrmReporting_NMQ_daily_PPAP.Form.FilterOn = True

Dim db As Database
Set db = CurrentDb()

Dim qdf As QueryDef
Set qdf = db.QueryDefs("frmReporting_NMQ_daily_PPAP_chart1_sub")

qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE " & filterCond

db.QueryDefs.refresh

Set qdf = Nothing
Set db = Nothing

Me.PPAP_chart1.Requery

Me.PPAP_lblTitle.Caption = filterName

End Function

Function trials_applyFilter(filterCond As String, filterName As String)

If Nz(Me.trials_fltUnit) <> "" Then
    filterCond = Split(filterCond, " AND unitId")(0) 'remove unit filter if present
    filterCond = filterCond & " AND unitId = " & Me.trials_fltUnit
    
    filterName = Split(filterName, " [")(0) 'remove unit from caption if present
    filterName = filterName & " [" & Me.trials_fltUnit.column(1) & "]"
End If

Me.sfrmReporting_NMQ_daily_trials.Form.filter = filterCond
Me.sfrmReporting_NMQ_daily_trials.Form.FilterOn = True

Dim db As Database
Set db = CurrentDb()

Dim qdf As QueryDef
Set qdf = db.QueryDefs("frmReporting_NMQ_daily_trials_chart1_sub")

qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE " & filterCond

db.QueryDefs.refresh

Set qdf = Nothing
Set db = Nothing

Me.trials_chart1.Requery

Me.trials_lblTitle.Caption = filterName

End Function

Function events_applyFilter(filterCond As String, filterName As String)

If Nz(Me.events_fltModel) <> "" Then
    filterCond = Split(filterCond, " AND programId")(0) 'remove model filter if present
    filterCond = filterCond & " AND programId = " & Me.events_fltModel
    
    filterName = Split(filterName, " [")(0) 'remove unit from caption if present
    filterName = filterName & " [" & Me.events_fltModel.column(1) & "]"
End If

Me.sfrmReporting_NMQ_daily_events.Form.filter = filterCond
Me.sfrmReporting_NMQ_daily_events.Form.FilterOn = True

Dim db As Database
Set db = CurrentDb()

Dim qdf As QueryDef
Set qdf = db.QueryDefs("frmReporting_NMQ_daily_events_chart1_sub")

qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE " & filterCond

db.QueryDefs.refresh

Set qdf = Nothing
Set db = Nothing

Me.events_chart1.Requery

Me.events_lblTitle.Caption = filterName

End Function

Function pvtp_applyFilter(filterCond As String, filterName As String)

If Nz(Me.pvtp_fltUnit) <> "" Then
    filterCond = Split(filterCond, " AND unitId")(0) 'remove unit filter if present
    filterCond = filterCond & " AND unitId = " & Me.pvtp_fltUnit
    
    filterName = Split(filterName, " [")(0) 'remove unit from caption if present
    filterName = filterName & " [" & Me.pvtp_fltUnit.column(1) & "]"
End If

Me.sfrmReporting_NMQ_daily_pvtp.Form.filter = filterCond
Me.sfrmReporting_NMQ_daily_pvtp.Form.FilterOn = True

Dim db As Database
Set db = CurrentDb()

Dim qdf As QueryDef
Set qdf = db.QueryDefs("frmReporting_NMQ_daily_pvtp_chart1_sub")

qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE " & filterCond

db.QueryDefs.refresh

Set qdf = Nothing
Set db = Nothing

Me.pvtp_chart1.Requery

Me.pvtp_lblTitle.Caption = filterName

End Function

Private Sub events_export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Reporting_NMQ_daily_events_" & nowString & ".xlsx"
filt = " WHERE " & Me.sfrmReporting_NMQ_daily_events.Form.filter
If Me.sfrmReporting_NMQ_daily_events.Form.FilterOn = False Then filt = ""
sqlString = "SELECT * FROM sfrmReporting_NMQ_daily_events " & filt
                    
Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub events_fltModel_AfterUpdate()
On Error GoTo Err_Handler

If Me.events_lblTitle.Caption = "No Report Selected" Or Me.events_lblTitle.Caption = "Please select a report first" Then 'no report selected
    Me.events_lblTitle.Caption = "Please select a report first"
    Me.ActiveControl = Null
    Exit Sub
End If

Call events_applyFilter(Me.sfrmReporting_NMQ_daily_events.Form.filter, Me.events_lblTitle.Caption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub events_next_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "eventDate < Date() + 90 AND eventDate > Date() - 1"

If Nz(Me.events_fltModel) <> "" Then
    filt = filt & " AND programId = " & Me.events_fltModel
End If

Call events_applyFilter(filt, "Events in the Next 3 Months")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub events_past_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "eventDate < Date()"

If Nz(Me.events_fltModel) <> "" Then
    filt = filt & " AND programId = " & Me.events_fltModel
End If

Call events_applyFilter(filt, "Past Events")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub ppap_export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Reporting_NMQ_daily_PPAP_" & nowString & ".xlsx"
filt = " WHERE " & Me.sfrmReporting_NMQ_daily_PPAP.Form.filter
If Me.sfrmReporting_NMQ_daily_PPAP.Form.FilterOn = False Then filt = ""
sqlString = "SELECT * FROM sfrmReporting_NMQ_daily_PPAP " & filt
                    
Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub PPAP_fltUnit_AfterUpdate()
On Error GoTo Err_Handler

If Me.PPAP_lblTitle.Caption = "No Report Selected" Or Me.PPAP_lblTitle.Caption = "Please select a report first" Then 'no report selected
    Me.PPAP_lblTitle.Caption = "Please select a report first"
    Me.ActiveControl = Null
    Exit Sub
End If

Call PPAP_applyFilter(Me.sfrmReporting_NMQ_daily_PPAP.Form.filter, Me.PPAP_lblTitle.Caption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim db As Database
Set db = CurrentDb()
Dim qdf As QueryDef

Set qdf = db.QueryDefs("frmReporting_NMQ_daily_events_chart1_sub")
qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE ID > 0"

Set qdf = db.QueryDefs("frmReporting_NMQ_daily_trials_chart1_sub")
qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE tblPartTrials.recordId > 0"

Set qdf = db.QueryDefs("frmReporting_NMQ_daily_PPAP_chart1_sub")
qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE recordId > 0"

db.QueryDefs.refresh

Set qdf = Nothing
Set db = Nothing

Me.trials_chart1.Requery
Me.events_chart1.Requery
Me.PPAP_chart1.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub PPAP_subNotApproved_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "PPAPsubmit IS NOT NULL AND PPAPapproval IS NULL"

If Nz(Me.PPAP_fltUnit) <> "" Then
    filt = filt & " AND unitId = " & Me.PPAP_fltUnit
End If

Call PPAP_applyFilter(filt, "Submitted but Not Approved")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub PPAP_upcLate_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "PPAPsubmit IS NULL AND PPAPdue < Date() + 30"

If Nz(Me.PPAP_fltUnit) <> "" Then
    filt = filt & " AND unitId = " & Me.PPAP_fltUnit
End If

Call PPAP_applyFilter(filt, "Upcoming (<3 months) and Late")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub pvtp_empty_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "tblPartProject.recordId NOT IN (SELECT projectId From tblPartTesting GROUP BY projectId HAVING (Count(recordId)>0))" 'no tests found

If Nz(Me.pvtp_fltUnit) <> "" Then
    filt = filt & " AND unitId = " & Me.pvtp_fltUnit
End If

Call pvtp_applyFilter(filt, "Parts with No Testing Enterred")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub pvtp_export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Reporting_NMQ_daily_pvtp_" & nowString & ".xlsx"
filt = " WHERE " & Me.sfrmReporting_NMQ_daily_pvtp.Form.filter
If Me.sfrmReporting_NMQ_daily_pvtp.Form.FilterOn = False Then filt = ""
sqlString = "SELECT * FROM sfrmReporting_NMQ_daily_pvtp " & filt
                    
Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub pvtp_incPastPlan_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "tblPartTesting.plannedEnd < Date() AND tblPartTesting.testStatus < 3" 'Not Started or In Progress

If Nz(Me.pvtp_fltUnit) <> "" Then
    filt = filt & " AND unitId = " & Me.pvtp_fltUnit
End If

Call pvtp_applyFilter(filt, "Incomplete, Past Planned End")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub pvtp_incPastStart_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "tblPartTesting.plannedStart < Date() AND tblPartTesting.actualStart is null AND tblPartTesting.testStatus < 3" 'Not Started or In Progress

If Nz(Me.pvtp_fltUnit) <> "" Then
    filt = filt & " AND unitId = " & Me.pvtp_fltUnit
End If

Call pvtp_applyFilter(filt, "Past Planned Start, No Actual Start")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trials_export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Reporting_NMQ_daily_trials_" & nowString & ".xlsx"
filt = " WHERE " & Me.sfrmReporting_NMQ_daily_trials.Form.filter
If Me.sfrmReporting_NMQ_daily_trials.Form.FilterOn = False Then filt = ""
sqlString = "SELECT * FROM sfrmReporting_NMQ_daily_trials " & filt
                    
Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trials_fltUnit_AfterUpdate()
On Error GoTo Err_Handler

If Me.trials_lblTitle.Caption = "No Report Selected" Or Me.trials_lblTitle.Caption = "Please select a report first" Then 'no report selected
    Me.trials_lblTitle.Caption = "Please select a report first"
    Me.ActiveControl = Null
    Exit Sub
End If

Call trials_applyFilter(Me.sfrmReporting_NMQ_daily_trials.Form.filter, Me.trials_lblTitle.Caption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trials_upcoming_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "tblPartTrials.trialDate < Date() + 7 AND tblDropDownsSP.trialStatus = 'scheduled'"

If Nz(Me.trials_fltUnit) <> "" Then
    filt = filt & " AND unitId = " & Me.trials_fltUnit
End If

Call trials_applyFilter(filt, "Upcoming (7 days) Trials")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trials_yesterday_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "tblPartTrials.trialDate = Date() -1"

If Nz(Me.trials_fltUnit) <> "" Then
    filt = filt & " AND unitId = " & Me.trials_fltUnit
End If

Call trials_applyFilter(filt, "Yesterday's Trials")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
