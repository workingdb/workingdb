Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub closedSteps_export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Reporting_PE_proj_closedSteps_" & nowString & ".xlsx"
filt = " WHERE " & Me.sfrmReporting_PE_proj_closedSteps.Form.filter
If Me.sfrmReporting_PE_proj_closedSteps.Form.FilterOn = False Then filt = ""
sqlString = "SELECT * FROM sfrmReporting_PE_proj_closedSteps " & filt
                    
Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.sfrmReporting_PE_proj_lateSteps.Form.filter = ""
Me.sfrmReporting_PE_proj_lateSteps.Form.FilterOn = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub lateSteps_export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Reporting_PE_proj_lateSteps_" & nowString & ".xlsx"
filt = " WHERE " & Me.sfrmReporting_PE_proj_lateSteps.Form.filter
If Me.sfrmReporting_PE_proj_lateSteps.Form.FilterOn = False Then filt = ""
sqlString = "SELECT * FROM sfrmReporting_PE_proj_lateSteps " & filt
                    
Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
