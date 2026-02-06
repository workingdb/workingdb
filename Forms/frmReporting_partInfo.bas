Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.sfrmReporting_partInfo_outsource.Form.filter = ""
Me.sfrmReporting_partInfo_outsource.Form.FilterOn = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub outsource_export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Reporting_partInfo_outsource_" & nowString & ".xlsx"
filt = " WHERE " & Me.sfrmReporting_partInfo_outsource.Form.filter
If Me.sfrmReporting_partInfo_outsource.Form.FilterOn = False Then filt = ""
sqlString = "SELECT * FROM sfrmReporting_partInfo_outsource " & filt
                    
Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
