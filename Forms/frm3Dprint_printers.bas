Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function filterIt(controlName As String)
On Error GoTo Err_Handler

Me(controlName).SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Function
Err_Handler:
    Call handleError(Me.name, "filterIt", Err.DESCRIPTION, Err.number)
End Function

Private Sub allHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryDRSupdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.previousData.ControlSource = "previous"
Form_frmHistory.newData.ControlSource = "new"
Form_frmHistory.filter = "tableName = 'tbl3Dprinters'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub printerChampion_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3Dprinters", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub printerItemHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryDRSupdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.previousData.ControlSource = "previous"
Form_frmHistory.newData.ControlSource = "new"
Form_frmHistory.filter = "tableRecordId = " & Me.recordId & " AND tableName = 'tbl3Dprinters'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub printerName_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3Dmaterials", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub printerType_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3Dprinters", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
    Call registerDRSUpdates("tbl3Dprinters", Me.recordId, "Request", "", "Deleted", Me.name)
    dbExecute ("DELETE FROM tbl3Dprinters WHERE [recordId] = " & Me.recordId)
    Me.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
