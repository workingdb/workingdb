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

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.filter = "[requestStatus] <> 'Completed'"
Me.FilterOn = True

'are you a champion?
TempVars.Add "printerChampion", Nz(DCount("recordId", "tbl3Dprinters", "printerChampion = " & userData("ID")), 0) > 0

Me.printers.Visible = TempVars!printerChampion = "True"
Me.materials.Visible = TempVars!printerChampion = "True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub materials_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frm3Dprint_materials"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newPrintRequest_Click()
On Error GoTo Err_Handler

TempVars.Add "new3Dreq", "True"
DoCmd.OpenForm "frm3Dprint_requestDetails", , , , acFormAdd

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openDetails_Click()
On Error GoTo Err_Handler

TempVars.Add "new3Dreq", "False"
DoCmd.OpenForm "frm3Dprint_requestDetails", , , "recordId = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub printers_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frm3Dprint_printers"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub showClosedToggle_Click()
On Error GoTo Err_Handler

Dim filt As String

If Me.showClosedToggle.Value Then
        filt = "[requestStatus] = 'Completed'"
    Else
        filt = "[requestStatus] <> 'Completed'"
End If

Me.filter = filt
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
