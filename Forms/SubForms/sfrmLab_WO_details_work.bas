Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub addWork_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

db.Execute "INSERT INTO tbllab_wo_work(woid) VALUES (" & Form_frmLab_WO_details.recordId & ");"
TempVars.Add "workoId", db.OpenRecordset("SELECT @@identity")(0).Value

Set db = Nothing

Call registerLabUpdates("tbllab_wo_work", TempVars!workoId, "WO Work", "", "Created", Form_frmLab_WO_details.recordId, Me.name)
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub deleteItem_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
   Call registerLabUpdates("tbllab_wo_work", Me.recordId, "WO Work", "", "Deleted", Form_frmLab_WO_details.recordId, Me.name)
    dbExecute "DELETE from tbllab_wo_work WHERE recordid = " & Me.recordId
    Me.Requery
    Call snackBox("success", "Success!", "Work deleted", Me.Parent.name)
End If


Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub esthours_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_work", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub itemCount_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_work", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub perCavity_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_work", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmLab_WO_details.recordId, Me.name)

Me.sampleCount = Nz(Form_frmLab_WO_details.Cavities) * Nz(Me.perCavity)
Call registerLabUpdates("tbllab_wo_work", Me.recordId, Me.sampleCount.name, Me.sampleCount.OldValue, Me.sampleCount, Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub reportFormat_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_work", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub retain_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_work", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub sampleCount_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_work", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub worktype_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_work", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
