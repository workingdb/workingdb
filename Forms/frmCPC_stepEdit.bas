Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub deletebtn_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this entire step?", vbYesNo, "Please confirm") <> vbYes Then Exit Sub

Call registerCPCUpdates("tblCPC_Steps", Me.ID, "Step", Me.stepName, "Deleted", Me.projectId, Me.stepName)

Dim db As Database
Set db = CurrentDb()
db.Execute "DELETE FROM tblCPC_StepApprovals WHERE stepId = " & Me.recordId
db.Execute "DELETE FROM tblCPC_Steps WHERE ID = " & Me.recordId
DoCmd.CLOSE
Form_sfrmCPC_Dashboard.Requery

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub moveDown_Click()
On Error GoTo Err_Handler

If Me.indexOrder = DMax("indexOrder", "tblCPC_Steps", "projectId = " & Me.projectId) Then Exit Sub

Dim oldIndex, newIndex
oldIndex = Me.indexOrder
newIndex = oldIndex + 1

dbExecute "UPDATE tblCPC_Steps SET indexOrder = " & oldIndex & " WHERE projectId = " & Me.projectId & " AND indexOrder = " & newIndex

Me.indexOrder = newIndex
Me.Dirty = False

Call registerCPCUpdates("tblCPC_Steps", Me.ID, "indexOrder", oldIndex, newIndex, Me.projectId, Me.stepName)

DoCmd.CLOSE
Form_sfrmCPC_Dashboard.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub moveUp_Click()
On Error GoTo Err_Handler

If Me.indexOrder = 1 Then Exit Sub
Dim oldIndex, newIndex
oldIndex = Me.indexOrder
newIndex = oldIndex - 1

dbExecute "UPDATE tblCPC_Steps SET indexOrder = " & oldIndex & " WHERE projectId = " & Me.projectId & " AND indexOrder = " & newIndex

Me.indexOrder = newIndex
Me.Dirty = False

Call registerCPCUpdates("tblCPC_Steps", Me.ID, "indexOrder", oldIndex, newIndex, Me.projectId, Me.stepName)

DoCmd.CLOSE
Form_sfrmCPC_Dashboard.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
