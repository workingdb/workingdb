Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub attachHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblStratPlanAttachmentsSP' AND referenceId = " & Me.referenceId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Detail_Paint()
On Error Resume Next

Me.openAttachment.Transparent = Me.fileStatus <> "Uploaded"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
If CurrentProject.AllForms("frmCapacityRequestDetails").IsLoaded Then Form_frmCapacityRequestDetails.fileCount.Requery
If CurrentProject.AllForms("frmDropFile").IsLoaded Then DoCmd.CLOSE acForm, "frmDropFile"

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgUploadedBy_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.uploadedBy & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newAttachment_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmDropFile"
'custom title stuff
Form_frmDropFile.customName.Visible = True
Form_frmDropFile.customName.Locked = False
Form_frmDropFile.TdocumentLibary = Me.TdocLibrary
Form_frmDropFile.Label62.Visible = True
Form_frmDropFile.Command63.Visible = True

Form_frmDropFile.lblDocCategory.Caption = "Strategic Planning Document"
Form_frmDropFile.TprojectId = Me.TreferenceId
Form_frmDropFile.TpartNumber = Me.TreferenceTable
Form_frmDropFile.documentType = 30
Form_frmDropFile.documentType.Locked = True
Form_frmDropFile.documentType.Visible = False
Form_frmDropFile.docTypeCard.Visible = False
Form_frmDropFile.Label51.Visible = False
Form_frmDropFile.Box57.Visible = False
Form_frmDropFile.Command58.Visible = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openAttachment_Click()
On Error GoTo Err_Handler

If Me.fileStatus = "Uploaded" Then
    Application.FollowHyperlink Me.directLink
    Call registerStratPlanUpdates("tblStratPlanAttachmentsSP", Me.ID, "File Attachment", Me.attachName, "Opened", Me.referenceId, Me.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If Me.fileStatus <> "Uploaded" Then
    MsgBox "File must be fully uploaded in order to delete.", vbInformation, "Wait a second..."
    Exit Sub
End If

If MsgBox("Are you sure?", vbYesNo, "Please confirm") <> vbYes Then Exit Sub

Call registerStratPlanUpdates("tblStratPlanAttachmentsSP", Me.ID, "File Attachment", Me.attachName, "Deleted", Me.referenceId, Me.name)

Me.fileStatus = "Deleting"
If Me.Dirty Then Me.Dirty = False

Me.Requery
Me.refresh

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
