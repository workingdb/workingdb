Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub allHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryDRSupdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.previousData.ControlSource = "previous"
Form_frmHistory.newData.ControlSource = "new"
Form_frmHistory.filter = "tableRecordId = " & Me.recordId & " AND tableName = 'tbl3DprintRequests'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnEditLocation_Click()
On Error GoTo Err_Handler

Dim strFolder As String
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Choose a Folder"
    .AllowMultiSelect = False
    .Show
    
    On Error Resume Next
    strFolder = .SelectedItems(1)
End With

If Nz(strFolder, "") = "" Then Exit Sub

strFolder = replaceDriveLetters(addLastSlash(strFolder))

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.folderLocation.name, Me.folderLocation, strFolder, Me.name)

Me.folderLocation = strFolder
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnOpenLocation_Click()
On Error GoTo Err_Handler

If FolderExists(Me.folderLocation) Then
    FollowHyperlink Me.folderLocation
Else
    Call snackBox("error", "Hmm...", "Something is wrong with the folder", Me.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cubicVolume_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub dropDeadDate_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.Text, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub folderLocation_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub forDepartment_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

TempVars.Add "req3Ddelete", "False"

Me.createdBy.DefaultValue = DLookup("ID", "tblPermissions", "user = '" & Environ("username") & "'")

Me.allowEdits = True

If TempVars!printerChampion = "True" Then
    Me.requestStatus.Locked = False
    Me.remove.Visible = True
    GoTo permOk
End If

'non champions
Me.requestStatus.Locked = True
Me.remove.Visible = False

'if new request, just lock requestStatus
If Me.requestStatus = 1 And TempVars!new3Dreq = "False" Then 'if not accepted yet, allow a few fields to be edited
    Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.tag Like "*lckReq*" Then ctl.Locked = True
    Next
    Me.materialid.Locked = False
    Me.requestQuantity.Locked = False
ElseIf Me.requestStatus > 1 Then
    Me.allowEdits = False
End If


permOk:

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Function validate() As Boolean

validate = False

If IsNull(Me.recordId) Then
    validate = True
    Exit Function
End If

Dim errorArray As Collection
Set errorArray = New Collection

Dim frm As Form, ctl As Control

Set frm = Me
For Each ctl In frm.Controls
    Select Case ctl.ControlType
        Case acLabel
            If Left(ctl.Parent.name, 3) = "frm" Then GoTo nextLabel
            
            frm.Controls(ctl.Parent.name).SetFocus
            If Right(ctl.Caption, 1) = "*" And Nz(frm.Controls(ctl.Parent.name).Text) = "" Then
                errorArray.Add ctl.Caption
                frm.Controls(ctl.Parent.name).tag = Replace(frm.Controls(ctl.Parent.name).tag, "txt", "txtErr")
            End If
            
nextLabel:
    End Select
Next ctl

If errorArray.count > 0 Then
    Dim errorTxtLines As String, element
    errorTxtLines = ""
    For Each element In errorArray
        errorTxtLines = errorTxtLines & vbNewLine & element
    Next element
    
    Call setTheme(Me)
    
    MsgBox "Please fill out these items: " & vbNewLine & errorTxtLines, vbOKOnly, "No can do!"
    Exit Function
End If

validate = True

End Function

Private Sub btnSave_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If validate = False Then Exit Sub

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, "Request", "", "Saved", Me.name)

If TempVars!new3Dreq = "True" Then
    Dim printerChampion As String, champId As Long
    champId = DLookup("printerChampion", "tbl3Dprinters", "recordId = " & Me.Controls("printer"))
    printerChampion = DLookup("user", "tblPermissions", "ID = " & champId)
    
    Dim body As String
    body = emailContentGen("New Print Request", _
        "New " & Me.requestReason.column(1), _
        "Notes: " & Me.requestNotes, _
         "Title: " & Me.requestTitle & " for: " & Me.createdBy.column(1), _
        "Requested: " & CStr(Date) & ", by: " & Me.createdBy.column(1), _
        "Priority: " & Me.requestPriority.column(1), _
        "Printer: " & Me.Controls("printer").column(1), "3D Print", Me.recordId)
        
    Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, "Submission", "", "New Request", Me.name)
    
    Call sendNotification(printerChampion, 6, 2, "New Print Request", body, "3D Print", Me.recordId)
End If

DoCmd.CLOSE acForm, "frm3Dprint_requestDetails"

If CurrentProject.AllForms("frm3Dprint_requests").IsLoaded Then Form_frm3Dprint_requests.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

If TempVars!req3Ddelete = "True" Then Exit Sub

If Me.Dirty Then Me.Dirty = False

If validate = False Then
    If MsgBox("Are you sure?" & vbNewLine & "Your current record will be deleted.", vbYesNo, "Please confirm") <> vbYes Then
        Cancel = True
        Exit Sub
    End If

    DoCmd.SetWarnings False
    If Nz(Me.recordId) <> "" Then DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
ElseIf Not IsNull(Me.recordId) Then 'passes validation! new record being saved
    Call registerDRSUpdates("tbl3Dprint_Requests", Me.recordId, "Created", "", Me.requestTitle, Me.name)
End If

If CurrentProject.AllForms("frm3Dprint_requests").IsLoaded Then Form_frm3Dprint_requests.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub materialid_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub printer_AfterUpdate()
On Error GoTo Err_Handler

If Me.ActiveControl = 2 Then 'sinterit only
    Me.lblcubicVolume.Caption = "Cubic Volume*"
Else
    Me.lblcubicVolume.Caption = "Cubic Volume"
End If

Me.materialid.RowSource = "SELECT " & _
    "mat.recordId,dd.materialType as Type,dd_1.materialColor as Color,mat.materialQuantity as Spools " & _
    "From (tbl3Dmaterials As mat INNER JOIN tbl3Ddropdowns as dd ON mat.materialType = dd.recordId) " & _
    "INNER JOIN tbl3Ddropdowns AS dd_1 ON mat.materialColor = dd_1.recordId " & _
    "WHERE (mat.materialPrinter = " & Me.ActiveControl & ") " & _
    "ORDER BY mat.materialQuantity DESC;"
    
Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this request?", vbYesNo, "Please confirm") = vbYes Then
    Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, "Request", Me.requestTitle, "Deleted", Me.name)
    dbExecute ("DELETE FROM tbl3DprintRequests WHERE [recordId] = " & Me.recordId)
    TempVars.Add "req3Ddelete", "True"
    DoCmd.CLOSE
    If CurrentProject.AllForms("frm3Dprint_requests").IsLoaded Then Form_frm3Dprint_requests.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub requestNotes_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub requestPriority_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub requestQuantity_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub requestReason_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub requestStatus_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub requestTitle_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tbl3DprintRequests", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
