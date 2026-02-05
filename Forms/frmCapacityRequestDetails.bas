Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnSave_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If validate = False Then Exit Sub

DoCmd.CLOSE acForm, "frmCapacityRequestDetails"

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
            
            If Right(ctl.Caption, 1) = "*" And (Nz(frm.Controls(ctl.Parent.name).Value) = "" Or Nz(frm.Controls(ctl.Parent.name).Value) = 0) Then
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

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

If TempVars!capAdd = "True" Then
    If validate = False Then
        If MsgBox("Are you sure?" & vbNewLine & "Your current record will be deleted.", vbYesNo, "Please confirm") <> vbYes Then
            Cancel = True
            Exit Sub
        End If
    
        DoCmd.SetWarnings False
        If Nz(Me.recordId) <> "" Then DoCmd.RunCommand acCmdDeleteRecord
        DoCmd.SetWarnings True
    Else 'passes validation! new record being saved
        Dim body As String
        body = emailContentGen("New Capacity Request", _
            "New " & Me.requestType.column(1), _
            "Notes: " & Me.Notes, _
             "PN: " & Me.NAM & " on Program: " & Me.Program.column(0), _
            "Requested: " & CStr(Date) & ", by: " & Me.Requestor.column(1), _
            Me.volumeTiming.column(1) & " Volume: " & Me.Volume, _
            "Vehicle: " & Me.Program.column(1))
        Call sendNotification("capacityrequest@us.nifco.com", 6, 2, "New Capacity Request", body, customEmail:=True)
    End If
End If

If CurrentProject.AllForms("frmCapacityRequestTracker").IsLoaded = True Then Form_frmCapacityRequestTracker.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub NAM_AfterUpdate()
On Error GoTo Err_Handler

If Nz(Me.NAM, "") = "" Then Exit Sub

'find current unit
Dim db As Database
Set db = CurrentDb()
Dim invId, currentUnit As String, rsCat As Recordset
invId = Nz(idNAM(Me.NAM, "NAM"), "")

currentUnit = ""

If invId <> "" Then
    Set rsCat = db.OpenRecordset("SELECT SEGMENT1 FROM INV_MTL_ITEM_CATEGORIES LEFT JOIN APPS_MTL_CATEGORIES_VL ON INV_MTL_ITEM_CATEGORIES.CATEGORY_ID = APPS_MTL_CATEGORIES_VL.CATEGORY_ID " & _
    "GROUP BY INV_MTL_ITEM_CATEGORIES.INVENTORY_ITEM_ID, APPS_MTL_CATEGORIES_VL.SEGMENT1, APPS_MTL_CATEGORIES_VL.STRUCTURE_ID HAVING STRUCTURE_ID = 101 AND [INVENTORY_ITEM_ID] = " & invId, dbOpenSnapshot)
    If rsCat.RecordCount > 0 Then currentUnit = Nz(rsCat!SEGMENT1, "")

    rsCat.CLOSE
    Set rsCat = Nothing
End If

If currentUnit <> "" Then
    Dim unitId
    unitId = Nz(DLookup("recordId", "tblUnits", "unitName = '" & currentUnit & "'"), 0)
    Me.unit = unitId
End If

'Dim rs1 As DAO.Recordset
'Set rs1 = db.OpenRecordset("SELECT ITEM_TYPE FROM APPS_MTL_SYSTEM_ITEMS WHERE SEGMENT1 = '" & Me.NAM & "'", dbOpenSnapshot)
'
'Me.ITEM_TYPE = rs1("ITEM_TYPE")
'
'rs1.CLOSE
'Set rs1 = Nothing

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub requestType_AfterUpdate()
On Error GoTo Err_Handler

Me.Requestor = Nz(DLookup("ID", "tblPermissions", "user = '" & Environ("username") & "'"), 0)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
