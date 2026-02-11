Option Compare Database
Option Explicit

Public Function registerStratPlanUpdates( _
    table As String, _
    ID As Variant, _
    column As String, _
    oldVal As Variant, _
    newVal As Variant, _
    referenceId As Long, _
    formName As String, _
    Optional tag0 As String = "", _
    Optional customUpdatedBy As String = "")
On Error GoTo Err_Handler

If (VarType(oldVal) = vbDate) Then oldVal = Format(oldVal, "mm/dd/yyyy")
If (VarType(newVal) = vbDate) Then newVal = Format(newVal, "mm/dd/yyyy")

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("tblStratPlan_UpdateTracking")

Dim updatedBy As String
updatedBy = Environ("username")
If customUpdatedBy <> "" Then updatedBy = customUpdatedBy

If Len(oldVal) > 255 Then oldVal = Left(oldVal, 255)
If Len(newVal) > 255 Then newVal = Left(newVal, 255)
If Len(tag0) > 100 Then tag0 = Left(tag0, 100)
If ID = "" Then ID = Null

With rs1
    .addNew
        !tableName = table
        !tableRecordId = ID
        !updatedBy = updatedBy
        !updatedDate = Now()
        !columnName = column
        !previousData = StrQuoteReplace(CStr(Nz(oldVal, "")))
        !newData = StrQuoteReplace(CStr(Nz(newVal, "")))
        !referenceId = referenceId
        !formName = formName
        !dataTag0 = StrQuoteReplace(tag0)
    .Update
    .Bookmark = .lastModified
End With

rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbStratPlan", "registerStratPlanUpdates", Err.DESCRIPTION, Err.number)
End Function