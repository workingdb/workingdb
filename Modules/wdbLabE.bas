Option Compare Database
Option Explicit

Public Function registerLabUpdates(table As String, ID As Variant, column As String, _
    oldVal As Variant, newVal As Variant, referenceId As String, _
    formName As String, Optional tag0 As Variant = "")
On Error GoTo Err_Handler

Dim sqlColumns As String, sqlValues As String

If (VarType(oldVal) = vbDate) Then oldVal = Format(oldVal, "mm/dd/yyyy")
If (VarType(newVal) = vbDate) Then newVal = Format(newVal, "mm/dd/yyyy")

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("tbllab_updatetracking")

If Len(oldVal) > 255 Then oldVal = Left(oldVal, 255)
If Len(newVal) > 255 Then newVal = Left(newVal, 255)
If Len(tag0) > 255 Then newVal = Left(tag0, 255)
If ID = "" Then ID = Null

With rs1
    .addNew
        !tableName = table
        !tableRecordId = ID
        !updatedBy = Environ("username")
        !updatedDate = Now()
        !columnName = column
        !previousData = StrQuoteReplace(CStr(Nz(oldVal, "")))
        !newData = StrQuoteReplace(CStr(Nz(newVal, "")))
        !referenceId = referenceId
        !formName = StrQuoteReplace(formName)
        !dataTag0 = StrQuoteReplace(tag0)
    .Update
    .Bookmark = .lastModified
End With

rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbProjectE", "registerPartUpdates", Err.DESCRIPTION, Err.number)
End Function