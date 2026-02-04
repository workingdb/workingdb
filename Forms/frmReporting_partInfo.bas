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
