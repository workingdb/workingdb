Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub


Function resetLabels()

Dim i, ctrl
For i = 0 To 2
    Set ctrl = Form_frmHelp.Controls("lbl" & i)
    ctrl.Visible = True
Next i

End Function
