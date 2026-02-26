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

Public Function setHelpScreen(topicId As Long)

Form_sfrmHelp_content.setHelpScreen (topicId)
Form_sfrmHelp_navigation.setHelpScreen (topicId)

End Function
