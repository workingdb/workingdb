Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub nmqMorningMeeting_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmReporting_NMQ_daily"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub pePartInfo_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmReporting_partInfo"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub peProjInfo_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmReporting_PE_proj"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
