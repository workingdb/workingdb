Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim dev As Boolean
dev = privilege("Developer") 'allow all
dev = False
Me.setf.SetFocus

Me.tabCtrl.Pages("tabNMQ").Enabled = Not dev
Me.tabCtrl.Pages("tabProject").Enabled = Not dev
Me.tabCtrl.Pages("tabSupplierQuality").Enabled = Not dev

Select Case userData("dept")
    Case "New Molder Quality"
        'this also requires supervisor
        Me.tabCtrl.Pages("tabNMQ").Enabled = Not restrict(Environ("username"), "New Model Quality", "Supervisor", True)
    Case "Project"
        'this also requires supervisor
        Me.tabCtrl.Pages("tabProject").Enabled = Not restrict(Environ("username"), "Project", "Supervisor", True)
    Case "Supplier Quality"
        'this does NOT require supervisor
        Me.tabCtrl.Pages("tabSupplierQuality").Enabled = True
End Select

'set messages
If Not Me.tabCtrl.Pages("tabNMQ").Enabled Then Me.tabCtrl.Pages("tabNMQ").Caption = " DISABLED: Must be NMQ Supervisor"
If Not Me.tabCtrl.Pages("tabProject").Enabled Then Me.tabCtrl.Pages("tabProject").Caption = " DISABLED: Must be Project Supervisor"
If Not Me.tabCtrl.Pages("tabSupplierQuality").Enabled Then Me.tabCtrl.Pages("tabSupplierQuality").Caption = " DISABLED: Must be SQ"

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

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
