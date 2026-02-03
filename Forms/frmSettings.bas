Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim dbLoc, fso

Function checkIfAdminDev() As Boolean

checkIfAdminDev = False

Dim errorTxt As String: errorTxt = ""
If (privilege("admin") = False) Then errorTxt = "You need admin privilege to do this"
If (privilege("developer") = False) Then errorTxt = "You need developer privilege to do this"

If errorTxt <> "" Then
    MsgBox errorTxt, vbCritical, "Access Denied"
    checkIfAdminDev = False
    Exit Function
End If

checkIfAdminDev = True

End Function

Private Sub disShift_Click()
On Error GoTo Err_Handler
If Not checkIfAdminDev Then Exit Sub
ap_DisableShift
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub enableShift_Click()
On Error GoTo Err_Handler
If Not checkIfAdminDev Then Exit Sub
ap_EnableShift
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub hideNav_Click()
On Error GoTo Err_Handler
Call DoCmd.NavigateTo("acNavigationCategoryObjectType")
Call DoCmd.RunCommand(acCmdWindowHide)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub hideRibbon_Click()
On Error GoTo Err_Handler

DoCmd.ShowToolbar "Ribbon", acToolbarNo

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openSettings_Click()
On Error GoTo Err_Handler

openPath ("\\data\mdbdata\WorkingDB\Batch\Working DB SETTINGS.lnk")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub sendNewUserInvitation_Click()
On Error GoTo Err_Handler

Dim body As String, primaryMessage As String

Dim tblHeading As String, tblFooter As String, strHTMLBody As String


primaryMessage = "<a href = '\\data\mdbdata\WorkingDB\Batch\Shortcut\Working DB.lnk'>Click Here to Open WorkingDB</a>"

tblHeading = "<table style=""width: 100%; margin: 0 auto; padding: 2em 3em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & "WorkingDB Invitation" & "</h2></td></tr>" & _
                                "<tr><td><p style=""color: rgb(73, 73, 73);"">" & "Once you open workingDB,<br/>please reply to this email so your permissions can be set.<br/>Please open inside of VMWare" & _
                                "</p></td></tr>" & _
                                 "<tr><td><table style=""padding: 1em; text-align: center;"">" & _
                                     "<tr><td style=""padding: 1em 1.5em; background: #FF6B00; "">" & primaryMessage & "</td></tr>" & _
                                "</table></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
tblFooter = "<table style=""width: 100%; margin: 0 auto; padding: 3em; background: #2b2b2b; color: rgba(255,255,255,.5);"">" & _
                        "<tbody>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: 1em; color: #c9c9c9;"">Details</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;""><a href = 'https://nifcoam.sharepoint.com/:p:/r/sites/WorkingDB/Work Instructions/First Time User Overview.pptx?d=wde9823f7cab448a08f1e8f9ab8ba0a7e&csf=1&web=1&e=YmLuWG'>First Time User Overview</a></td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em 1em 2em;""><a href = 'https://nifcoam.sharepoint.com/:p:/r/sites/WorkingDB/Work Instructions/Full Application.pptx?d=w2761e00c353245808c717e4d443f0884&csf=1&web=1&e=KV6VDq'>Full Application Work Instructions</a></td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & "The shortcut should auto-copy to your desktop" & "</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & "Sent by: " & getFullName & "</td></tr>" & _
                        "</tbody>" & _
                    "</table>"

strHTMLBody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 15px; line-height: 1.8;"">" & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center; "">" & _
            "<tbody>" & _
                "<tr><td>" & tblHeading & "</td></tr>" & _
                "<tr><td>" & tblFooter & "</td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"


Call wdbEmail("", "brownj@us.nifco.com", "WorkingDB Invitation", strHTMLBody)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub showRibbon_Click()
On Error GoTo Err_Handler

DoCmd.ShowToolbar "Ribbon", acToolbarYes

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub showNav_Click()
On Error GoTo Err_Handler

Call DoCmd.SelectObject(acTable, , True)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim dev As Boolean
dev = privilege("Developer") 'dev dashboard

Me.enableShift.Visible = dev
Me.disShift.Visible = dev
Me.showNav.Visible = dev
Me.showRibbon.Visible = dev
Me.hideNav.Visible = dev
Me.hideRibbon.Visible = dev

If privilege("Edit") Then
    Me.openSettings.Enabled = True
    Me.openSettings.Caption = " Open Settings App"
Else
    Me.openSettings.Enabled = False
    Me.openSettings.Caption = " Open Settings App - NEED EDIT PRIVILEGE"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub plmSettings_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPLMsettings"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
