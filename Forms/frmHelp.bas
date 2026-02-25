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

Public Function setHelpScreen()

Dim db As Database
Set db = CurrentDb()

Dim rsTopic As Recordset
Dim rsSections As Recordset
Dim rsItems As Recordset

Dim h As Long, lineHeight As Long
Dim xRunning As Long
Dim ctlTopic As Control, ctlSection As Control, ctlItem As Control
Dim i As Long

Set rsTopic = db.OpenRecordset("SELECT * from tblHelpTopics WHERE recordId = " & 1)

xRunning = 100
lineHeight = 500

Set ctlTopic = Me.Controls("lbl" & i)
ctlTopic.Caption = rsTopic!topicName
ctlTopic.Top = xRunning
ctlTopic.Left = 0
ctlTopic.fontSize = 24
ctlTopic.FontBold = 1
ctlTopic.Height = 700
ctlTopic.Width = 15840
ctlTopic.tag = "lbl.L0"
ctlTopic.TextAlign = 2
ctlTopic.Visible = True

xRunning = xRunning + 1000

Set rsSections = db.OpenRecordset("SELECT * from tblHelpSections WHERE helpTopicId = " & rsTopic!recordId)

Do While Not rsSections.EOF
    i = i + 1
    Set ctlSection = Me.Controls("lbl" & i)
    ctlSection.Caption = rsSections!sectionTitle
    ctlSection.Top = xRunning
    ctlSection.Left = 0
    ctlSection.fontSize = 14
    ctlSection.FontBold = 1
    ctlSection.Height = 500
    ctlSection.Width = 15840
    ctlSection.tag = "lbl_wBack.L1"
    ctlSection.BackStyle = 1
    ctlSection.TextAlign = 2
    ctlSection.Visible = True
    
    xRunning = xRunning + 500

    Set rsItems = db.OpenRecordset("SELECT * from tblHelpItems WHERE helpSectionId = " & rsSections!recordId)
    
    Do While Not rsItems.EOF
        i = i + 1
        
        Set ctlItem = Me.Controls("lbl" & i)
        
        h = -Int(-Len(rsItems!helpContentText) / 106) * lineHeight
        
        ctlItem.Caption = rsItems!helpContentText
        ctlItem.Top = xRunning
        ctlItem.Left = 0
        ctlItem.fontSize = 12
        ctlItem.FontBold = 0
        ctlItem.Height = h
        ctlItem.Width = 15840
        ctlItem.tag = "lbl_wBack.L1"
        ctlItem.BackStyle = 1
        ctlItem.TextAlign = 2
        ctlItem.LeftMargin = 2500
        ctlItem.RightMargin = 2500
        ctlItem.Visible = True
        
        xRunning = xRunning + h
        
        If Nz(rsItems!helpContentImage, "") <> "" Then
            Set ctlItem = Me.Controls("pic" & i)
            ctlItem.Picture = "\\data\mdbdata\WorkingDB\Pictures\help\" & rsItems!helpContentImage & ".png"
            ctlItem.Top = xRunning
            ctlItem.Left = 0
            ctlItem.Height = 5000
            ctlItem.Width = 15840
            ctlItem.tag = "pic.L0"
            ctlItem.Visible = True
            
            xRunning = xRunning + 5000 + 300
        End If
        
        rsItems.MoveNext
    Loop
    xRunning = xRunning + 200
    
    rsSections.MoveNext
Loop



If Not rsTopic Is Nothing Then rsTopic.CLOSE
If Not rsSections Is Nothing Then rsSections.CLOSE
If Not rsItems Is Nothing Then rsItems.CLOSE
Set rsTopic = Nothing
Set rsSections = Nothing
Set rsItems = Nothing
Set db = Nothing

Call setTheme(Me)

End Function

Function resetLabels()

Dim i, ctrl As Control
For i = 0 To 100
    Set ctrl = Form_frmHelp.Controls("lbl" & i)
    ctrl.tag = ""
    ctrl.Visible = False
    Set ctrl = Form_frmHelp.Controls("pic" & i)
    ctrl.tag = ""
    ctrl.Visible = False
Next i

End Function
