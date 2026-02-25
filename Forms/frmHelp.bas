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

Dim db As Database
Set db = CurrentDb()

Dim rsTopic As Recordset
Dim rsSections As Recordset
Dim rsItems As Recordset

Dim h As Long, lineHeight As Long
Dim xRunning As Long
Dim ctlTopic As Control, ctlSection As Control, ctlItem As Control, ctlBtn As Control
Dim i As Long
Dim xMargin As Long, pageWidth As Long
xMargin = 2300
pageWidth = 15840

Set rsTopic = db.OpenRecordset("SELECT * from tblHelpTopics WHERE recordId = " & topicId)

xRunning = 200
lineHeight = 500

Set ctlTopic = Me.Controls("lbl" & i)
ctlTopic.Caption = rsTopic!topicName
ctlTopic.Top = xRunning
ctlTopic.Left = 0
ctlTopic.fontSize = 24
ctlTopic.FontBold = 1
ctlTopic.Height = 700
ctlTopic.Width = pageWidth
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
    ctlSection.Left = xMargin
    ctlSection.TopMargin = 50
    ctlSection.fontSize = 14
    ctlSection.FontBold = 1
    ctlSection.Height = 500
    ctlSection.Width = pageWidth - xMargin * 2
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
        ctlItem.Left = xMargin
        ctlItem.fontSize = 12
        ctlItem.FontBold = 0
        ctlItem.Height = h
        ctlItem.Width = pageWidth - xMargin * 2
        ctlItem.TopMargin = 100
        ctlItem.LeftMargin = 100
        ctlItem.RightMargin = 100
        ctlItem.tag = "lbl_wBack.L1"
        ctlItem.BackStyle = 1
        ctlItem.TextAlign = 2
        ctlItem.Visible = True
        
        xRunning = xRunning + h
        
        If Nz(rsItems!helpContentImage, "") <> "" Then
            h = 5000
            Set ctlItem = Me.Controls("pic" & i)
            ctlItem.Picture = "\\data\mdbdata\WorkingDB\Pictures\help\" & rsItems!helpContentImage & ".png"
            ctlItem.Top = xRunning
            ctlItem.Left = xMargin
            ctlItem.Height = h
            ctlItem.Width = pageWidth - xMargin * 2
            ctlItem.tag = "pic.L1"
            ctlItem.BackStyle = 1
            ctlItem.Visible = True
            
            xRunning = xRunning + h
        End If
        
        rsItems.MoveNext
    Loop
    
    'set up cardBtn
    i = i + 1
    Set ctlBtn = Me.Controls("btnBack" & i)
    ctlBtn.Top = (ctlSection.Top) - 25
    ctlBtn.Left = (xMargin) - 25
    ctlBtn.Height = (xRunning - ctlSection.Top) + 50
    ctlBtn.Width = (pageWidth - xMargin * 2) + 50
    ctlBtn.tag = "cardBtn.L1"
    ctlBtn.BackStyle = 0
    ctlBtn.BorderStyle = 1
    ctlBtn.BorderWidth = 3
    
    ctlBtn.Visible = True
    
    xRunning = xRunning + 300
    
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
