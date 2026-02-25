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
Dim ctlTopic As Control, ctlSection As Control, ctlBtn As Control
Dim i As Long
Dim xMargin As Long, pageWidth As Long
Dim ctltag As String

xMargin = 100
pageWidth = Me.Width

Set rsTopic = db.OpenRecordset("tblHelpTopics")

xRunning = 200

Do While Not rsTopic.EOF
    If rsTopic!recordId = topicId Then
        ctltag = "btnContrastBorder.L0"
    Else
        ctltag = "btn.L0"
    End If

    i = i + 1
    h = 500
    Set ctlTopic = Me.Controls("btnBack" & i)
    ctlTopic.Caption = rsTopic!topicName
    ctlTopic.Top = xRunning
    ctlTopic.Left = xMargin
    ctlTopic.Height = h
    ctlTopic.FontBold = 1
    ctlTopic.Width = pageWidth - xMargin * 2
    ctlTopic.tag = ctltag
    ctlTopic.BackStyle = 1
    ctlTopic.BorderStyle = 1
    ctlTopic.BorderWidth = 3
    ctlTopic.fontSize = 14
    ctlTopic.BackStyle = 1
    ctlTopic.Alignment = 1
    
    ctlTopic.Visible = True
    
    xRunning = xRunning + h + 100
    
    Set rsSections = db.OpenRecordset("SELECT * from tblHelpSections WHERE helpTopicId = " & rsTopic!recordId)
    
    Do While Not rsSections.EOF
        'set up btn
        i = i + 1
        h = 500
        Set ctlBtn = Me.Controls("btnBack" & i)
        ctlBtn.Caption = rsSections!sectionTitle
        ctlBtn.Top = xRunning
        ctlBtn.Left = xMargin * 5
        ctlBtn.Height = h
        ctlBtn.FontBold = 0
        ctlBtn.Width = pageWidth - xMargin * 6
        ctlBtn.tag = "btn.L0"
        ctlBtn.BackStyle = 1
        ctlBtn.BorderStyle = 1
        ctlBtn.BorderWidth = 3
        ctlBtn.fontSize = 12
        ctlBtn.BackStyle = 1
        ctlBtn.Alignment = 1
        
        ctlBtn.Visible = True
        
        xRunning = xRunning + h + 100
        
        rsSections.MoveNext
    Loop
    
    rsTopic.MoveNext
Loop



If Not rsTopic Is Nothing Then rsTopic.CLOSE
If Not rsSections Is Nothing Then rsSections.CLOSE
Set rsTopic = Nothing
Set rsSections = Nothing
Set rsItems = Nothing
Set db = Nothing

Call setTheme(Me)

End Function

Function resetLabels()

Dim i, ctrl As Control
For i = 0 To 100
    Set ctrl = Me.Controls("btnBack" & i)
    ctrl.tag = ""
    ctrl.Visible = False
Next i

End Function
