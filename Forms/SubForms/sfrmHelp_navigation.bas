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

Function helpTopicClick()

Dim topicId As Long
topicId = Split(Me.ActiveControl.tag, ",")(0)

Call setHelpScreen(topicId)
Call Form_sfrmHelp_content.setHelpScreen(topicId)

End Function

Public Function setHelpScreen(topicId As Long)

resetLabels

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
Dim currentTopic As Boolean

xMargin = 100

pageWidth = Me.Width

Set rsTopic = db.OpenRecordset("SELECT * FROM tblHelpTopics ORDER BY indexOrder")

xRunning = 200

Do While Not rsTopic.EOF
    If rsTopic!recordId = topicId Then
        currentTopic = True
        ctltag = rsTopic!recordId & ",btnContrastBorder.L0"
    Else
        currentTopic = False
        ctltag = rsTopic!recordId & ",btn.L0"
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
    
    xRunning = xRunning + h + 50
    
    If currentTopic Then
        Set rsSections = db.OpenRecordset("SELECT * from tblHelpSections WHERE helpTopicId = " & rsTopic!recordId & " ORDER BY indexOrder")
        
        Do While Not rsSections.EOF
            'set up btn
            i = i + 1
            h = 450
            Set ctlBtn = Me.Controls("btnBack" & i)
            ctlBtn.Caption = rsSections!sectionTitle
            ctlBtn.Top = xRunning
            ctlBtn.Left = xMargin * 5
            ctlBtn.Height = h
            ctlBtn.FontBold = 0
            ctlBtn.Width = pageWidth - xMargin * 7
            ctlBtn.tag = "cardBtn.L0"
            ctlBtn.BackStyle = 1
            ctlBtn.BorderStyle = 1
            ctlBtn.BorderWidth = 3
            ctlBtn.fontSize = 12
            ctlBtn.BackStyle = 1
            ctlBtn.Alignment = 1
            ctlBtn.OnClick = ""
            ctlBtn.CursorOnHover = 2
            
            ctlBtn.Visible = True
            
            xRunning = xRunning + h + 50
            
            rsSections.MoveNext
        Loop
    End If
    
    rsTopic.MoveNext
Loop

Me.Detail.Height = xRunnning + h + 50


If Not rsTopic Is Nothing Then rsTopic.CLOSE
If Not rsSections Is Nothing Then rsSections.CLOSE
Set rsTopic = Nothing
Set rsSections = Nothing
Set rsItems = Nothing
Set db = Nothing

Call setTheme(Me)

End Function

Function resetLabels()

Dim ctrl As Control

Set ctrl = Me.Controls("btnBack0")
ctrl.Caption = ""
ctrl.tag = "btn.L0"
ctrl.Height = 1
ctrl.Width = 1
ctrl.Left = 1
ctrl.Top = 1
ctrl.Visible = True
ctrl.SetFocus

Dim i
For i = 1 To 100
    Set ctrl = Me.Controls("btnBack" & i)
    ctrl.Left = 1
    ctrl.Top = 1
    ctrl.Width = 1
    ctrl.Height = 1
    ctrl.tag = ""
    ctrl.Visible = False
Next i

End Function
