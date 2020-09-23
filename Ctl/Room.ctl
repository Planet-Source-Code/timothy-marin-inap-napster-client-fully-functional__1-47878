VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.UserControl Room 
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8910
   ScaleHeight     =   4335
   ScaleWidth      =   8910
   Begin MSScriptControlCtl.ScriptControl Script 
      Left            =   8160
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton cmdScrip 
      Caption         =   "Script"
      Height          =   315
      Left            =   6720
      TabIndex        =   8
      Top             =   0
      Width           =   1005
   End
   Begin VB.PictureBox picBlack 
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   0
      ScaleHeight     =   3915
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   0
      Width           =   75
   End
   Begin VB.PictureBox picSlide 
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   5520
      ScaleHeight     =   3915
      ScaleWidth      =   75
      TabIndex        =   6
      Top             =   315
      Width           =   75
   End
   Begin VB.CommandButton cmdLEave 
      Caption         =   "Leave"
      Height          =   315
      Left            =   7785
      TabIndex        =   5
      Top             =   0
      Width           =   1005
   End
   Begin VB.Timer tmrUsers 
      Interval        =   1000
      Left            =   8280
      Top             =   2760
   End
   Begin VB.TextBox Topic 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   10
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Topic"
      Top             =   0
      Width           =   6645
   End
   Begin VB.TextBox txtSend 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10
      TabIndex        =   1
      Top             =   3915
      Width           =   5520
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5575
      TabIndex        =   0
      Text            =   "0 Users in channel."
      Top             =   3915
      Width           =   3225
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   3615
      Left            =   5565
      TabIndex        =   2
      Top             =   300
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgSpeed"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Shared"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Speed"
         Object.Width           =   1940
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   3615
      Left            =   0
      TabIndex        =   3
      Top             =   300
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6376
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Room.ctx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSpeed 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Room.ctx":0077
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Room.ctx":0611
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Room.ctx":0BAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Room.ctx":1145
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuChan 
      Caption         =   "mnuP"
      Visible         =   0   'False
      Begin VB.Menu mnuPM 
         Caption         =   "Private Message"
      End
      Begin VB.Menu mnuATHL 
         Caption         =   "Add To Hotlist"
      End
      Begin VB.Menu mnuBrowse 
         Caption         =   "Browse Files"
      End
      Begin VB.Menu mnuIgnore 
         Caption         =   "Ignore"
      End
      Begin VB.Menu mnudash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPing 
         Caption         =   "Ping"
      End
      Begin VB.Menu mnuWhoIs 
         Caption         =   "Who Is"
      End
   End
   Begin VB.Menu mnuS 
      Caption         =   "mnuS"
      Begin VB.Menu mnuLS 
         Caption         =   "Load"
      End
   End
End
Attribute VB_Name = "Room"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_CServer = 0
Const m_def_CChannel = 0
'Property Variables:
Dim m_CServer As Variant
Dim m_CChannel As Variant
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Topic,Topic,-1,Text
Public Property Get CTopic() As String
Attribute CTopic.VB_Description = "Returns/sets the text contained in the control."
    CTopic = Topic.Text
End Property

Sub ClearRoom()
    RTB.Text = Time
End Sub
Public Property Let CTopic(ByVal New_CTopic As String)
    Topic.Text() = New_CTopic
    PropertyChanged "CTopic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CServer() As Variant
    CServer = m_CServer
End Property

Public Property Let CServer(ByVal New_CServer As Variant)
    m_CServer = New_CServer
    PropertyChanged "CServer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CChannel() As Variant
    CChannel = m_CChannel
End Property

Public Property Let CChannel(ByVal New_CChannel As Variant)
    m_CChannel = New_CChannel
    PropertyChanged "CChannel"
End Property



Private Sub cmdLEave_Click()
    LeaveRoom
End Sub

Private Sub cmdScrip_Click()
    On Error Resume Next
    frmMain.CD.Filter = "Script File(*.txt)|*.txt"
    frmMain.CD.ShowOpen
    Dim TS As String
    If Len(frmMain.CD.FileName) > 0 Then
        Dim FF As Integer
        Dim LI As String
        FF = FreeFile
        Open frmMain.CD.FileName For Input As #FF
            Line Input #FF, LI
            TS = LI
            Do Until EOF(FF)
                DoEvents
                Line Input #FF, LI
                TS = TS & vbCrLf & LI
            Loop
        Close #FF
    On Error GoTo Err
    ResetScript
    Script.AddCode TS
    Script.Run "SetChan", CChannel, CServer
    Exit Sub
Err:
    MsgBox Err.Description
    Script.Reset
    End If

End Sub

Private Sub LV1_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub LV1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If LV1.SortOrder = lvwAscending Then
        LV1.SortOrder = lvwDescending
    Else
        LV1.SortOrder = lvwAscending
    End If
    LV1.SortKey = ColumnHeader.Index - 1
    LV1.Sorted = True
    LV1.Sorted = False
End Sub

Private Sub LV1_DblClick()
On Error Resume Next
            Fchan = PMtoChan(CStr(LV1.SelectedItem), CServer)
            If Fchan > 50 Then
                Fchan = OpenPM
                If Fchan < 51 Then
                    PM(Fchan).Show
                    PM(Fchan).CServer = CServer
                    PM(Fchan).CName = LV1.SelectedItem
                    PM(Fchan).Caption = "(" & LV1.SelectedItem & ") Instant Message"
                End If
            Else
                    PM(Fchan).Caption = "(" & LV1.SelectedItem & ") Instant Message"
                    PM(Fchan).WindowState = vbNormal
                    PM(Fchan).ZOrder 0
                    
            End If
End Sub

Private Sub LV1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuChan
End Sub

Private Sub mnuATHL_Click()
    'add a user to hotlist
    frmMain.LVNC.ListItems.Add , LV1.SelectedItem, LV1.SelectedItem, , 6
    For i = 1 To frmMain.Sock.UBound
        If frmMain.Sock(i).State = 7 Then frmMain.SendData LV1.SelectedItem, NumToChr(207), Int(i)
    Next
End Sub

Private Sub mnuBrowse_Click()
    '211 nick
    On Error Resume Next
    frmMain.LVB.ListItems.Clear
    frmMain.SendData LV1.SelectedItem, NumToChr(211), CServer
End Sub

Private Sub mnuIgnore_Click()
    On Error Resume Next
    frmMain.lstIgnore.AddItem LV1.SelectedItem
    For i = 0 To frmMain.Sock.UBound
        If frmMain.Sock(i).State = 7 Then
            frmMain.SendData LV1.SelectedItem, NumToChr(322), Int(i)
        End If
    Next
End Sub

Private Sub mnuPing_Click()
    On Error Resume Next
    'frmMain.mnuPrivate_Click
    IsCommand "/ping " & LV1.SelectedItem, CChannel, CServer
End Sub

Private Sub mnuPM_Click()
    On Error Resume Next
    LV1_DblClick
End Sub

Private Sub mnuWhoIs_Click()
    On Error Resume Next
    'frmMain.mnuPrivate_Click
    IsCommand "/whois " & LV1.SelectedItem, CChannel, CServer
End Sub

Private Sub picBlack_Click()
    picSlide_MouseUp 1, 0, 0, 0
End Sub

Private Sub picBlack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picSlide_MouseUp Button, Shift, X, Y
End Sub

Private Sub picSlide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        UserControl.MousePointer = 9
        picBlack.ZOrder 0
        'picSlide.ZOrder 0
        picBlack.Visible = True
        'picSlide.BackColor = vbBlue
        picBlack.BackColor = vbBlack
        'frmMain.Caption = X + picSlide.Left
        If X + picSlide.Left > 0 Then
            If X + picSlide.Left > UserControl.Width Then
                picBlack.Left = UserControl.Width - picBlack.Width
            Else
                picBlack.Left = X + picSlide.Left
            End If
        Else
            picBlack.Left = 0
        End If
                
    End If
End Sub

Private Sub picSlide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        DoEvents
        If X + picSlide.Left > 0 Then
            If X + picSlide.Left > UserControl.Width Then
                picSlide.Left = UserControl.Width - picBlack.Width
            Else
                picSlide.Left = X + picSlide.Left
            End If
        Else
            picSlide.Left = 0
        End If
        UserControl.MousePointer = 0
        LV1.Left = picSlide.Left + picSlide.Width
        RTB.Width = picSlide.Left
        LV1.Width = UserControl.Width - LV1.Left
        txtAmt.Width = LV1.Width
        picSlide.ZOrder 0
        picBlack.Visible = False
        picSlide.BackColor = &H8000000F
        UserControl_Resize
End Sub

Private Sub tmrUsers_Timer()
txtAmt.Text = LV1.ListItems.Count & " Users in channel."
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_CServer = m_def_CServer
    m_CChannel = m_def_CChannel
    ClearRoom
End Sub

Public Sub ChanMsg(User As String, msg As String)
On Error GoTo Err:
If Script.Procedures.Count > 0 Then
    Script.Run "MMsg", User, msg
End If
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, "Script Error"
End Sub

Public Sub JoinMsg(User As String)
On Error GoTo Err:
If Script.Procedures.Count > 0 Then
    Script.Run "Join", User
End If
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, "Script Error"
End Sub

Public Sub PartMsg(User As String)
On Error GoTo Err:
If Script.Procedures.Count > 0 Then
    Script.Run "Part", User
End If
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical, "Script Error"
End Sub

Public Sub ResetScript()
    Script.Reset
    Script.AddObject "frmMain", frmMain, True
    'Script.AddObject "Room", UserControl, True
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Topic.Text = PropBag.ReadProperty("CTopic", "Topic")
    m_CServer = PropBag.ReadProperty("CServer", m_def_CServer)
    m_CChannel = PropBag.ReadProperty("CChannel", m_def_CChannel)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Topic.Width = UserControl.Width + 10 - 50 - cmdLEave.Width - cmdScrip.Width
    cmdLEave.Left = UserControl.Width - cmdLEave.Width
    cmdScrip.Left = cmdLEave.Left - cmdScrip.Width - 20
    RTB.Width = Topic.Width + cmdLEave.Width + cmdScrip.Width - LV1.Width
    LV1.Left = RTB.Width + 20
    txtSend.TOP = UserControl.Height - txtSend.Height
    txtSend.Width = RTB.Width - 30
    txtAmt.Left = LV1.Left + 20
    txtAmt.TOP = txtSend.TOP
    RTB.Height = UserControl.Height - RTB.TOP - txtAmt.Height - 20
    LV1.Height = RTB.Height
    picSlide.Left = RTB.Width - 10
    picSlide.Width = 40
    picSlide.TOP = Topic.TOP + Topic.Height
    picSlide.Height = UserControl.Height - Topic.Height
    picBlack.Left = RTB.Width - 10
    picBlack.Width = 40
    picBlack.TOP = Topic.TOP + Topic.Height
    picBlack.Height = UserControl.Height - Topic.Height
    picSlide.ZOrder 0
    SizeBy LV1, 1
    UserControl.MousePointer = 0
End Sub

Sub LeaveRoom()
    Dim X As Integer
    frmMain.SendData CChannel, Chr(145) & Chr(1), CServer
    X = ChantoChan(CChannel, CServer)
    LV1.ListItems.Clear
    RTB.Text = ""
    txtSend.Text = ""
    frmMain.mnuRoom(X).Visible = False
    frmMain.Room(X).Visible = False
    frmMain.Room(X).CServer = ""
    frmMain.Room(X).CChannel = ""
    frmMain.Room(X).CTopic = ""
    frmMain.mnuPrivate_Click
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CTopic", Topic.Text, "Topic")
    Call PropBag.WriteProperty("CServer", m_CServer, m_def_CServer)
    Call PropBag.WriteProperty("CChannel", m_CChannel, m_def_CChannel)
End Sub

Sub AddText(msg As String, Vcol As ColorConstants)
    RTB.SelStart = Len(RTB.Text)
    RTB.SelColor = Vcol
    RTB.SelText = vbCrLf & msg
End Sub

Sub AddUser(Nick As String, Files As String, Speed As String, Ico As String)
    On Error Resume Next
    LV1.ListItems.Add , , Nick, , Int(Ico)
    LV1.ListItems.Item(LV1.ListItems.Count).SubItems(1) = Files
    LV1.ListItems.Item(LV1.ListItems.Count).SubItems(2) = Speed
End Sub

Sub RemoveUser(Nick As String)
    On Error Resume Next
    For Y = 1 To LV1.ListItems.Count
        If LV1.ListItems.Item(Y) = Nick Then
            LV1.ListItems.Remove Y
        End If
    Next
End Sub

Private Sub RTB_Change()
    RTB.SelStart = Len(RTB.Text)
End Sub

'    <channel> <message>
Private Sub txtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsCommand(txtSend.Text, CChannel, CServer) = False Then
            frmMain.SendData CChannel & " " & txtSend.Text, Chr(146) & Chr(1), CServer
        End If
        txtSend.Text = ""
    End If
End Sub
