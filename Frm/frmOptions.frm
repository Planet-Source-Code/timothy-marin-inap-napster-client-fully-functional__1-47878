VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5820
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   315
      Left            =   4560
      TabIndex        =   33
      Top             =   3765
      Width           =   1215
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   635
      ButtonWidth     =   2037
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgTool"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "My Info   "
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Shared    "
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Connect   "
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Transfer   "
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Other     "
            ImageIndex      =   11
            Style           =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imgTool 
      Left            =   6720
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":39FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":5DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":64AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":8890
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":AC72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":D054
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":D5EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":F9D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":FD6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":11A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":11E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":121A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":12542
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmOpt 
      Height          =   3375
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   5775
      Begin VB.CheckBox chkPsv 
         Caption         =   "Passive Mode"
         Height          =   255
         Left            =   3240
         TabIndex        =   34
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbVersion 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOptions.frx":128DC
         Left            =   1320
         List            =   "frmOptions.frx":128E6
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1680
         Width           =   3855
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Show options on startup"
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   2280
         Width           =   4695
      End
      Begin VB.ComboBox cmbSpeed 
         Height          =   315
         ItemData        =   "frmOptions.frx":128FC
         Left            =   1320
         List            =   "frmOptions.frx":1291E
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtOption 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Text            =   "Tim"
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtOption 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Text            =   "Piggy"
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtOption 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   1
         Text            =   "6699"
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Pass :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Port :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Speed :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Version :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame frmOpt 
      Height          =   3375
      Index           =   1
      Left            =   30
      TabIndex        =   14
      Top             =   360
      Width           =   5775
      Begin VB.CommandButton cmdDirRem 
         Caption         =   "Remove"
         Height          =   255
         Left            =   4320
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdDirAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   2655
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   2880
         TabIndex        =   16
         Top             =   600
         Width           =   2775
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblInfo 
         Caption         =   "Subdirectories also shared !"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   20
         Top             =   2880
         Width           =   2655
      End
   End
   Begin VB.Frame frmOpt 
      Height          =   3375
      Index           =   2
      Left            =   30
      TabIndex        =   9
      Top             =   360
      Width           =   5775
      Begin VB.CheckBox chkGator 
         Caption         =   "Also connect to new napigator list (one server per network)"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Value           =   1  'Checked
         Width           =   5535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Remove"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   2655
      End
      Begin MSComctlLib.ListView LVS 
         Height          =   2295
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4048
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Address"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Network"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame frmOpt 
      Height          =   3375
      Index           =   4
      Left            =   30
      TabIndex        =   25
      Top             =   360
      Width           =   5775
      Begin VB.Line Line1 
         BorderStyle     =   2  'Dash
         X1              =   240
         X2              =   5520
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.Frame frmOpt 
      Height          =   3375
      Index           =   3
      Left            =   30
      TabIndex        =   24
      Top             =   360
      Width           =   5775
      Begin VB.Frame Frame1 
         Caption         =   "Speed Settings(when i get to it)"
         Height          =   1455
         Left            =   240
         TabIndex        =   31
         Top             =   1800
         Width           =   5295
         Begin VB.Line Line2 
            BorderStyle     =   2  'Dash
            X1              =   120
            X2              =   5160
            Y1              =   720
            Y2              =   720
         End
      End
      Begin VB.CommandButton cmdDL 
         Caption         =   "..."
         Height          =   285
         Left            =   4920
         TabIndex        =   30
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtDL 
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Top             =   1320
         Width           =   4575
      End
      Begin MSComctlLib.Slider sldMaxUp 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   20
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label lblInfo 
         Caption         =   "Download Dir :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   28
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblInfo 
         Caption         =   "Max Upload :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IsOpen As Boolean

Private Sub chkPsv_Click()
If chkPsv.Value = vbChecked Then
    txtOption(2) = "0"
    txtOption(2).Enabled = False
Else
    txtOption(2) = GetSetting("INap", "Options", "Port", "6699")
    If txtOption(2) = 0 Then txtOption(2) = 6699
    txtOption(2).Enabled = True
End If
End Sub

Private Sub cmdApply_Click()
    Unload Me
End Sub

Private Sub cmdDirAdd_Click()
    List1.AddItem Dir1.path
End Sub

Private Sub cmdDirRem_Click()
    On Error Resume Next
    List1.RemoveItem List1.ListIndex
End Sub

Private Sub cmdDL_Click()
    Dim X As String
    X = SelectFolder(Me.hwnd)
    If Len(X) > 0 Then
        txtDL.Text = X
    End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim Address As String
Dim Network As String
Address = InputBox("Please enter address of server.", "Address", "127.0.0.1:8888")
If Address > "" Then
    Network = InputBox("Please enter the correct network name.", "Network", "Home")
        If Network > "" Then
            LVS.ListItems.Add , , Address
            LVS.ListItems.Item(LVS.ListItems.Count).SubItems(1) = Network
        End If
End If
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    LVS.ListItems.Remove LVS.SelectedItem.Index
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
End Sub

Private Sub Form_Load()
On Error Resume Next
IsOpen = False
Randomize
    txtOption(1) = GetSetting("INap", "Options", "Password", "Password")
    txtOption(0) = GetSetting("INap", "Options", "Nickname", "Nickname" & Int(Rnd * 5700))
    txtOption(2) = GetSetting("INap", "Options", "Port", "6699")
    cmbSpeed.ListIndex = GetSetting("INap", "Options", "Speed", "0")
    chkGator.Value = GetSetting("INap", "Options", "GetGator", "1")
    chkShow.Value = GetSetting("INap", "Options", "ShowOptions", "1")
    sldMaxUp.Value = GetSetting("INap", "Options", "MaxUp", "3")
    txtDL.Text = GetSetting("INap", "Options", "DownloadDir", App.path & "\Downloads")
    cmbVersion.ListIndex = 0
    chkPsv.Value = GetSetting("INap", "Options", "isPasive", "1")
    If chkPsv.Value = vbChecked Then
        chkPsv_Click
    Else
        If txtOption(2) = 0 Then
            txtOption(2) = 6699
        End If
    End If
frmOpt(0).ZOrder 0
Toolbar1.Buttons.Item(1).Value = tbrPressed
Open App.path & "\options\share.txt" For Input As #1
Do Until EOF(1)
    Dim a As String
    Line Input #1, a
    List1.AddItem a
Loop
Close #1
Open App.path & "\options\servers.txt" For Input As #1
Do Until EOF(1)
    Dim b() As String
    Line Input #1, a
    b = Split(a, "|")
    LVS.ListItems.Add , , b(0)
    LVS.ListItems.Item(LVS.ListItems.Count).SubItems(1) = b(1)
Loop
Close #1
If chkShow.Value = vbUnchecked And frmMain.Visible = False Then
    Unload Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Me.Hide
    Password = txtOption(1)
    Nickname = txtOption(0)
    Port = txtOption(2)
    Client = cmbVersion.Text  '"audioGnome" '"XNap 2.4-pre5"
    Speed = cmbSpeed.ListIndex
    GetGator = False
    DownloadDir = txtDL.Text
    MaxUp = sldMaxUp.Value
    SaveSetting "INap", "Options", "isPasive", chkPsv.Value
    SaveSetting "INap", "Options", "Password", Password
    SaveSetting "INap", "Options", "Nickname", Nickname
    SaveSetting "INap", "Options", "Port", Port
    SaveSetting "INap", "Options", "Speed", Speed
    SaveSetting "INap", "Options", "GetGator", chkGator.Value
    SaveSetting "INap", "Options", "ShowOptions", chkShow.Value
    SaveSetting "INap", "Options", "MaxUp", sldMaxUp.Value
    SaveSetting "INap", "Options", "DownloadDir", txtDL.Text
    frmShare.Show , frmMain
frmShare.lblShare.Caption = "Geting Server List!!"
    If chkGator.Value = vbChecked Then GetGator = True

    For i = 1 To frmOptions.LVS.ListItems.Count
        Err.Clear
        Dim hoe As Integer
        hoe = 0
        hoe = LCase(frmOptions.LVS.ListItems.Item(i).SubItems(1))
        If hoe = 0 Then
            frmMain.LVC.ListItems.Add , LCase(frmOptions.LVS.ListItems.Item(i).SubItems(1)), frmOptions.LVS.ListItems.Item(i)
            frmMain.LVC.ListItems(LCase(frmOptions.LVS.ListItems.Item(i).SubItems(1))).SubItems(1) = "Disconnected"
            frmMain.LVC.ListItems(LCase(frmOptions.LVS.ListItems.Item(i).SubItems(1))).SubItems(2) = "0"
            frmMain.LVC.ListItems(LCase(frmOptions.LVS.ListItems.Item(i).SubItems(1))).SubItems(3) = "0"
            frmMain.LVC.ListItems(LCase(frmOptions.LVS.ListItems.Item(i).SubItems(1))).SubItems(4) = "0"
            frmMain.LVC.ListItems(LCase(frmOptions.LVS.ListItems.Item(i).SubItems(1))).SubItems(5) = frmOptions.LVS.ListItems.Item(i).SubItems(1)
        Else
            frmMain.lvSL.ListItems.Add , frmOptions.LVS.ListItems.Item(i), frmOptions.LVS.ListItems.Item(i)
            frmMain.lvSL.ListItems(frmOptions.LVS.ListItems.Item(i)).SubItems(1) = frmOptions.LVS.ListItems.Item(i).SubItems(1)

        End If
    Next
    
If GetGator = True And IsOpen = False Then frmMain.LoadGator
DoEvents
frmMain.Show
Report:
    On Error Resume Next
    Err.Clear
    frmMain.Transfer(0).Close
    frmMain.Transfer(0).LocalPort = Port
    If Port <> 0 Then frmMain.Transfer(0).Listen
    If Err.Description > "" Then
    '    Port = InputBox(Err.Description & " Please Select a different port.", "Error", Port)
    '    GoTo Report
    MsgBox Err.Description, vbCritical, "Error"
    End If
    If IsOpen = False Then
        frmMain.LVS.ListItems.Clear
        For i = 0 To List1.ListCount - 1
        DirSize List1.List(i)
        Next
    End If
Unload frmShare
Open App.path & "\options\servers.txt" For Output As #1
    For i = 1 To LVS.ListItems.Count
        Print #1, LVS.ListItems.Item(i) & "|" & LVS.ListItems.Item(i).SubItems(1)
    Next
Close #1
frmMain.List1.Clear
Open App.path & "\options\share.txt" For Output As #1
    For i = 0 To List1.ListCount - 1
        Print #1, List1.List(i)
        frmMain.List1.AddItem List1.List(i)
    Next
Close #1
DoEvents
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    frmOpt(Button.Index - 1).ZOrder 0
    For i = 1 To Toolbar1.Buttons.Count
    Toolbar1.Buttons.Item(i).Value = tbrUnpressed
    Next
    Toolbar1.Buttons.Item(Button.Index).Value = tbrPressed
    
End Sub
