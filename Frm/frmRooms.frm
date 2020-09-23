VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRooms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Channels"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   Icon            =   "frmRooms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   7335
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Server"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Room"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Users"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Topic"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdJoin_Click()
    On Error Resume Next
    If LV1.SelectedItem.SubItems(1) <> "" Then
        frmMain.SendData LV1.SelectedItem.SubItems(1), Chr(144) & Chr(1), LV1.SelectedItem.Tag
        '400 (0x190) join channel
    End If
    Unload Me
End Sub

Private Sub Form_Load()
SizeBy LV1, 4
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
cmdJoin_Click
End Sub
