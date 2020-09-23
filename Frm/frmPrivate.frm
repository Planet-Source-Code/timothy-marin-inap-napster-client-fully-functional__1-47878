VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPrivate 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmPrivate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHot 
      Caption         =   "Add User To Hotlist"
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   6720
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "Ignore"
      Height          =   330
      Left            =   2640
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Veiw  Info"
      Height          =   330
      Left            =   5280
      TabIndex        =   1
      Top             =   3360
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmPrivate.frx":014A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPrivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CServer As Integer
Public CName As String


Private Sub cmdHot_Click()
    On Error GoTo Err
    frmMain.LVNC.ListItems.Add , CName, CName, , 6
    For i = 1 To frmMain.Sock.UBound
        If frmMain.Sock(i).State = 7 Then frmMain.SendData CName, NumToChr(207), Int(i)
    Next
Err:
End Sub

Private Sub cmdIgnore_Click()
    On Error Resume Next
    frmMain.lstIgnore.AddItem CName
    For i = 0 To frmMain.Sock.UBound
        If frmMain.Sock(i).State = 7 Then
            frmMain.SendData CName, NumToChr(322), Int(i)
        End If
    Next
End Sub

Private Sub cmdView_Click()
    frmMain.SendData CName, Chr(91) & Chr(2), CServer '603 (0x25b) whois request
End Sub

Sub AddText(msg As String, Vcol As ColorConstants)
    RTB.SelStart = Len(RTB.Text)
    RTB.SelColor = Vcol
    RTB.SelText = vbCrLf & msg
End Sub

Private Sub Form_Unload(Cancel As Integer)
CServer = 0
CName = ""
End Sub

Private Sub RTB_Change()
    RTB.SelStart = Len(RTB.Text)
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        frmMain.SendData CName & " " & txtSend.Text, Chr(205) & Chr(0), CServer
        'MsgBox Trim(Me.Caption) & " " & txtSend.Text & " " & CServer
        AddText "<" & Nickname & "> " & txtSend.Text, 8388608
        txtSend.Text = ""
    End If
End Sub

