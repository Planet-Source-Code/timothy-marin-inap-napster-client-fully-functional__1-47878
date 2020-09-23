VERSION 5.00
Begin VB.Form frmIgnore 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ignore List"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remomve"
      Height          =   375
      Left            =   15
      TabIndex        =   1
      Top             =   2610
      Width           =   2160
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmIgnore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRemove_Click()
    On Error Resume Next
    For i = 0 To frmMain.Sock.UBound
        If frmMain.Sock(i).State = 7 Then
            frmMain.SendData List1.Text, NumToChr(323), Int(i)
        End If
    Next
    List1.RemoveItem List1.ListIndex
End Sub
