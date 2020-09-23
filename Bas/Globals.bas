Attribute VB_Name = "Globals"
Public Nickname As String
Public Password As String
Public Share As Long
Public Client As String
Public Speed As Integer
Public Port As String
'------------------------
Public GetGator As Boolean
'------------------------
Public DownloadDir As String
Public MaxUp As Integer
'------------------------
Public Network() As String 'assign a network to a socket
'


Public PM(50) As New frmPrivate

'Hold d/l u/l info
'UL
Public NSend() As Long
Public LSend() As Long
Public UFile() As String
'-------------------------
Public TState() As Integer
Public TSpeed() As Long
Public TSize() As Long
'
Function OpenChan()
    OpenChan = 51
    For i = 0 To 50
        If frmMain.Room(i).Visible = False Then
            OpenChan = i
            Exit Function
        End If
    Next
End Function

Function ChantoChan(Msg As String, Index As Integer)
    ChantoChan = 51
    For i = 0 To 50
        If frmMain.Room(i).Visible = True Then
            If frmMain.Room(i).CChannel = Msg And frmMain.Room(i).CServer = Index Then
            ChantoChan = i
            Exit Function
            End If
        End If
    Next
End Function

Function OpenPM()
    OpenPM = 51
    For i = 0 To 50
        If PM(i).Visible = False Then
            OpenPM = i
            Exit Function
        End If
    Next
End Function

Function RemChr(Msg As String)
    Dim RC() As String
    RemChr = Msg
    RC = Split("/,\,:,*,?,"",<,>,|", ",")
    For i = 0 To UBound(RC)
        RemChr = Replace(RemChr, RC(i), "")
    Next
End Function
Function PMtoChan(Msg As String, Index As Integer)
    PMtoChan = 51
    For i = 0 To 50
        If frmMain.Room(i).Visible = True Then
            If PM(i).CName = Msg And PM(i).CServer = Index Then
            PMtoChan = i
            Exit Function
            End If
        End If
    Next
End Function

Public Function DoConv(DC As String)
    On Error GoTo Errors:
    Dim NType As Integer
    Dim CType() As String
    CType = Split("B,KB,MB,GB,TB,PB,EB,ZB,YB", ",")
    Do Until DC < 1024
        DC = DC / 1024
        NType = NType + 1
    Loop
    If InStr(DC, ".") Then DC = Format(DC, "#.0#")
    DoConv = DC & " " & CType(NType)
    Exit Function
Errors:
    DoConv = "Error"
End Function


Function OpenSocket(Wsck As String, FRM As Form) As Integer
    Dim i As Integer
    For i = 1 To FRM.Controls(Wsck).UBound
        If FRM.Controls(Wsck)(i).State = 0 Then
            OpenSocket = i
            Exit Function
        End If
    Next
    i = FRM.Controls(Wsck).UBound + 1
    Load FRM.Controls(Wsck)(i)
    OpenSocket = i
End Function

Function NumToIco(Num As String)
NumToIco = 1
    Select Case Num
        Case 0, 1, 2, 3
            NumToIco = 1
        Case 4, 5, 6
            NumToIco = 2
        Case 7, 8
            NumToIco = 3
        Case 9, 10
            NumToIco = 4
    End Select
End Function

Function NumToSpeed(Num As String)
    NumToSpeed = "Unknown"
    Select Case Num
        Case 0
            NumToSpeed = "Unknown"
        Case 1
            NumToSpeed = "14.4K"
        Case 2
            NumToSpeed = "28.8K"
        Case 3
            NumToSpeed = "33.6K"
        Case 4
            NumToSpeed = "56K"
        Case 5
            NumToSpeed = "ISDN-56K"
        Case 6
            NumToSpeed = "ISDN-128K"
        Case 7
            NumToSpeed = "Cable"
        Case 8
            NumToSpeed = "DSL"
        Case 9
            NumToSpeed = "T1"
        Case 10
            NumToSpeed = "T3"
    End Select
End Function

Function SizeBy(ByVal L As MSComctlLib.ListView, Index As Integer)
    On Error Resume Next
    Dim LH As Long
    If L.ListItems.Count > 0 Then
        LH = L.ListItems.Item(1).Height * L.ListItems.Count
    End If
    Dim i As Integer
    Dim Hei As Long
    For i = 1 To L.ColumnHeaders.Count
        Hei = Hei + L.ColumnHeaders.Item(i).Width
    Next
    Dim Ad As Long
    Ad = L.Width - Hei
    If L.ColumnHeaders.Item(Index).Width + Ad - 100 > 1400 Then
        If LH > L.Height Then
            L.ColumnHeaders.Item(Index).Width = L.ColumnHeaders.Item(Index).Width + Ad - 400
        Else
            L.ColumnHeaders.Item(Index).Width = L.ColumnHeaders.Item(Index).Width + Ad - 100
        End If
    Else
        L.ColumnHeaders.Item(Index).Width = 1400
    End If
End Function

Function NumToChr(Num As Long)
    Dim z As Integer
    z = 0
    Do Until Num < 255
        z = z + 1
        Num = Num - 256
    Loop
    NumToChr = Chr(Num) & Chr(z)
End Function
