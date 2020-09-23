Attribute VB_Name = "basWinsock"
Option Explicit
Public Const INADDR_NONE = &HFFFF
Public Const hostent_size = 16
    Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
    Declare Function lstrlen Lib "kernel32.dll" (ByVal lpString As Any) As Integer
    Public Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
    Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
    Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
    Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal Host_Name As String) As Long
Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Function GetAscIp(ByVal ipl$) As String
    On Error Resume Next
    Dim lpStr&
    Dim nStr&
    Dim retString$
    Dim inn As Long
    retString = String(32, 0)
    inn = Val(ipl)
    inn = ntohl(inn)
    lpStr = inet_ntoa(inn)
    If lpStr = 0 Then
        GetAscIp = "255.255.255.255"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    GetAscIp = retString
    If Err Then GetAscIp = "255.255.255.255"
End Function

Function AddrToIP(ByVal AddrOrIP$) As String
    On Error Resume Next
    AddrToIP$ = GetAscIp(GetHostByNameAlias(AddrOrIP$))
    'MsgBox AddrToIP$
    If Err Then AddrToIP$ = "255.255.255.255"
End Function

Function GetHostByNameAlias(ByVal Hostname$) As Long
  On Error Resume Next
  'Return IP address as a long, in network byte order
  Dim phe&    ' pointer to host information entry
  Dim heDestHost As HostEnt 'hostent structure
  Dim addrList&
  Dim retIP&
  'first check to see if what we have been passed is a valid IP
    retIP = inet_addr(Hostname)
    If retIP = INADDR_NONE Then
        'it wasn't an IP, so do a DNS lookup
        phe = gethostbyname(Hostname)
        If phe <> 0 Then
            'Pointer is non-null, so copy in hostent structure
            MemCopy heDestHost, ByVal phe, hostent_size
            'Now get first pointer in address list
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
        Else
            'its not a valid address
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
    If Err Then GetHostByNameAlias = INADDR_NONE
End Function
