Attribute VB_Name = "Handle"
Public Buffer() As String
Public Busy() As Boolean
Public Sub HandleData(Sata As String, Index As Integer)
Busy(Index) = True
DoEvents
Dim HMR As Integer
HMR = 0
TOP:
HMR = HMR + 1
If HMR >= 5 Then
    HMR = 0
    DoEvents
End If
If Len(Sata) < 4 Then 'not even enough for protocoll specs sumthins fcked cut your losses
    Busy(Index) = False
    Buffer(Index) = ""
    Exit Sub
End If
On Error Resume Next
    Dim LENGTH As Long
    Dim What As String
    Dim Title As String
    Dim Data As String
    Dim Fchan As Integer
    Dim ChFo() As String
    Dim Chan() As String
    Dim Topic() As String
    LENGTH = Asc(Mid(Sata, 2, 1)) * 255 + Asc(Mid(Sata, 1, 1)) + 1
    What = Asc(Mid(Sata, 3, 1)) & "-" & Asc(Mid(Sata, 4, 1))
    Data = Mid(Sata, 5, LENGTH - 1)
    Data = Replace(Data, Chr(0), " ")
    If frmMain.mnuDebug.Checked = True Then
        frmMain.AddText Network(Index) & " : " & LENGTH & " : " & What & " : " & Data, vbBlack
    End If
If LENGTH > 2400 Or Asc(Mid(Sata, 4, 1)) > 3 Then 'not even enough for protocoll specs sumthins fcked cut your losses
    Busy(Index) = False
    Buffer(Index) = ""
    Exit Sub
End If

    If Len(Data) < LENGTH - 2 Then
        Buffer(Index) = Sata
        Busy(Index) = False
        Exit Sub
    End If
    Select Case What
        Case "0-0" 'Error
            frmMain.AddText Network(Index) & " : " & "*" & Data, vbRed
            frmMain.LVC.ListItems(Network(Index)).SubItems(1) = Data
            frmMain.Sock_Close Index
            'frmMain.Sock_Close Index
        Case "3-0" 'Login Ok Your Email Is
            '<email>
            For i = 0 To frmMain.lstIgnore.ListCount - 1
               frmMain.SendData frmMain.lstIgnore.List(i), NumToChr(322), Index
            Next
            For i = 1 To frmMain.LVIC.ListItems.Count
               frmMain.SendData frmMain.LVIC.ListItems.Item(i), NumToChr(208), Index
            Next
            For i = 1 To frmMain.LVNC.ListItems.Count
                frmMain.SendData frmMain.LVNC.ListItems.Item(i), NumToChr(208), Index
            Next
            frmMain.SendShare Index
        Case "5-0" 'Auto Upgrade
            '<version> <hostname:filename>
        Case "201-0" 'Search Reponce
            '"<filename>" <md5> <size> <bitrate> <frequency> <length> <nick> <ip> <link-type> [weight]
            '"<filename>" <md5> <size> 128 44100 182 PORCODIOPORCO 3616330558 0
            ChFo = Split(Data, """")
            Title = ChFo(1)
            Topic = Split(ChFo(2), " ")
            'If UBound(Topic) > 6 Then
                frmMain.LV1.ListItems.Add , , Topic(6)
                frmMain.LV1.ListItems.Item(frmMain.LV1.ListItems.Count).SubItems(5) = Topic(2)
                frmMain.LV1.ListItems.Item(frmMain.LV1.ListItems.Count).SubItems(1) = Mid(Title, InStrRev(Title, "\") + 1, Len(Title) - InStrRev(Title, "\"))
                frmMain.LV1.ListItems.Item(frmMain.LV1.ListItems.Count).SubItems(2) = DoConv(Topic(2))
                frmMain.LV1.ListItems.Item(frmMain.LV1.ListItems.Count).SubItems(3) = NumToSpeed(Topic(8))
                frmMain.LV1.ListItems.Item(frmMain.LV1.ListItems.Count).SubItems(4) = Topic(3)
                frmMain.LV1.ListItems.Item(frmMain.LV1.ListItems.Count).SmallIcon = NumToIco(Topic(8))
                frmMain.LV1.ListItems.Item(frmMain.LV1.ListItems.Count).Tag = Index
                frmMain.LV1.ListItems.Item(frmMain.LV1.ListItems.Count).ToolTipText = Title
                If IsNumeric(Topic(2)) = False Or IsNumeric(Topic(8)) = False Then
                    frmMain.LV1.ListItems.Remove frmMain.LV1.ListItems.Count
                End If
             
            'End If
        Case "202-0" 'Search Complete
            'Search.txtStat.Text = "Search Complete Found " & Search.LV1.ListItems.Count & " Matches..."
        Case "204-0" 'Start Dl Responce to 203-0
            '<nick> <ip> <port> "<filename>" <md5> <linespeed>
            ChFo = Split(Data, """")
            Dim DN, DA, DP, DS
            DS = Split(ChFo(0), " ")
            DN = DS(0)
            DA = DS(1)
            DP = DS(2)
            'MsgBox Data
            If DP = "0" Then
                frmMain.LVD.ListItems(ChFo(1) & DN).SubItems(2) = "Listening"
                'MsgBox DN & " """ & ChFo(1) & """"
                frmMain.SendData DN & " """ & ChFo(1) & """", Chr(244) & Chr(1), Index
            Else
                'MsgBox DP
                frmMain.LVD.ListItems(ChFo(1) & DN).SubItems(2) = "Connecting"
                frmMain.LVD.ListItems(ChFo(1) & DN).SubItems(3) = "0 B/s"
                Dim TSock As Integer
                TSock = OpenSocket("Transfer", frmMain)
                'MsgBox TSock
                'On Error GoTo 0
                frmMain.LVD.ListItems(ChFo(1) & DN).Key = "x" & TSock
                'frmMain.Caption = frmMain.LVD.ListItems("x" & TSock).SubItems(1)
                
                'On Error GoTo ConErr:
                'MsgBox Trim(AddrToIP(DA)) & " : """ & DP & """" & DA
                ReDim Preserve TState(frmMain.Transfer.UBound)
                ReDim Preserve TSpeed(frmMain.Transfer.UBound)
                ReDim Preserve TSize(frmMain.Transfer.UBound)
                TState(TSock) = 0
                TSpeed(TSock) = 0
                TSize(TSock) = frmMain.LVD.ListItems("x" & TSock).SubItems(4)
                'MsgBox TSize(TSock)
                frmMain.Transfer(TSock).LocalPort = 0
                frmMain.Transfer(TSock).Connect Trim(AddrToIP(DA)), DP

                GoTo IOK:
ConErr:
                'frmMain.LVD.ListItems(Index & ChFo(1)).Key = Index & ChFo(1)
                'frmMain.LVD.ListItems(Index & ChFo(1)).SubItems(2) = Err.Description
IOK:
End If
                On Error Resume Next
                'Connect to client on that ip/port
            'GetAscIp
        Case "205-0" 'Private Msg
            'Nick MSG
            ChFo = Split(Data, " ")
            Fchan = PMtoChan(CStr(ChFo(0)), Index)
            If Fchan > 50 Then
                Fchan = OpenPM
                If Fchan < 51 Then
                    PM(Fchan).Show
                    PM(Fchan).CServer = Index
                    PM(Fchan).Caption = "(" & ChFo(0) & ") Instant Message"
                    PM(Fchan).CName = ChFo(0)
                    PM(Fchan).AddText "<" & ChFo(0) & "> " & Right(Data, Len(Data) - InStr(Data, " ")), vbBlack
                End If
            Else
                    PM(Fchan).Caption = "(" & ChFo(0) & ") Instant Message"
                    PM(Fchan).AddText "<" & ChFo(0) & "> " & Right(Data, Len(Data) - InStr(Data, " ")), vbBlack
            End If
        Case "206-0" 'Download Err
            '<nick> "<filename>"
            ChFo = Split(Data, """")
            Title = ChFo(0)
            'MsgBox """" & LCase(frmMain.LVD.ListItems(Index & ChFo(1))) & """" & " """ & RTrim(LCase(Title)) & """"
            'If LCase(frmMain.LVD.ListItems(Index & ChFo(1))) = RTrim(LCase(Title)) Then
                frmMain.LVD.ListItems(ChFo(1) & Trim(Title)).SubItems(2) = "Error"
            'End If
        Case "209-0" 'Hotlist User Signon
            '<user> <speed>
            
            ChFo = Split(Data, " ")
            Dim CheckHL As Integer
            CheckHL = 0
            CheckHL = frmMain.LVNC.ListItems(ChFo(0)).Index
            If CheckHL = 0 Then GoTo ends
            frmMain.LVIC.ListItems.Add , ChFo(0), ChFo(0), , NumToIco(ChFo(1))
            frmMain.LVIC.ListItems(ChFo(0)).Tag = Index
            frmMain.LVNC.ListItems.Remove frmMain.LVNC.ListItems(ChFo(0)).Index
        Case "210-0" 'Hotlist User Signoff
            '<user>
            CheckHL = 0
            CheckHL = frmMain.LVIC.ListItems(Data).Index
            If CheckHL = 0 Then GoTo ends
            'Remove From Connected but resend entry
            frmMain.LVNC.ListItems.Add , Data, Data, , 6
            frmMain.LVIC.ListItems.Remove frmMain.LVIC.ListItems(Data).Index
                For i = 1 To frmMain.Sock.UBound
                    If frmMain.Sock(i).State = 7 Then frmMain.SendData Data, NumToChr(303), Int(i)
                    If frmMain.Sock(i).State = 7 Then frmMain.SendData Data, NumToChr(207), Int(i)
                Next
        Case "212-0" 'Browse Users Files
            '<nick> "<filename>" <md5> <size> <bitrate> <frequency> <time> <nick>
            '"<filename>" <md5> <size> <bitrate> <frequency> <length> <nick> <ip> <link-type> [weight]
            ChFo = Split(Data, """")


            Title = ChFo(1)
            Topic = Split(ChFo(2), " ")
            If UBound(Topic) > 4 Then
            frmMain.LVB.ListItems.Add , , Trim(ChFo(0))
            frmMain.LVB.ListItems.Item(frmMain.LVB.ListItems.Count).SubItems(5) = Topic(2)
            frmMain.LVB.ListItems.Item(frmMain.LVB.ListItems.Count).SubItems(1) = Mid(Title, InStrRev(Title, "\") + 1, Len(Title) - InStrRev(Title, "\"))
            frmMain.LVB.ListItems.Item(frmMain.LVB.ListItems.Count).SubItems(2) = DoConv(Topic(2))
            'frmMain.LVB.ListItems.Item(frmMain.LVB.ListItems.Count).SubItems(3) = NumToSpeed(Topic(8))
            frmMain.LVB.ListItems.Item(frmMain.LVB.ListItems.Count).SubItems(4) = Topic(3)
            frmMain.LVB.ListItems.Item(frmMain.LVB.ListItems.Count).SmallIcon = NumToIco(Topic(8))
            frmMain.LVB.ListItems.Item(frmMain.LVB.ListItems.Count).Tag = Index
            frmMain.LVB.ListItems.Item(frmMain.LVB.ListItems.Count).ToolTipText = Title
            If IsNumeric(Topic(2)) = False Then
            
                frmMain.LVB.ListItems.Remove frmMain.LVB.ListItems.Count
            End If
            End If
        Case "213-0" 'End Browse
            '<nick> [ip]
            frmMain.P(7).ZOrder 0
                For i = 1 To frmMain.TB1.Buttons.Count
                frmMain.TB1.Buttons.Item(i).Value = tbrUnpressed
                Next
                
                frmMain.TB1.Buttons.Item(11).Value = tbrPressed
        Case "214-0" 'Stats
            '<users> <# files> <size>
            Dim Stats() As String
            Stats = Split(Data, " ")
            'frmMain.S1.Panels.Item(2).Text = "Currently " & Stats(1) & " Files (" & Stats(2) & " Gigabytes) available in " & Stats(0) & " libraries."
            frmMain.LVC.ListItems(Network(Index)).SubItems(2) = Stats(0)
            frmMain.LVC.ListItems(Network(Index)).SubItems(3) = Stats(1)
            frmMain.LVC.ListItems(Network(Index)).SubItems(4) = Stats(2)
        Case "216-0" 'Resume File
            '<user> <ip> <port> <filename> <checksum> <size> <speed>
        Case "217-0" 'End Resume List
            '
        Case "45-1" 'User Aded to hotlist
            'User
        Case "46-1" '302 (0x12e) hotlist error
            '<user>
        Case "64-1" '320 (0x140) user ignore list
            '<count>
        Case "65-1" '321 (0x141) user ignore list entry
            '<user>
        Case "66-1" '322 (0x142) add user to ignore list
            '<user>
        Case "68-1" '324 (0x144) user is not ignored
            'User
        Case "69-1" '325 (0x145) user is already ignored
            'user
        Case "70-1" '326 (0x146) clear ignore list
            '<count>
        Case "145-1" '401 Disconnected from room
            '<channel>
            ChFo = Split(Data, " ")
            Fchan = ChantoChan(CStr(ChFo(0)), Index)
            If Fchan < 51 Then
                frmMain.Room(Fchan).LeaveRoom
            End If
        Case "147-1" '403 (0x193) public message
            '<channel> <nick> <text>
            ChFo = Split(Data, " ")
            Fchan = ChantoChan(CStr(ChFo(0)), Index)
            If Fchan < 51 Then
            If ChFo(1) = Nickname Then
                frmMain.Room(Fchan).AddText "<" & ChFo(1) & "> " & Right(Data, Len(Data) - 2 - Len(ChFo(0)) - Len(ChFo(1))), 8388608
            Else
                frmMain.Room(Fchan).AddText "<" & ChFo(1) & "> " & Right(Data, Len(Data) - 2 - Len(ChFo(0)) - Len(ChFo(1))), vbBlack
            End If
            frmMain.Room(Fchan).ChanMsg ChFo(1), Right(Data, Len(Data) - 2 - Len(ChFo(0)) - Len(ChFo(1)))
                'If InStr(LCase(Data), "!client") Then
                '    SendData ChFo(0) & " " & "INap1.1", Chr(146) & Chr(1)
                'End If
            End If
        Case "148-1" '404 (0x194) error message
            '<error>
            frmMain.AddText Network(Index) & " : " & Data, vbRed
        Case "149-1" '405 (0x195) join acknowledge
            '<Channel>
            frmMain.P(2).ZOrder 0
            Fchan = OpenChan
            If Fchan < 51 Then
                frmMain.Room(Fchan).Visible = True
                frmMain.Room(Fchan).ZOrder 0
                frmMain.mnuRoom(Fchan).Visible = True
                frmMain.mnuRoom(Fchan).Caption = Network(Index) & " : " & Data
                frmMain.mnuRoom_Click Fchan
                frmMain.Room(Fchan).CServer = Index
                frmMain.Room(Fchan).CChannel = Data
                frmMain.Room(Fchan).AddText "Joined Room " & Data & "...", 8421376
                frmMain.Room(Fchan).Left = 0
                frmMain.Room(Fchan).TOP = 0
                frmMain.Room(Fchan).Width = frmMain.P(1).Width
                frmMain.Room(Fchan).Height = frmMain.P(1).Height
            End If
        Case "150-1" '406 (0x196) join message
            '<channel> <user> <sharing> <link-type>
            ChFo = Split(Data, " ")
            Fchan = ChantoChan(CStr(ChFo(0)), Index)
            If Fchan < 51 Then
                frmMain.Room(Fchan).AddText "*User " & ChFo(1) & " (" & NumToSpeed(ChFo(3)) & ") [sharing " & ChFo(2) & " files] has joined...", 8421376
                frmMain.Room(Fchan).AddUser ChFo(1), ChFo(2), NumToSpeed(ChFo(3)), NumToIco(ChFo(3))
            End If
        Case "151-1" '407 (0x197) user parted channel
            '<channel> <nick> <sharing> <linespeed>
            ChFo = Split(Data, " ")
            Fchan = ChantoChan(CStr(ChFo(0)), Index)
            If Fchan < 51 Then
                ' * Tim5456 (Cable) [sharing 1871 files] has joined.
                ' * Tim5456 (Cable) [sharing 1871 files] has left.
                frmMain.Room(Fchan).AddText "*User " & ChFo(1) & " (" & NumToSpeed(ChFo(3)) & ") [sharing " & ChFo(2) & " files] Has Left...", 192
                frmMain.Room(Fchan).RemoveUser ChFo(1)
            End If
        Case "152-1" '408 (0x198) channel user list entry
            '<channel> <user> <sharing> <link-type>
            ChFo = Split(Data, " ")
            Fchan = ChantoChan(CStr(ChFo(0)), Index)
            If Fchan < 51 Then
                frmMain.Room(Fchan).AddUser ChFo(1), ChFo(2), NumToSpeed(ChFo(3)), NumToIco(ChFo(3))
            End If
        Case "153-1" '409 (0x199) end of channel user list
            '<channel>
        Case "154-1" '410 (0x19a) channel topic
            '<channel> <topic>
            ChFo = Split(Data, " ")
            Fchan = ChantoChan(CStr(ChFo(0)), Index)
            If Fchan < 51 Then
                frmMain.Room(Fchan).CTopic = "Topic : " & Right(Data, Len(Data) - (Len(ChFo(0)) + 1))
            End If
        Case "164-1" '420 (0x1a4) channel ban list
            '<channel>
        Case "165-1" '421 (0x1a5) channel ban list entry
            '
        Case "169-1" '425 channel motd
            '<message>
        Case "174-1" '430 invite a user
            '<nick> <channel> "<topic>" <unknown_digit> <unknown_text>
        Case "245-1" '501 (0x1f5) alternate download ack
            '<nick> <ip> <port> "<filename>" <md5> <speed>
                'this message is sent to the uploader when their data port is set to
            '0 to indicate they are behind a firewall and need to push all data
            'outware.  the uploader is responsible for connecting to the
            'downloader to transfer the file.
            ChFo = Split(Data, """")
            Dim XUP() As String
            XUP = Split(ChFo(0), " ")
            Dim isShare As Long
            isShare = 0
            isShare = frmMain.LVU.ListItems(ChFo(1) & XUP(0)).Index
            If isShare > 0 Then
                TSock = OpenSocket("Transfer", frmMain)
                ReDim Preserve TState(frmMain.Transfer.UBound)
                ReDim Preserve TSpeed(frmMain.Transfer.UBound)
                ReDim Preserve TSize(frmMain.Transfer.UBound)
                TState(TSock) = 0
                TSpeed(TSock) = 0
                frmMain.LVU.ListItems(ChFo(1) & XUP(0)).SubItems(2) = "Connecting"
                frmMain.LVU.ListItems(ChFo(1) & XUP(0)).Key = "x" & TSock
                frmMain.Transfer(TSock).LocalPort = 0
                frmMain.Transfer(TSock).Connect AddrToIP(XUP(1)), XUP(2)
            End If
        Case "89-2" '601 (0x259) link speed response
            '<nick> <linespeed>
        Case "92-2" '604 (0x25c) whois response [SERVER]
            '<nick> "<user-level>" <time> "<channels>" "<status>" <shared>
            '<downloads> <uploads> <link-type> "<client-info>" [ <total downloads>
            '<total_uploads> <ip> <connecting port> <data port> <email> ]
            'Tim5456 "Elite" 10 "Alternative " Active 1888 0 0 7 "INap" 0 0 127.0.0.1 54281 6699 anon@localhost localhost
            'Tim5456 "Human" 33 "Tater'sRetreat " "Active" 1888 0 0 7 "INap" 0 0 68.59.144.106 57353 6699 anon@God God 0 8
            ChFo = Split(Data, """")
            Dim WIChan As String
            Dim WILevel As String
            Dim WIStatus As String
            Dim WINick As String
            Dim WIVersion As String
            Dim WITime As String
                WINick = Trim(ChFo(0))
                WIChan = ChFo(3)
                WILevel = ChFo(1)
                WITime = Trim(ChFo(2))
                Unload frmInfo
                frmInfo.Show
                Dim WIOther() As String
                Dim WIOps() As String
            If UBound(ChFo) > 6 Then
                WIStatus = ChFo(5)
                WIVersion = ChFo(7)
                WIOther = Split(ChFo(6), " ")
            Else
                WIVersion = ChFo(5)
                WIOther = Split(ChFo(4), " ")
                WIStatus = WIOther(1)
                WIOther = Split(Right(ChFo(4), Len(ChFo(4)) - (Len(WIStatus) + 1)), " ")
            End If
            WIOps = Split(ChFo(UBound(ChFo)), " ")
            ' 0 0 127.0.0.1 54281 6699 anon@localhost localhost
            frmInfo.LvI.ListItems.Add , , "Status"
            frmInfo.LvI.ListItems.Add , , "Time Connected"
            frmInfo.LvI.ListItems.Add , , "User"
            frmInfo.LvI.ListItems.Add , , "Client Version"
            frmInfo.LvI.ListItems.Add , , "User Level"
            frmInfo.LvI.ListItems.Add , , "Connection"
            frmInfo.LvI.ListItems.Add , , "Channels"
            frmInfo.LvI.ListItems.Add , , "Number Shared"
            frmInfo.LvI.ListItems.Add , , "Current Downloads"
            frmInfo.LvI.ListItems.Add , , "Current Uploads"
            frmInfo.LvI.ListItems.Add , , "Total Downloads"
            frmInfo.LvI.ListItems.Add , , "Total Uploads"
            frmInfo.LvI.ListItems.Add , , "Address"
            frmInfo.LvI.ListItems.Add , , "Data Port"
            frmInfo.LvI.ListItems.Add , , "Port"
            frmInfo.LvI.ListItems.Add , , "Email"
            frmInfo.LvI.ListItems.Item(1).SubItems(1) = WIStatus
            Dim WId, WIh, WIm
            WId = 0
            WIh = 0
            WIm = 0
            Do Until WITime < 60
                DoEvents
                WIm = WIm + 1
                WITime = WITime - 60
            Loop
            Do Until WIm < 60
                DoEvents
                WIh = WIh + 1
                WIm = WIm - 60
            Loop
            Do Until WIh < 24
                DoEvents
                WId = WId + 1
                WIh = WIh - 24
            Loop
            If WId = 1 Then
                WId = "1 Day "
            ElseIf WId > 1 Then
                WId = WId & " Days "
            Else
                WId = ""
            End If
            frmInfo.LvI.ListItems.Item(2).SubItems(1) = WId & WIh & ":" & WIm & ":" & WITime
            frmInfo.LvI.ListItems.Item(3).SubItems(1) = WINick
            frmInfo.LvI.ListItems.Item(4).SubItems(1) = WIVersion
            frmInfo.LvI.ListItems.Item(5).SubItems(1) = WILevel
            frmInfo.LvI.ListItems.Item(6).SubItems(1) = NumToSpeed(WIOther(4))
            frmInfo.LvI.ListItems.Item(7).SubItems(1) = WIChan
            frmInfo.LvI.ListItems.Item(8).SubItems(1) = WIOther(1)
            frmInfo.LvI.ListItems.Item(9).SubItems(1) = WIOther(2)
            frmInfo.LvI.ListItems.Item(10).SubItems(1) = WIOther(3)
            frmInfo.LvI.ListItems.Item(11).SubItems(1) = WIOps(1)
            frmInfo.LvI.ListItems.Item(12).SubItems(1) = WIOps(2)
            frmInfo.LvI.ListItems.Item(13).SubItems(1) = WIOps(3)
            frmInfo.LvI.ListItems.Item(14).SubItems(1) = WIOps(4)
            frmInfo.LvI.ListItems.Item(15).SubItems(1) = WIOps(5)
            frmInfo.LvI.ListItems.Item(16).SubItems(1) = WIOps(6)
            frmMain.AddText Data, vbRed
        Case "93-2" '605 (0x25d) whowas response
            '<user> <level> <last-seen>
        Case "95-2" '607 (0x25f) upload request ( someone wants a file from you )
            '<nick> "<filename>" <speed>
            ChFo = Split(Data, """")
            Title = Trim(ChFo(0))
            'if you have open slots send an ok msg
            '608(96-2) accept 609(97-2) deny
            'frmMain.SendData Title & """" & ChFo(1) & """", Chr(97) & Chr(2), Index
            'Check if your sharing the file.
            isShare = 0
            isShare = frmMain.LVS.ListItems(ChFo(1)).Index
            If isShare > 0 And frmMain.LVU.ListItems.Count < MaxUp Then
                frmMain.LVU.ListItems.Add , ChFo(1) & Title, Title
                frmMain.LVU.ListItems(ChFo(1) & Title).SubItems(1) = ChFo(1)
                frmMain.LVU.ListItems(ChFo(1) & Title).SubItems(2) = "Waiting"
                frmMain.LVU.ListItems(ChFo(1) & Title).SubItems(3) = "0 B/s"
                frmMain.LVU.ListItems(ChFo(1) & Title).Tag = Index
                frmMain.tmrDL.Enabled = False
                frmMain.tmrDL.Enabled = True
                frmMain.SendData Title & " """ & ChFo(1) & """", Chr(96) & Chr(2), Index
            Else
                frmMain.SendData Title & " """ & ChFo(1) & """ " & MaxUp, Chr(107) & Chr(2), Index
            End If
            
        Case "97-2" '609 (0x261)     accept failed
            '<nick> "<filename>"
            ChFo = Split(Data, " ")
            Title = ChFo(0)
            Dim EFname As String
            For i = 1 To UBound(ChFo)
                EFname = EFname & " " & ChFo(i)
            Next
            EFname = Replace(LTrim(EFname), """", "")
            'MsgBox """" & LCase(frmMain.LVD.ListItems(Index & ChFo(1))) & """" & " """ & RTrim(LCase(Title)) & """"
            'If LCase(frmMain.LVD.ListItems(Index & ChFo(1))) = RTrim(LCase(Title)) Then
                frmMain.LVD.ListItems(EFname & Trim(Title)).SubItems(2) = "Accept Failed"
            'End If
        Case "103-2" '613 (0x265) set data port for user
            '<port>
        Case "103-2" '615 (0x267) End Ban List
            '
        Case "104-2" '616 (0x268) (ip?) ban list entry
            '<ip> <nick> "<reason>" <time> <n>
        Case "105-2" '617 (0x269) End list channels
            'MsgBox "Received Channel List...", , "INap"
            frmMain.AddText Network(Index) & " : " & "Received Channel List...", vbBlue
        Case "106-2" '618 (0x26a) channel list entry
            '<channel-name> <number-of-users> <topic>
            Chan = Split(Data, " ")
            frmRooms.LV1.ListItems.Add , , Network(Index)
            frmRooms.LV1.ListItems.Item(frmRooms.LV1.ListItems.Count).Tag = Index
            frmRooms.LV1.ListItems.Item(frmRooms.LV1.ListItems.Count).SubItems(1) = Chan(0)
            frmRooms.LV1.ListItems.Item(frmRooms.LV1.ListItems.Count).SubItems(2) = Chan(1)
            frmRooms.LV1.ListItems.Item(frmRooms.LV1.ListItems.Count).SubItems(3) = Mid(Data, Len(Chan(1)) + Len(Chan(0)) + 2, Len(Data))
        Case "108-2" '620 (0x26c) queue limit
            '<nick> "<filename>" <filesize> <digit>
            ChFo = Split(Data, """")
            Title = ChFo(0)
            'MsgBox """" & LCase(frmMain.LVD.ListItems(Index & ChFo(1))) & """" & " """ & RTrim(LCase(Title)) & """"
            'If LCase(frmMain.LVD.ListItems(Index & ChFo(1))) = RTrim(LCase(Title)) Then
                frmMain.LVD.ListItems(ChFo(1) & Trim(Title)).SubItems(2) = "Remotely Queued"
            'End If
        Case "109-2" '621 (0x26d) message of the day
            frmMain.AddText Network(Index) & " : " & "*" & Data, 8388608
        Case "115-2" '627 (0x273) operator message
            '<nick> <text>
            ChFo = Split(Data, " ")
            frmMain.AddText Network(Index) & " : " & "<" & ChFo(0) & "> " & ChFo(1), 32768
        Case "116-2" '628 (0x274) global message
            '<nick> <text>
            ChFo = Split(Data, " ")
            frmMain.AddText Network(Index) & " : " & "<" & ChFo(0) & "> " & ChFo(1), 128
        Case "117-2" '629 (0x275) banned users
            '<nick>
        Case "128-2" '640 direct browse request
            '<nick> [ip port] 'Someone wants to browse ur files [] if passive
        Case "129-2" '641 direct browse accept
            '<nick> <ip> <port>
        Case "130-2" '642 direct browse error [SERVER]
            '<nick> "message"
        Case "236-2" '748 login attempt
            '
        Case "238-2" '750 (0x2ee) server ping
            '
        Case "239-2" '751 (0x2ef) ping user
            '<user> 'Someone Pinged you
            frmMain.SendData Data, Chr(240) & Chr(2), Index
            frmMain.AddText Network(Index) & " : " & "!Pong " & Data, vbMagenta
        Case "240-2" '752 (0x2f0) pong response
            '<user> 'Receaved Ping Responce
            frmMain.AddText Network(Index) & " : " & "*" & Now & " " & Data & " Pong!", vbBlue
        Case "53-3" '821 (0x335) redirect client to another server
            '<server> <port>
        Case "54-3" '822 (0x336) cycle client
            '<metaserver> 'Disco and Connect to Meta
        Case "56-3" '824 (0x338) emote
            '<channel> <user> "<text>"
            ChFo = Split(Data, " ")
            Fchan = ChantoChan(CStr(ChFo(0)), Index)
            If Fchan < 51 Then
                frmMain.Room(Fchan).AddText "*" & ChFo(1) & " " & Left(Right(Data, Len(Data) - 3 - Len(ChFo(0)) - Len(ChFo(1))), Len(Right(Data, Len(Data) - 3 - Len(ChFo(0)) - Len(ChFo(1)))) - 1), 8388608
            End If
        Case "57-3" '825 (0x339) user list entry
            '<channel> <user> <files shared> <speed>
        Case "59-3" '827 (0x33b) show all channels
            'Channel list Complete
            frmMain.AddText Network(Index) & " : " & "Received Channel List...", vbBlue
        Case "60-3" '828 (0x33c) channel list
            '<channel> <users> <n1> <level> <limit> "<topic>"
            Topic = Split(Data, """")
            Chan = Split(Topic(0), " ")
            frmRooms.LV1.ListItems.Add , , Network(Index)
            frmRooms.LV1.ListItems.Item(frmRooms.LV1.ListItems.Count).Tag = Index
            frmRooms.LV1.ListItems.Item(frmRooms.LV1.ListItems.Count).SubItems(1) = Chan(0)
            frmRooms.LV1.ListItems.Item(frmRooms.LV1.ListItems.Count).SubItems(2) = Chan(1)
            frmRooms.LV1.ListItems.Item(frmRooms.LV1.ListItems.Count).SubItems(3) = Topic(1)
        Case "62-3" '830 (0x33e) list users in channel
            '<channel>
        Case "132-3" '900 connection test
            '<ip> <port> <data> 'Connect and send
        Case "133-3" '901 Listen test
            '<port> <timeout> <data> 'Listen On Connect Send
        Case Else
            frmMain.AddText Network(Index) & " : " & "UNKNOWN : " & LENGTH & " : " & What & " : " & Data, vbGreen
    End Select
ends:
Sata = Right(Sata, Len(Sata) - 4 - Len(Data)) & Buffer(Index)
    If Len(Sata) > 0 Then
        GoTo TOP:
    End If
    DoEvents
    Busy(Index) = False
    Buffer(Index) = ""
End Sub

