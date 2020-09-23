Attribute VB_Name = "Commands"
Function IsCommand(Msg As String, Chan As String, Index As Integer)
    Dim CInfo() As String
    Dim IH As String
    IsCommand = False
    If Left(Msg, 1) <> "/" Then Exit Function
    CInfo = Split(Msg, " ")
    Select Case LCase(CInfo(0))
        Case "/ban" '<user|address>
                If UBound(CInfo) > 0 Then
                    Msg = Mid(Msg, InStr(Msg, " "), Len(Msg))
                    frmMain.SendData Msg, NumToChr(612), Index
                Else
                    If Chan = 0 Then
                        frmMain.AddText "*Proper Syntax /ban user|ip [""reson"" timeout]", vbRed
                    Else
                        IH = ChantoChan(Chan, Index)
                        If IH < 51 Then
                            frmMain.Room(IH).AddText "*Proper Syntax /ban user|ip [""reson"" timeout]", vbRed
                        End If
                    End If
                End If
        Case "/setuserlevel" '<user> <level> 01234 leech user mod admin elite
                If UBound(CInfo) > 1 Then
                    Msg = Mid(Msg, InStr(Msg, " "), Len(Msg))
                    frmMain.SendData Msg, NumToChr(606), Index
                Else
                    If Chan = 0 Then
                        frmMain.AddText "*Proper Syntax /setuserlevel level", vbRed
                    Else
                        IH = ChantoChan(Chan, Index)
                        If IH < 51 Then
                            frmMain.Room(IH).AddText "*Proper Syntax /setuserlevel level", vbRed
                        End If
                    End If
                End If
        Case "/clear" '<channel> 'Clear ur channel text if none then clear main
            'If UBound(CInfo) > 0 Then
                If Chan = "0" Then
                    frmMain.RTB.Text = ""
                Else
                    IH = ChantoChan(Chan, Index)
                    If IH < 51 Then
                        frmMain.Room(IH).ClearRoom
                    End If
                End If
        Case "/kick" '<channel> <user> ["reason"]
            If channel = "0" Then
                If UBound(CInfo) > 1 Then
                    Msg = Mid(Msg, InStr(Msg, " "), Len(Msg))
                    frmMain.SendData Msg, NumToChr(829), Index
                Else
                    frmMain.AddText "*Proper Syntax /kick Channel user [""reson""]", vbRed
                End If
            Else
                If UBound(CInfo) > 0 Then
                    frmMain.SendData Chan & " " & Msg, NumToChr(829), Index
                Else
                    IH = ChantoChan(Chan, Index)
                    If IH < 51 Then
                        frmMain.Room(IH).AddText "*Proper Syntax /kick user [""reson""]", vbRed
                    End If
                End If
            End If
        Case "/whois" '<user>
            If UBound(CInfo) > 0 Then
                frmMain.SendData CInfo(1), Chr(91) & Chr(2), Index '603 (0x25b) whois request
            Else
                frmMain.AddText "*Proper Syntax /whois user", vbRed
            End If
        Case "/muzzle" '<user> <reson>
                If UBound(CInfo) > 0 Then
                    Msg = Mid(Msg, InStr(Msg, " "), Len(Msg))
                    frmMain.SendData Msg, NumToChr(829), Index
                Else
                    If Chan = 0 Then
                        frmMain.AddText "*Proper Syntax /muzzle user [""reson""]", vbRed
                    Else
                        IH = ChantoChan(Chan, Index)
                        If IH < 51 Then
                            frmMain.Room(IH).AddText "*Proper Syntax /muzzle user [""reson""]", vbRed
                        End If
                    End If
                End If
        Case "/topic" '<channel> [topic] - display/set channel topic
            If channel = "0" Then
                If UBound(CInfo) > 1 Then
                    Msg = Mid(Msg, InStr(Msg, " "), Len(Msg))
                    frmMain.SendData Msg, NumToChr(410), Index
                Else
                    frmMain.AddText "*Proper Syntax /Topic Channel Topic", vbRed
                End If
            Else
                If UBound(CInfo) > 0 Then
                    frmMain.SendData Chan & " " & Msg, NumToChr(410), Index
                Else
                    IH = ChantoChan(Chan, Index)
                    If IH < 51 Then
                        frmMain.Room(IH).AddText "*Proper Syntax /kick user [""reson""]", vbRed
                    End If
                End If
            End If
        Case "/unban" '<user|address>
                If UBound(CInfo) > 0 Then
                    Msg = Mid(Msg, InStr(Msg, " "), Len(Msg))
                    frmMain.SendData Msg, NumToChr(423), Index
                Else
                    If Chan = 0 Then
                        frmMain.AddText "*Proper Syntax /unban user|ip [""reson""]", vbRed
                    Else
                        IH = ChantoChan(Chan, Index)
                        If IH < 51 Then
                            frmMain.Room(IH).AddText "*Proper Syntax /unban user|ip [""reson""]", vbRed
                        End If
                    End If
                End If
        Case "/unmuzzle" '<user|address>
                If UBound(CInfo) > 0 Then
                    Msg = Mid(Msg, InStr(Msg, " "), Len(Msg))
                    frmMain.SendData Msg, NumToChr(623), Index
                Else
                    If Chan = 0 Then
                        frmMain.AddText "*Proper Syntax /unmuzzle nick [""reson""]", vbRed
                    Else
                        IH = ChantoChan(Chan, Index)
                        If IH < 51 Then
                            frmMain.Room(IH).AddText "*Proper Syntax /unmuzzle nick [""reson""]", vbRed
                        End If
                    End If
                End If
        Case "/cloak" ' - toggle invisibility to normal users"
            frmMain.SendData "", NumToChr(652), Index
        Case "/msg" '<nick> <msg>
            If UBound(CInfo) > 1 Then
                frmMain.SendData CInfo(1) & " " & Right(Msg, Len(Msg) - InStr(6, Msg, " ")), Chr(205) & Chr(0), Index
                frmMain.AddText "*You tell " & CInfo(1) & " " & Right(Msg, Len(Msg) - InStr(6, Msg, " ")), vbBlue
            Else
                frmMain.AddText "*Proper Syntax /msg user msg", vbRed
            End If
        Case "/me" '<channel> <msg>
            If UBound(CInfo) > 0 Then
                frmMain.SendData Chan & " """ & Right(Msg, Len(Msg) - InStr(3, Msg, " ")) & """", Chr(56) & Chr(3), Index
            Else
                frmMain.AddText "*Proper Syntax /me msg", vbRed
            End If
        Case "/setpassword" '<user> <password> [reason]
            If UBound(CInfo) > 1 Then
                '
            Else
                frmMain.AddText "*Proper Syntax /setpassword user password [reson]", vbRed
            End If
        Case "/ping" '<user>
            If UBound(CInfo) > 0 Then
                frmMain.SendData CInfo(1), Chr(239) & Chr(2), Index 'Ping 751 = 256 * (2) + (239) = 2EF
                frmMain.AddText "*" & Now & " " & CInfo(1) & " Ping...", vbBlue
            Else
                frmMain.AddText "*Proper Syntax /ping user", vbRed
            End If
        Case "/join" 'channel
            If UBound(CInfo) > 0 Then
                frmMain.SendData CInfo(1), Chr(144) & Chr(1), Index
            Else
                frmMain.AddText "*Proper Syntax /Join Channel", vbRed
            End If
        Case Else
            IsCommand = False
            Exit Function
    End Select
    IsCommand = True
End Function
