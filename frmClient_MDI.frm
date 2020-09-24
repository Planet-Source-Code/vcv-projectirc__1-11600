VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm Client 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "projectIRC"
   ClientHeight    =   5655
   ClientLeft      =   7230
   ClientTop       =   1980
   ClientWidth     =   9345
   Icon            =   "frmClient_MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picToolMain 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   0
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   623
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9345
   End
   Begin MSWinsockLib.Winsock IDENT 
      Left            =   360
      Top             =   1095
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   113
      LocalPort       =   113
   End
   Begin MSWinsockLib.Winsock sock 
      Left            =   360
      Top             =   675
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnu_file 
      Caption         =   "&File"
      Begin VB.Menu mnu_File_Connect 
         Caption         =   "&Connect"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu_File_Disconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu_File_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Options 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_File_LB02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Quit 
         Caption         =   "&Quit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "&View"
      Begin VB.Menu mnu_View_Status 
         Caption         =   "&Status Window"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnu_Window 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnu_Window_Cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnu_Window_TileH 
         Caption         =   "&Tile Horizontally"
      End
      Begin VB.Menu mnu_Tile_Vertically 
         Caption         =   "&Tile Vertically"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_Help_About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const clrSep = &H808080



Sub CTCPReply(strNick As String, strReply As String)
    Client.SendData "NOTICE " & strNick & " :" & strAction & strReply & strAction
End Sub

Public Sub HandleCTCP(strNick As String, strData As String)
    strData = RightR(strData, 1)
    strData = LeftR(strData, 1)
    
    Dim strCom As String, strParam As String
    Seperate strData, " ", strCom, strParam
    
    Select Case LCase(strCom)
        Case "version"
            PutData Status.DataIn, strColor & "05" & strBold & strNick & strBold & " has just requested your client version"
            CTCPReply strNick, "VERSION projectIRC for Windows"
        Case "ping"
            PutData Status.DataIn, strColor & "05" & strBold & strNick & strBold & " has just pinged you"
            CTCPReply strNick, "PING 0 seconds, cause projectIRC is elite"
    End Select
End Sub


Sub interpret(strData As String)
    Dim parsed As ParsedData, AllParams As String, intTemp As Integer
    Dim i As Integer, strChan As String, strTemp As String
    
    strData = Replace(strData, Chr(13), "")
    strData = Replace(strData, Chr(10), "")
    ParseData Replace(strData, Chr(13), ""), parsed
    AllParams = Params(parsed, 1, -1)
    If parsed.strCommand = "" Then Exit Sub
    
    'PutData Status.DataIn, "*****" & parsed.strNick & "~" & parsed.strCommand & "~" & AllParams
    
    Select Case LCase(parsed.strCommand)
        Case "ping"
            SendData "PONG :" & AllParams
            PutData Status.DataIn, strColor & "03Ping? Pong! [" & AllParams & "]"
            Exit Sub
        Case "join"
            If parsed.strNick = strMyNick Then
                intTemp = NewChannel(AllParams)
                Client.SendData "MODE " & AllParams
            Else
                intTemp = GetChanIndex(parsed.strParams(1))
                If intTemp = -1 Then Exit Sub
                Channels(intTemp).AddNick parsed.strNick
                PutData Channels(intTemp).DataIn, strColor & "03" & strBold & parsed.strNick & strBold & " has joined " & strBold & Channels(intTemp).strName
            End If
            Exit Sub
        Case "privmsg"
            strChan = Params(parsed, 1, 1)
            If Left(strChan, 1) = "#" Then  'privmsg to channel
                intTemp = GetChanIndex(strChan)
                If intTemp <> -1 Then Channels(intTemp).PutText parsed.strNick, Params(parsed, 2, -1)               '
            ElseIf parsed.strNick = strMyNick Then
                If Params(parsed, 2, 2) = strAction & "VERSION" & strAction Then    'version
                    'Client.SendData "CTCP REPLY " & strChan & " VERSION :jIRC for Windows9x"
                    Client.SendData "NOTICE " & parsed.strNick & " :VERSION projectIRC for Win32"
                End If
                GoTo msg
            
            Else    'send to query window
msg:
                strTemp = Params(parsed, 2, -1)
                If Left(strTemp, 1) = strAction Then
                    HandleCTCP parsed.strNick, strTemp
                    Exit Sub
                End If
                
                If QueryExists(parsed.strNick) Then
                    'MsgBox "exists"
                    intTemp = GetQueryIndex(parsed.strNick)
                    If intTemp = -1 Then Exit Sub
                    
                    If queries(intTemp).strHost <> parsed.strFullHost Then
                        queries(intTemp).strHost = RightOf(parsed.strFullHost, "!")
                        queries(intTemp).lblHost = RightOf(parsed.strFullHost, "!")
                        
                    End If
                    queries(intTemp).Caption = parsed.strNick
                    queries(intTemp).strNick = parsed.strNick
                    queries(intTemp).lblNick = parsed.strNick
                    queries(intTemp).PutText parsed.strNick, strTemp
                Else
                    'MsgBox "doesnt"
                    NewQuery parsed.strNick, parsed.strFullHost
                    intTemp = GetQueryIndex(parsed.strNick)
                    If intTemp = -1 Then Exit Sub
                    queries(intTemp).Caption = parsed.strNick
                    queries(intTemp).strNick = parsed.strNick
                    queries(intTemp).lblNick = parsed.strNick
                    queries(intTemp).PutText parsed.strNick, strTemp
                End If
            End If
            Exit Sub
        Case "nick"
            If parsed.strNick = strMyNick Then
                strMyNick = Params(parsed, 1, 1)
                PutData Status.DataIn, strColor & "03Your nick is now " & strBold & strMyNick
                ChangeNick parsed.strNick, Params(parsed, 1, -1)
            Else
                ChangeNick parsed.strNick, Params(parsed, 1, 1)
            End If
            Exit Sub
        Case "part"
            If parsed.strNick = strMyNick Then Exit Sub
            intTemp = GetChanIndex(parsed.strParams(1))
            'MsgBox intTemp & "~" & parsed.strParams(1) & "~"
            If intTemp = -1 Then Exit Sub
            Channels(intTemp).RemoveNick parsed.strNick
            PutData Channels(intTemp).DataIn, strColor & "03" & strBold & parsed.strNick & strBold & " has left " & strBold & Channels(intTemp).strName
            If parsed.strNick = strMyNick Then Unload Channels(intTemp)
            Exit Sub
        Case "353" 'nick list!
            'MsgBox parsed.strParams(3)
            intTemp = GetChanIndex(parsed.strParams(3))
            If intTemp = -1 Then Exit Sub
            Dim strNicks() As String
            strNicks = Split(Params(parsed, 4, -1), " ")
            For i = LBound(strNicks) To UBound(strNicks)
                Channels(intTemp).AddNick strNicks(i)
            Next i
            Exit Sub
        Case "mode"     'set mode
            intTemp = GetChanIndex(Params(parsed, 1, 1))
            If intTemp = -1 Then Exit Sub
            PutData Channels(intTemp).DataIn, strColor & "03" & strBold & parsed.strNick & strBold & " sets mode: " & Params(parsed, 2, -1)
            ParseMode Params(parsed, 1, 1), Params(parsed, 2, -1)
            Exit Sub
        Case "quit"
            NickQuit parsed.strNick, Params(parsed, 1, -1)
            Exit Sub
        Case "kick"
            intTemp = GetChanIndex(Params(parsed, 1, 1))
            If intTemp = -1 Then Exit Sub
            PutData Channels(intTemp).DataIn, strColor & "03" & strBold & Params(parsed, 2, 2) & strBold & " was kicked from " & strBold & Params(parsed, 1, 1) & strBold & " by " & strBold & parsed.strNick & strBold & " [ " & Params(parsed, 3, -1) & " ]"
            Channels(intTemp).RemoveNick Params(parsed, 2, 2)
            Exit Sub
        Case "332"  'topic!
            intTemp = GetChanIndex(Params(parsed, 2, 2))
            If intTemp = -1 Then Exit Sub
            Channels(intTemp).rtbTopic.Text = ""
            PutData Channels(intTemp).rtbTopic, Params(parsed, 3, -1)
            Channels(intTemp).rtbTopic.SelStart = 0
            Channels(intTemp).rtbTopic.SelLength = 1
            Channels(intTemp).rtbTopic.SelText = ""
            PutData Channels(intTemp).DataIn, strColor & "03Topic is """ & strColor & Params(parsed, 3, -1) & strColor & "03"""
            Channels(intTemp).rtbTopic.SelStart = 0
            Channels(intTemp).rtbTopic.Tag = "locked"
            Exit Sub
        Case "topic"    'change in topic!
            intTemp = GetChanIndex(Params(parsed, 1, 1))
            If intTemp = -1 Then Exit Sub
            Channels(intTemp).rtbTopic.Text = ""
            PutData Channels(intTemp).rtbTopic, Params(parsed, 2, -1)
            Channels(intTemp).rtbTopic.SelStart = 0
            Channels(intTemp).rtbTopic.SelLength = 1
            Channels(intTemp).rtbTopic.SelText = ""
            PutData Channels(intTemp).DataIn, strColor & "03Topic changed by " & strBold & parsed.strNick & strBold & " : " & Params(parsed, 2, -1)
            Exit Sub
        Case "333"  'topic on param2 set by param3, on param4
            intTemp = GetChanIndex(Params(parsed, 2, 2))
            If intTemp = -1 Then Exit Sub
            PutData Channels(intTemp).DataIn, strColor & "03Topic set by " & strBold & Params(parsed, 3, 3) & strBold
        Case "324"  'set channel modes
            ParseMode Params(parsed, 2, 2), Params(parsed, 3, -1)
    End Select
    PutData Status.DataIn, "*** " & strBold & parsed.strCommand & strBold & " " & AllParams ' & " [" & parsed.strFullHost & "]"
End Sub


Sub SendData(strData As String)
    On Error Resume Next
    sock.SendData strData & Chr(10)
End Sub


Private Sub IDENT_ConnectionRequest(ByVal requestID As Long)
    IDENT.Close
    IDENT.Accept requestID
End Sub

Private Sub IDENT_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String
    IDENT.GetData dat, vbString
    
    If dat Like "*, *" Then
        dat = LeftR(dat, 2)
        PutData Status.DataIn, "*** IDENT : " & dat
        dat = dat & " : USERID : UNIX : " & strMyIdent
        Client.SendData dat
        PutData Status.DataIn, "*** IDENT reply : " & dat
        'MsgBox "~" & dat & "~"
        Dim i As Integer
        IDENT.Close
    End If
End Sub

Private Sub IDENT_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    PutData Status.DataIn, Chr(Color) & "04IDENT Error " & strColor & Description
End Sub

Private Sub MDIForm_Load()
    '* Use this until setting files implemented
    strServer = "irc.otherside.com"
    strMyNick = "YourNick"
    strOtherNick = "OtherNick"
    strFullName = "jIRC User"
    strMyIdent = "jIRC"
    lngPort = 6667
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Client.SendData "QUIT :using projectIRC, closed"
    Dim i As Integer
    For i = 0 To 50
        DoEvents
    Next i
    Cancel = 0
End Sub

Private Sub MDIForm_Terminate()
    Client.SendData "QUIT :Client closed, using jIRC"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Client.SendData "QUIT :Client closed, using jIRC"
End Sub

Sub mnu_File_Connect_Click()
    Select Case mnu_File_Connect.Caption
        Case "&Connect"
            '* Connect
            sock.Close
            mnu_File_Connect.Caption = "&Cancel"
            sock.RemoteHost = strServer
            sock.RemotePort = lngPort
            sock.Connect
            IDENT.Close
            On Error Resume Next
            IDENT.Listen
            PutData Status.DataIn, strColor & "02Connecting to " & strBold & strServer & strBold & " port " & strBold & lngPort
        Case "&Cancel"
            '* Cancel
            IDENT.Close
            sock.Close
            mnu_File_Connect.Caption = "&Connect"
            PutData Status.DataIn, strColor & "05Connection attempt cancelled"
    End Select
End Sub


Sub mnu_File_Disconnect_Click()
    sock.Close
    mnu_File_Connect.Enabled = True
    mnu_File_Disconnect.Enabled = False
    PutData Status.DataIn, strColor & "05Disconnected from " & strServer
End Sub


Private Sub mnu_File_Options_Click()
    Options.Show 1
End Sub

Private Sub mnu_File_Quit_Click()
    Client.SendData "QUIT :Using projectIRC, closed"
    IDENT.Close
    sock.Close
    Dim i As Integer
    For i = 1 To 1000
        DoEvents
    Next i
    Unload Me
End Sub

Private Sub mnu_Help_About_Click()
    About.Show vbModal
End Sub

Private Sub mnu_Tile_Vertically_Click()
    Client.Arrange vbTileVertical
End Sub

Private Sub mnu_View_Status_Click()
    mnu_View_Status.Checked = Not mnu_View_Status.Checked
    Status.Visible = mnu_View_Status.Checked
End Sub

Private Sub mnu_Window_Cascade_Click()
    Client.Arrange vbCascade
End Sub

Private Sub mnu_Window_TileH_Click()
    Client.Arrange vbTileHorizontal
End Sub


Private Sub sock_Close()
    PutData Status.DataIn, strColor & "02Disconnected by SERVER from " & strServer
    sock.Close
    IDENT.Close
    mnu_File_Connect.Enabled = True
    mnu_File_Connect.Caption = "&Connect"
    mnu_File_Disconnect.Enabled = False

End Sub

Private Sub sock_Connect()
    mnu_File_Connect.Enabled = False
    mnu_File_Disconnect.Enabled = True
    mnu_File_Connect.Caption = "&Connect"
    PutData Status.DataIn, strColor & "03Connected to " & strServer
    
    SendData "PASS password"
    SendData "NICK " & strMyNick
    SendData "USER " & strMyNick & " " & sock.LocalHostName & " irc :" & strFullName
    
    '* Let's close all open windows
    Dim i As Integer
    
    For i = 1 To intChannels
        Channels(i).Tag = "NOPART"
        Unload Channels(i)
    Next i
    
    For i = 1 To intQueries
        Unload queries(i)
    Next i
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String, AllParams As String
    Dim strData() As String, i As Integer
    
    sock.GetData dat, vbString
    
    '* this'll stay for about half a second
    strData = Split(dat, Chr(10))
    
    For i = LBound(strData) To UBound(strData)
        interpret strData(i)
    Next i
    
End Sub


Private Sub sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    PutData Status.DataIn, strColor & "04ERROR : " & Description
End Sub


