Attribute VB_Name = "modIRC"
'* projectIRC version 1.0
'* By Matt C, sappy@adelphia.net
'* Feel free to EMail me with any questions you may have.

'* Module that handles many of the Client's procedures,
'* Feel free to use in your project as long as you give me proper credit

'/me is using projectIRC $version $+ , on $server port $port me = $me on channel $chan

'* Channels and Queries
Global Const MAX_CHANNELS = 30
Global Const MAX_QUERIES = 30
Public Channels(1 To MAX_CHANNELS)  As Channel
Public queries(1 To MAX_QUERIES)    As Query
Public intChannels  As Integer
Public intQueries   As Integer

'* Server settings
Global strServer    As String
Global strMyNick    As String
Global strOtherNick As String
Global strFullName  As String
Global strMyIdent   As String
Global lngPort      As Long

'* Variables for incoming commands
Type ParsedData
    bHasPrefix   As Boolean
    strParams()  As String
    intParams    As Integer
    strFullHost  As String
    strCommand   As String
    strNick      As String
    strIdent     As String
    strHost      As String
    AllParams    As String
End Type

'* ANSI Formatting character values
Global Const BOLD = 2
Global Const UNDERLINE = 31
Global Const Color = 3
Global Const REVERSE = 22
Global Const ACTION = 1

'* ANSI Formatting characters
Global strBold As String
Global strUnderline As String
Global strColor As String
Global strReverse As String
Global strAction As String

'Nick storage for nick list inchannels
Type Nick
    Nick    As String
    op      As Boolean
    voice   As Boolean
    helper  As Boolean
    host    As String
    IDENT   As String
End Type

'Mode storage for each channel
Type typMode
    mode    As String
    bPos    As Boolean
End Type

Function AnsiColor(intColNum As Integer) As Long
    Select Case intColNum
        Case 0: AnsiColor = RGB(255, 255, 255)
        Case 1: AnsiColor = RGB(0, 0, 0)
        Case 2: AnsiColor = RGB(0, 0, 127)
        Case 3: AnsiColor = RGB(0, 127, 0)
        Case 4: AnsiColor = RGB(255, 0, 0)
        Case 5: AnsiColor = RGB(127, 0, 0)
        Case 6: AnsiColor = RGB(127, 0, 127)
        Case 7: AnsiColor = RGB(255, 127, 0)
        Case 8: AnsiColor = RGB(255, 255, 0)
        Case 9: AnsiColor = RGB(0, 255, 0)
        Case 10: AnsiColor = RGB(0, 0, 0)
        Case 11: AnsiColor = RGB(0, 255, 255)
        Case 12: AnsiColor = RGB(0, 0, 255)
        Case 13: AnsiColor = RGB(255, 0, 255)
        Case 14: AnsiColor = RGB(92, 92, 92)
        Case 15: AnsiColor = RGB(184, 184, 184)
        Case Else: AnsiColor = RGB(0, 0, 0)
    End Select
End Function



Sub ChangeNick(strOldNick As String, strNewNick As String)
    Dim i As Integer, bChangedQuery As Boolean, intTemp As Integer
    
    'MsgBox intChannels
    For i = 1 To intChannels
        'change in queries :)
        If Not bChangedQuery Then
            If Channels(i).InChannel(strOldNick) Then
               
                intTemp = GetQueryIndex(strOldNick)
                If intTemp <> -1 Then
                    queries(intTemp).lblNick = strNewNick
                    queries(intTemp).strNick = strNewNick
                    queries(intTemp).Caption = strNewNick
                    'Queries(intTemp).strHost = RightOf(parsed.strFullHost, "!")
                    'Queries(intTemp).lblHost = RightOf(parsed.strFullHost, "!")
                End If
                bChangedQuery = True
            End If
            
        'change in channel :)
        If Channels(i).strName <> "" Then
            Channels(i).ChangeNck strOldNick, strNewNick
        End If
        End If
    Next i
End Sub

Function Combine(arrItems() As String, intStart As Integer, intEnd As Integer) As String
    '* This returns the given parameter range as one string
    '* -1 specified as strEnd means from intStart to the last parameter
    '* Ex: strParams(1) is "#mIRC", then params 2 thru 16 are the nicks
    '*     of the users in the channel, simply do something like this
    '*     strNames = Params(2, -1)
    '*     that would return all the nicks into one string, similar to mIRC's
    '*      $#- identifier ($2-) in this case
    
    '* Declare variables
    Dim strFinal As String, intLast As Integer, i As Integer
    
    '* check for bad parameters
    If intStart < 1 Or intEnd > UBound(arrItems) + 1 Then Exit Function
    
    '* if intEnd = -1 then set it to param count
    If intEnd = -1 Then intLast = UBound(arrItems) + 1 Else intLast = intEnd
        
    For i = intStart To intLast
        strFinal = strFinal & arrItems(i - 1)
        If i <> intLast Then strFinal = strFinal & " "
    Next i
    
    Combine = Replace(strFinal, vbCr, "")
    strFinal = ""
End Function

Function DisplayNick(nckNick As Nick) As String
    Dim strPre As String
    If nckNick.voice Then strPre = "+"
    If nckNick.helper Then strPre = "%"
    If nckNick.op Then strPre = "@"
    DisplayNick = strPre & nckNick.Nick
End Function

Sub DoMode(strChannel As String, bAdd As Boolean, strMode As String, strParam As String)
    Dim intX As Integer, i As Integer
    intX = GetChanIndex(strChannel)
    
    Select Case strMode
        Case "v"
            Channels(intX).SetVoice strParam, bAdd
            
        Case "o"
            Channels(intX).SetOp strParam, bAdd
            If bAdd Then
                If strParam = strMyNick Then Channels(intX).rtbTopic.Tag = ""
            Else
                If strParam = strMyNick Then Channels(intX).rtbTopic.Tag = "locked"
            End If
        Case "h"
        Case "b"
        Case "k"
            If bAdd = True Then
                Channels(intX).strKey = strParam
                Channels(intX).AddMode strMode, bAdd
            Else
                Channels(intX).strKey = ""
                Channels(intX).RemoveMode strMode
            End If
        Case "l"
            If bAdd = True Then
                Channels(intX).intLimit = CInt(strParam)
                Channels(intX).AddMode strMode, bAdd
            Else
                Channels(intX).intLimit = 0
                Channels(intX).RemoveMode strMode
            End If
        Case Else
            If bAdd = True Then
                Channels(intX).AddMode strMode, bAdd
            Else
                Channels(intX).RemoveMode strMode
            End If
    End Select
End Sub

Function GetAlias(strChan As String, strData As String) As String
    Dim arrParams() As String, i As Integer, strP As String, strCom As String
    Dim strFinal As String, strAdd As String, bSpace As Boolean, intTemp As Integer
    Dim strTemp As String, strNck As String
    
    Seperate strData, " ", strCom, strData
    arrParams = Split(strData, " ")
    bSpace = True
    
    For i = LBound(arrParams) To UBound(arrParams)
        strP = arrParams(i)
        'MsgBox strP & ":" & i
        strAdd = ""
        If strP = "$+" Then
            strFinal = LeftR(strFinal, 1)
            bSpace = False
        ElseIf Left(strP, 1) = "$" Then
            strAdd = GetVar(strChan, RightR(strP, 1))
        Else
            strAdd = strP
        End If
        
        strFinal = strFinal & strAdd
        If bSpace Then
            strFinal = strFinal & " "
        Else
            bSpace = True
        End If
    Next i
    
    DoEvents
    
    If Len(strFinal) > 0 Then strFinal = LeftR(strFinal, 1)
    
    ReDim arrParams(1) As String
    arrParams = Split(strFinal, " ")
    
    Dim r As String 'return
    Select Case LCase(strCom)
        Case "query"
            strTemp = Combine(arrParams, 2, -1)
            strNck = Combine(arrParams, 1, 1)
            'MsgBox strNck & "~"
            If QueryExists(strNck) Then
                intTemp = GetQueryIndex(strNck)
                queries(intTemp).PutText strMyNick, strTemp
                r = "PRIVMSG " & strNck & " :" & strTemp
            Else
                intTemp = NewQuery(strNck, "")
                If UBound(arrParams) > 0 Then
                    queries(intTemp).PutText strMyNick, strTemp
                    r = "PRIVMSG " & strNck & " :" & strTemp
                End If
            End If
                
        Case "msg"
            r = "PRIVMSG " & Combine(arrParams, 1, 1) & " :" & Combine(arrParams, 2, -1)
        Case "me"
            strTemp = Combine(arrParams, 1, -1)
            r = "PRIVMSG " & strChan & " :" & strAction & "ACTION " & strTemp & strAction
            If Left(strChan, 1) = "#" Then
                intTemp = GetChanIndex(strChan)
                If intTemp = -1 Then Exit Function
                PutData Channels(intTemp).DataIn, strColor & "06" & strMyNick & " " & strTemp
            Else
                intTemp = GetQueryIndex(strChan)
                If intTemp = -1 Then Exit Function
                PutData queries(intTemp).DataIn, strColor & "06" & strMyNick & " " & strTemp
            End If
        Case "quit"
            r = "QUIT :" & Combine(arrParams, 1, -1)
        Case "raw"
            r = Combine(arrParams, 1, -1)
        Case "nick"
            If Client.sock.State = 0 Then
                strMyNick = Combine(arrParams, 1, 1)
            Else
                r = "NICK " & Combine(arrParams, 1, 1)
            End If
        Case "id"   'identify with nickserv
            r = "PRIVMSG NickServ :IDENTIFY " & Combine(arrParams, 1, 1)
        Case "part"
            strTemp = Combine(arrParams, 1, -1)
            If UBound(arrParams) = 0 Then
                r = "PART " & strChan
                strTemp = strTemp
            Else
                r = "PART " & Combine(arrParams, 1, 1) & " :" & Combine(arrParams, 2, -1)
                strTemp = LeftOf(strTemp, " ")
            End If
            
            intTemp = GetChanIndex(strTemp)
            
            'On Error Resume Next
            Channels(intTemp).Tag = "PARTNOW"
            
'            MsgBox r
        Case Else
            r = strCom & " " & Combine(arrParams, 1, -1)
    End Select
    
    GetAlias = r
    
End Function


Function GetChanIndex(strName As String) As Integer
    Dim i As Integer

    For i = 1 To intChannels
        If LCase(Channels(i).strName) = LCase(strName) Then
            GetChanIndex = i
            Exit Function
        End If
    Next i
    GetChanIndex = -1
End Function
Function GetQueryIndex(strNick As String) As Integer
    Dim i As Integer

    For i = 1 To intQueries
        If LCase(queries(i).strNick) = LCase(strNick) Then
            GetQueryIndex = i
            Exit Function
        End If
    Next i
    GetQueryIndex = -1
End Function

Function GetVar(strChan As String, strName As String)
    Dim r As String     'r is the return value
    Dim intTemp As String
    
    On Error Resume Next
    Select Case LCase(strName)
        Case "version"
            r = App.Major & "." & App.Minor & App.Revision
        Case "chan", "channel", "ch"
            r = strChan
        Case "me"
            r = strMyNick
        Case "server"
            r = Client.sock.RemoteHost
        Case "port"
            r = Client.sock.RemotePort
        Case "randnick"
            intTemp = GetChanIndex(strChan)
            If Left(strChan, "1") = "#" Then
                With Channels(intTemp)
                    Randomize
                    r = .GetNick(Int(Rnd * .intNicks) + 1)
                End With
            End If
    End Select
    GetVar = r
End Function

Function LeftOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    
    intPos = InStr(strData, strDelim)
    If intPos Then
        LeftOf = Left(strData, intPos - 1)
    Else
        LeftOf = strData
    End If
End Function

Function LeftR(strData As String, intMin As Integer)
    LeftR = Left(strData, Len(strData) - intMin)
End Function

Function NewChannel(strName As String) As Integer
    Dim i As Integer

    For i = 1 To intChannels
        If Channels(i).strName = "" Then
            Channels(i).Caption = strName
            Channels(i).lblName = strName
            Channels(i).strName = strName
            Channels(i).Visible = True
            Channels(i).Tag = i
            NewChannel = i
            Exit Function
        End If
    Next i
    intChannels = intChannels + 1
    Set Channels(intChannels) = New Channel
    Channels(intChannels).strName = strName
    Channels(intChannels).lblName = strName
    Channels(intChannels).Caption = strName
    Channels(intChannels).Visible = True
    Channels(intChannels).Tag = intChannels
    NewChannel = intChannels
End Function
Function NewQuery(strNick As String, strHost As String) As Integer
    Dim i As Integer, strHostX As String
    strHostX = RightOf(strHost, "!")

    For i = 1 To intQueries
        If queries(i).strNick = "" Then
            queries(i).Caption = strNick
            queries(i).lblNick = strNick
            queries(i).strNick = strNick
            queries(i).strHost = strHostX
            queries(i).lblHost = strHostX
            queries(i).Visible = True
            queries(i).Tag = i
            NewQuery = i
            Exit Function
        End If
    Next i
    
    intQueries = intQueries + 1
    Set queries(intQueries) = New Query
    queries(intQueries).strNick = strNick
    queries(intQueries).lblNick = strNick
    queries(intQueries).Caption = strNick
    queries(intQueries).lblHost = strHostX
    queries(intQueries).strHost = strHostX
    queries(intQueries).Visible = True
    queries(intQueries).Tag = intQueries
    NewQuery = intQueries
End Function

Sub NickQuit(strNick As String, strMsg As String)
    For i = 1 To intChannels
        If Channels(i).InChannel(strNick) Then
            Channels(i).RemoveNick strNick
            PutData Channels(i).DataIn, strColor & "02" & strBold & strNick & strBold & " has Quit IRC [ " & strMsg & " ]"
            Exit For
        End If
    Next i

    For i = 1 To intQueries
        If LCase(queries(i).strNick) = LCase(strNick) Then
            PutData queries(i).DataIn, strColor & "02" & strBold & strNick & strBold & " has Quit IRC [ " & strMsg & " ]"
            Exit Sub
        End If
    Next i
End Sub

Function Params(parsed As ParsedData, intStart As Integer, intEnd As Integer) As String
    '* This returns the given parameter range as one string
    '* -1 specified as strEnd means from intStart to the last parameter
    '* Ex: strParams(1) is "#mIRC", then params 2 thru 16 are the nicks
    '*     of the users in the channel, simply do something like this
    '*     strNames = Params(2, -1)
    '*     that would return all the nicks into one string, similar to mIRC's
    '*      $#- identifier ($2-) in this case
    
    '* Declare variables
    Dim strFinal As String, intLast As Integer, i As Integer
    
    '* check for bad parameters
    If intStart < 1 Or intEnd > parsed.intParams Then Exit Function
    
    '* if intEnd = -1 then set it to param count
    If intEnd = -1 Then intLast = parsed.intParams Else intLast = intEnd
        
    For i = intStart To intLast
        strFinal = strFinal & parsed.strParams(i)
        If i <> intLast Then strFinal = strFinal & " "
    Next i
    
    Params = Replace(strFinal, vbCr, "")
    strFinal = ""
End Function

Sub ParseData(ByVal strData As String, ByRef parsed As ParsedData)

    '*/ Here's what this does:
    '*/ : 1. Resets all the variables from previous parsing
    '*/ : 2. Checks for host prefix indicated by :
    '*/ :   a. parses data left of space into nick, ident and host (or just host in some cases)
    '*/ : 3. Checks for any parameters
    '*/ :   a. if no parameters, exit
    '*/ :   b. If param starts with :, indicates last parameter
    '*/ :   c. If space found left in string, more parameters, go back to b
    '*/ :   d. if no space, exit

    '* Declare variables
    Dim strTMP As String, i As Integer
    
    '* Reset variables
    bHasPrefix = False
    parsed.strNick = ""
    parsed.strIdent = ""
    parsed.strHost = ""
    parsed.strCommand = ""
    parsed.intParams = 1
    ReDim parsed.strParams(1 To 1) As String
    
    '* Check for prefix, if so, parse nick, ident and host (or just host)
    If Left(strData, 1) = ":" Then
        bHasPrefix = True
        strData = Right(strData, Len(strData) - 1)
        '* Put data left of " " in strHost, data right of " "
        '* into strData
        Seperate strData, " ", parsed.strHost, strData
        parsed.strFullHost = parsed.strHost
        
        '* Check to see if client host name
        If InStr(parsed.strHost, "!") Then
            Seperate parsed.strHost, "!", parsed.strNick, parsed.strHost
            Seperate parsed.strHost, "@", parsed.strIdent, parsed.strHost
        End If
    End If
    
    '* If any params, parse
    If InStr(strData, " ") Then
        Seperate strData, " ", parsed.strCommand, strData
        
        parsed.AllParams = strData
       '* Let's parse all the parameters.. yummy
Begin: '* OH NO I USED A LABEL!

        '* If begginning of param is :, indicates that its the last param
        If Left(strData, 1) = ":" Then
            parsed.strParams(parsed.intParams) = Right(strData, Len(strData) - 1)
            GoTo Finish
        End If
        '* If there is a space still, there is more params
        If InStr(strData, " ") Then
            Seperate strData, " ", parsed.strParams(parsed.intParams), strData
            'If Left(parsed.strParams(1), 1) = "#" Then MsgBox parsed.strParams(parsed.intParams) & "~~"
            parsed.intParams = parsed.intParams + 1
            ReDim Preserve parsed.strParams(1 To parsed.intParams) As String
            GoTo Begin
        Else
            parsed.strParams(parsed.intParams) = strData
        End If
    Else
        '* No params, strictly command
        parsed.intParams = 0
        parsed.strCommand = strData
    End If
Finish:
End Sub

Sub ParseMode(strChannel As String, strData As String)
    Dim strModes() As String, strChar As String
    Dim i As Integer, intParam As Integer
    Dim bAdd As Boolean
    
    bAdd = True
    strModes = Split(strData, " ")
    For i = 1 To Len(strModes(0))
        strChar = Mid(strModes(0), i, 1)
        Select Case strChar
            Case "+"
                bAdd = True
            Case "-"
                bAdd = False
            Case "v", "b", "o", "h", "k", "l"
                intParam = intParam + 1
                DoMode strChannel, bAdd, strChar, strModes(intParam)
            Case Else
                DoMode strChannel, bAdd, strChar, ""
        End Select
    Next i
End Sub

Sub PutData(RTF As RichTextBox, strData As String)
    Dim bBold As Boolean, bUnderLine As Boolean, bReverse As Boolean, bColor As Boolean
    Dim strBuffer As String, strChar As String, intAsc As Integer, i As Integer
    Dim inColor As Boolean, strCols() As String
    
    RTF.SelStart = Len(RTF.Text)
    RTF.SelText = " "
    For i = 1 To Len(strData)
        strChar = Mid(strData, i, 1)
        intAsc = Asc(strChar)
        
        If intAsc = Color Then
            bColor = Not bColor
            strBuffer = ""
        End If
        
        RTF.SelStart = Len(RTF.Text)
        If inColor And intAsc = Color Then
            bColor = False
            strCols = Split(strBuffer, ",")
            RTF.SelColor = AnsiColor(CInt(strCols(0)))
            RTF.SelText = strChar
            strBuffer = ""
            inColor = False
        ElseIf inColor And (strChar = "," Or IsNumeric(strChar)) Then
            If Len(RightOf(strBuffer, ",")) >= 2 And strChar <> "," Then
                strCols = Split(strBuffer, ",")
                RTF.SelColor = AnsiColor(CInt(strCols(0)))
                strBuffer = ""
                RTF.SelText = strChar
                inColor = False
            ElseIf RightOf(CInt(strBuffer & strChar), ",") > 15 Then
                strCols = Split(strBuffer, ",")
                RTF.SelColor = AnsiColor(CInt(strCols(0)))
                strBuffer = ""
                RTF.SelText = strChar
                inColor = False
            Else
                strBuffer = strBuffer & strChar
            End If
        ElseIf inColor And Not IsNumeric(strChar) Then
            If strBuffer = "" Then
                bColor = False
                inColor = False
                GoTo hah
            End If
            
            strCols = Split(strBuffer, ",")
            On Error Resume Next
            RTF.SelColor = AnsiColor(CInt(strCols(0)))
            strBuffer = ""
            
            'If intAsc <> COLOR Then rtf.SelText = strChar
            
            If intAsc = BOLD Then
                bBold = Not bBold
                RTF.SelBold = bBold
            ElseIf intAsc = UNDERLINE Then
                bUnderLine = Not bUnderLine
                RTF.SelUnderline = bUnderLine
            ElseIf intAsc = REVERSE Then
            Else
                RTF.SelText = strChar
            End If
            
            inColor = False
hah:
        ElseIf intAsc = Color Then
            If Not bColor Then
                RTF.SelColor = vbBlack
            Else
                inColor = True
            End If
        ElseIf intAsc = BOLD Then
            bBold = Not bBold
            RTF.SelBold = bBold
        ElseIf intAsc = UNDERLINE Then
            bUnderLine = Not bUnderLine
            RTF.SelUnderline = bUnderLine
        Else
            RTF.SelText = strChar
        End If
    Next i
    RTF.SelBold = False
    RTF.SelColor = vbBlack
    RTF.SelUnderline = False
    RTF.SelStart = Len(RTF.Text)
    RTF.SelText = vbCrLf
End Sub

Function QueryExists(strNick As String) As Boolean
    Dim i As Integer

    For i = 1 To intQueries
        If LCase(queries(i).strNick) = LCase(strNick) Then
            QueryExists = True
            Exit Function
        End If
    Next i
    QueryExists = False
End Function

Function RealNick(strNick As String) As String
    strNick = Replace(strNick, "@", "")
    strNick = Replace(strNick, "%", "")
    strNick = Replace(strNick, "+", "")
    RealNick = strNick
End Function

Sub RefreshList(lstBox As ListBox)
    'lstBox.AddItem "", 0
    'lstBox.RemoveItem 0
End Sub

Function RightOf(strData As String, strDelim As String) As String
    Dim intPos As Integer
    intPos = InStr(strData, strDelim)
    
    If intPos Then
        RightOf = Mid(strData, intPos + 1, Len(strData) - intPos)
    Else
        RightOf = strData
    End If
End Function


Function RightR(strData As String, intMin As Integer)
    On Error Resume Next
    RightR = Right(strData, Len(strData) - intMin)
End Function

Sub Seperate(strData As String, strDelim As String, ByRef strLeft As String, ByRef strRight As String)
    '* Seperates strData into 2 variables based on strDelim
    '* Ex: strData is "Bill Clinton"
    '*     Dim strFirstName As String, strLastName As String
    '*     Seperate strData, " ", strFirstName, strLastName
    
    Dim intPos As Integer
    intPos = InStr(strData, strDelim)
    
    If intPos Then
        strLeft = Left(strData, intPos - 1)
        strRight = Mid(strData, intPos + 1, Len(strData) - intPos)
    Else
        strLeft = strData
        strRight = strData
    End If
End Sub


