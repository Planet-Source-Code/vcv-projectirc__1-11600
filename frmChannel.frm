VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Channel 
   Caption         =   "#channel"
   ClientHeight    =   3870
   ClientLeft      =   4380
   ClientTop       =   3030
   ClientWidth     =   6690
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   6690
   Begin VB.PictureBox picTopic 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1770
      ScaleHeight     =   255
      ScaleWidth      =   4725
      TabIndex        =   8
      Top             =   105
      Width           =   4725
      Begin RichTextLib.RichTextBox rtbTopic 
         Height          =   375
         Left            =   -45
         TabIndex        =   9
         Top             =   -45
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   661
         _Version        =   393217
         MultiLine       =   0   'False
         MaxLength       =   512
         TextRTF         =   $"frmChannel.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBMPC"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picFlat 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   165
      ScaleHeight     =   3195
      ScaleWidth      =   6330
      TabIndex        =   1
      Top             =   390
      Width           =   6330
      Begin VB.PictureBox picNicks 
         BorderStyle     =   0  'None
         Height          =   2925
         Left            =   4680
         ScaleHeight     =   2925
         ScaleWidth      =   1650
         TabIndex        =   6
         Top             =   0
         Width           =   1650
         Begin VB.ListBox lstNicks 
            Height          =   3000
            IntegralHeight  =   0   'False
            Left            =   -30
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   7
            Top             =   -45
            Width           =   1710
         End
      End
      Begin VB.PictureBox picDO 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   6360
         TabIndex        =   2
         Top             =   2955
         Width           =   6360
         Begin RichTextLib.RichTextBox DataOut 
            Height          =   390
            Left            =   -45
            TabIndex        =   3
            Top             =   -45
            Width           =   6510
            _ExtentX        =   11483
            _ExtentY        =   688
            _Version        =   393217
            MultiLine       =   0   'False
            TextRTF         =   $"frmChannel.frx":00FA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBMPC"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin RichTextLib.RichTextBox DataIn 
         Height          =   2925
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4655
         _ExtentX        =   8202
         _ExtentY        =   5159
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmChannel.frx":01F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBMPC"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4215
      Top             =   3465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":02EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChannel.frx":0742
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   330
      Left            =   1440
      TabIndex        =   0
      Top             =   3765
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Shape shpTopic 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   300
      Left            =   1740
      Top             =   90
      Width           =   4785
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "#channel name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   90
      TabIndex        =   5
      Top             =   75
      Width           =   1425
   End
   Begin VB.Shape shpDI 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   3240
      Left            =   1650
      Top             =   375
      Width           =   4875
   End
   Begin VB.Shape shpBlue 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3930
      Left            =   0
      Top             =   0
      Width           =   1650
   End
End
Attribute VB_Name = "Channel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strTopic As String
Public strMode  As String
Public strName  As String
Public strKey   As String
Public intLimit As Integer

Dim Nicks()     As Nick
Public intNicks As Integer

Public bControl As Boolean

Dim Modes()     As typMode
Public intModes As Integer

Public Sub AddMode(strMode As String, bPlus As Boolean)
'    MsgBox strMode & "~" & bPlus
    Dim i As Integer
    For i = 1 To intModes
        If Modes(i).mode = strMode Then Exit Sub
    Next i
    
    intModes = intModes + 1
    ReDim Preserve Modes(1 To intModes) As typMode
    
    With Modes(intModes)
        .bPos = True
        .mode = strMode
    End With
    Update
End Sub


Public Sub AddNick(strNick As String)
    Dim strPre As String

    If strNick = "" Then Exit Sub
    intNicks = intNicks + 1
    ReDim Preserve Nicks(1 To intNicks) As Nick
    If InStr(strNick, "%") Then Nicks(intNicks).helper = True: strPre = "%": strNick = Replace(strNick, "%", "")
    If InStr(strNick, "+") Then Nicks(intNicks).voice = True: strPre = "+": strNick = Replace(strNick, "+", "")
    If InStr(strNick, "@") Then Nicks(intNicks).op = True: strPre = "@": strNick = Replace(strNick, "@", "")
    
    Nicks(intNicks).Nick = strNick
    lstNicks.AddItem DisplayNick(Nicks(intNicks))
End Sub



Public Sub ChangeNck(strOldNick, strNewNick)
    Dim i As Integer, bInd As Integer
    
    For i = 1 To intNicks
        'MsgBox Nicks(i).Nick & "~" & strOldNick
        If Nicks(i).Nick = strOldNick Then
            Nicks(i).Nick = strNewNick
            bInd = i
            Exit For
        End If
    Next i
    For i = 0 To lstNicks.ListCount - 1
        'MsgBox RealNick(lstNicks.List(i)) & "!" & strOldNick
        If RealNick(lstNicks.List(i)) = strOldNick Then
            lstNicks.List(i) = DisplayNick(Nicks(bInd))
            PutData DataIn, strColor & "03" & strBold & strOldNick & strBold & " is now known as " & strBold & strNewNick
            Exit For
        End If
    Next i
End Sub




Public Function GetNick(intIndex As Integer) As String
    GetNick = Nicks(intIndex).Nick
End Function


Function InChannel(strNick As String) As Boolean
    Dim i As Integer
    For i = 1 To intNicks
        If strNick = Nicks(i).Nick Then InChannel = True: Exit Function
    Next i
    InChannel = False
End Function

Public Function ModeString() As String
    If intModes = 0 Then Exit Function
    Dim strFinal As String, bWhich As Boolean, i As Integer
    If Modes(1).bPos = True Then bWhich = True
    
    If bWhich Then strFinal = strFinal & "+" Else strFinal = strFinal & "-"
    
    For i = 1 To intModes
        If Modes(i).bPos <> bWhich Then
            bWhich = Not bWhich
            If bWhich Then strFinal = strFinal & "+" Else strFinal = strFinal & "-"
        End If
        strFinal = strFinal & Modes(i).mode
    Next i
    ModeString = strFinal
End Function

Public Sub PutText(strNick As String, strText As String)
    If Left(strText, 8) = strAction & "ACTION " Then
        strText = RightR(strText, 8)
        strText = LeftR(strText, 1)
        PutData Me.DataIn, strColor & "06" & strNick & " " & strText
    ElseIf Left(strText, 9) = strAction & "VERSION" & strAction Then
        'MsgBox "hey"
        Client.SendData "CTCPREPLY " & strNick & " VERSION :jIRC for Windows9x"
    Else
        'Dim i As Integer
        'For i = 0 To lstNicks.ListCount
        '    If RealNick(lstNicks.List(i)) = strNick Then
        '        If lstNicks.Selected(i) Then
        '            Exit Sub
        '        End If
        '    End If
        'Next i
        PutData Me.DataIn, Trim("" & strNick & ": " & Chr(9) & strText)
    End If
End Sub


Public Sub RemoveMode(strMode As String)
    'MsgBox strMode & "~rem"
    Dim i As Integer, j As Integer
    
    For i = 1 To intModes
        'MsgBox Modes(i).mode & "~" & strMode & ".."
        If Modes(i).mode = strMode Then
            Modes(i).mode = ""
            For j = i To intModes - 1
                Modes(j) = Modes(j + 1)
            Next j
            intModes = intModes - 1
            ReDim Preserve Modes(1 To intModes) As typMode
            Update
            Exit Sub
        End If
    Next i
End Sub

Public Sub RemoveNick(strNick As String)
    Dim i As Integer, j As Integer, strTemp As String
    
    For i = 1 To intNicks
        If Nicks(i).Nick = strNick Then
            For j = i To intNicks - 1
                Nicks(j) = Nicks(j + 1)
            Next j
            intNicks = intNicks - 1
            ReDim Preserve Nicks(1 To intNicks) As Nick
            
            For j = 0 To lstNicks.ListCount - 1
                strTemp = lstNicks.List(j)
                strTemp = Replace(strTemp, "@", "")
                strTemp = Replace(strTemp, "+", "")
                strTemp = Replace(strTemp, "%", "")
                If strTemp = strNick Then
                    lstNicks.RemoveItem j
                    Exit Sub
                End If
            Next j
            Exit Sub
        End If
    Next i
End Sub

Public Sub SetHelper(strNick As String, bWhich As Boolean)
    Dim i As Integer, bInd As Integer
    
    For i = 1 To intNicks
        If Nicks(i).Nick = strNick Then
            Nicks(i).helper = bWhich
            bInd = i
            Exit For
        End If
    Next i
    For i = 0 To lstNicks.ListCount - 1
        If RealNick(lstNicks.List(i)) = strNick Then
            lstNicks.RemoveItem i
            lstNicks.AddItem DisplayNick(Nicks(bInd))
            Exit For
        End If
    Next i
End Sub

Public Sub SetOp(strNick As String, bWhich As Boolean)
    Dim i As Integer, bInd As Integer
    
    For i = 1 To intNicks
        If Nicks(i).Nick = strNick Then
            Nicks(i).op = bWhich
            bInd = i
            Exit For
        End If
    Next i
    For i = 0 To lstNicks.ListCount - 1
        If RealNick(lstNicks.List(i)) = strNick Then
            lstNicks.RemoveItem i
            lstNicks.AddItem DisplayNick(Nicks(bInd))

            Exit For
        End If
    Next i
End Sub

Public Sub SetVoice(strNick As String, bWhich As Boolean)
    Dim i As Integer, bInd As Integer
    
    'MsgBox strNick & "~" & bWhich
    For i = 1 To intNicks
        If Nicks(i).Nick = strNick Then
            Nicks(i).voice = bWhich
            bInd = i
            'MsgBox Nicks(i).voice
            Exit For
        End If
    Next i
    For i = 0 To lstNicks.ListCount - 1
        If RealNick(lstNicks.List(i)) = strNick Then
            lstNicks.RemoveItem i
            lstNicks.AddItem DisplayNick(Nicks(bInd))

            Exit For
        End If
    Next i
End Sub

Sub Update()
    strMode = ModeString()
    Dim strExtra As String
    If intLimit <> 0 Then strExtra = strExtra & " " & CStr(intLimit)
    If strKey <> "" Then strExtra = strExtra & " " & strKey
    Me.Caption = strName & " [" & strMode & strExtra & "]"
End Sub

Private Sub DataOut_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then bControl = True   'control
End Sub

Private Sub DataOut_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If Left(DataOut.Text, 1) = "/" Then
            Client.SendData GetAlias(strName, RightR(DataOut.Text, 1))
            If Me.Tag = "PARTNOW" Then
                Me.Tag = "NOPART"
                Unload Me
                Exit Sub
            End If
        Else
            Client.SendData "PRIVMSG " & strName & " :" & DataOut.Text
            PutData DataIn, "" & strMyNick & ":" & Chr(9) & DataOut.Text
        End If
        
        DataOut.Text = ""
    End If
    
    If bControl Then
        'MsgBox KeyAscii
        If KeyAscii = 11 Then
            DataOut.SelText = strColor
        ElseIf KeyAscii = 2 Then
            DataOut.SelText = strBold
        ElseIf KeyAscii = 21 Then
            DataOut.SelText = strUnderline
        ElseIf KeyAscii = 18 Then
            DataOut.SelText = strReverse
        End If
    End If
    
End Sub


Private Sub DataOut_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then bControl = False   'control
End Sub

Private Sub Form_Activate()
    DataOut.SetFocus
End Sub

Private Sub Form_GotFocus()
    DataOut.SetFocus
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 4500 Then Me.Width = 4500
    If Me.Height < 2500 Then Me.Height = 2500
    
    shpTopic.Width = Me.ScaleWidth - 1890
    picTopic.Width = shpTopic.Width - 60
    rtbTopic.Width = picTopic.Width + 150
    shpDI.Width = Me.ScaleWidth - 1800
    shpDI.Height = Me.ScaleHeight - 550
    picFlat.Width = Me.ScaleWidth - 330
    picFlat.Height = Me.ScaleHeight - 600
    DataIn.Width = Me.ScaleWidth - 2020
    DataIn.Height = Me.ScaleHeight - 870
    DataOut.Width = Me.ScaleWidth - 180
    picNicks.Left = Me.ScaleWidth - 1990
    picNicks.Height = DataIn.Height
    lstNicks.Height = DataIn.Height + 80
    picDO.Top = Me.ScaleHeight - 840
    picDO.Width = Me.ScaleWidth - 350
    shpBlue.Height = Me.ScaleHeight + 25
    Toolbar.Left = Me.ScaleWidth - 850
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    
    If Me.Tag <> "NOPART" Then
        Client.SendData "PART " & strName & " :closed channel"
    End If
    Me.Tag = ""
    
    strName = ""
    Me.Caption = ""
    lblName = ""
    strMode = ""
    intModes = 0
    On Error Resume Next
    Unload Channels(Me.Tag)
    strKey = ""
    intLimit = 0
    
    
End Sub


Private Sub rtbTopic_KeyPress(KeyAscii As Integer)
    If rtbTopic.Tag = "locked" Then KeyAscii = 0: Exit Sub
    If KeyAscii = 13 Then
        Client.SendData "TOPIC " & strName & " :" & rtbTopic.Text
        KeyAscii = 0
    End If
    
End Sub


