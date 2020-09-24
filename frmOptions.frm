VERSION 5.00
Begin VB.Form Options 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3855
   ClientLeft      =   3090
   ClientTop       =   1485
   ClientWidth     =   5625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   975
      TabIndex        =   23
      Top             =   3435
      Width           =   885
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   360
      Left            =   90
      TabIndex        =   22
      Top             =   3435
      Width           =   885
   End
   Begin VB.PictureBox picConnecting 
      BorderStyle     =   0  'None
      Height          =   3465
      Left            =   1980
      ScaleHeight     =   3465
      ScaleWidth      =   3630
      TabIndex        =   17
      Top             =   390
      Width           =   3630
      Begin VB.TextBox txtPort 
         Height          =   315
         Left            =   2955
         MaxLength       =   5
         TabIndex        =   27
         Text            =   "6667"
         Top             =   1830
         Width           =   600
      End
      Begin VB.TextBox txtRetry 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1650
         MaxLength       =   2
         TabIndex        =   25
         Text            =   "99"
         Top             =   3120
         Width           =   225
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   " &Retry Connect        times "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   3120
         Width           =   3240
      End
      Begin VB.TextBox txtIdent 
         Height          =   315
         Left            =   1530
         TabIndex        =   4
         Text            =   "~IDENT"
         Top             =   1185
         Width           =   2025
      End
      Begin VB.CheckBox chkInvisible 
         Appearance      =   0  'Flat
         Caption         =   " Invisible Mode "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   2835
         Width           =   3240
      End
      Begin VB.CheckBox chklReconnect 
         Appearance      =   0  'Flat
         Caption         =   " Reconnect to server on disconnect "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   2550
         Width           =   3240
      End
      Begin VB.CheckBox chkStartUp 
         Appearance      =   0  'Flat
         Caption         =   " Connect to server on Client load "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Top             =   2265
         Width           =   3240
      End
      Begin VB.ComboBox cbServer 
         Height          =   315
         Left            =   210
         TabIndex        =   5
         Text            =   "irc.otherside.com"
         Top             =   1830
         Width           =   2685
      End
      Begin VB.TextBox txtFullName 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Text            =   "jIRC User"
         Top             =   810
         Width           =   2025
      End
      Begin VB.TextBox txtOtherNick 
         Height          =   315
         Left            =   1530
         TabIndex        =   2
         Text            =   "OtherNick"
         Top             =   435
         Width           =   2025
      End
      Begin VB.TextBox txtNick 
         Height          =   315
         Left            =   1530
         TabIndex        =   1
         Text            =   "YourNick"
         Top             =   60
         Width           =   2025
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2790
         TabIndex        =   26
         Top             =   1545
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IDENT:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   24
         Top             =   1215
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   21
         Top             =   1545
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   20
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alternate Nick:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   19
         Top             =   465
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nick:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   18
         Top             =   90
         Width           =   390
      End
   End
   Begin VB.CheckBox chkOP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   " DCC... "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   4
      Left            =   135
      TabIndex        =   16
      Top             =   1785
      Width           =   1680
   End
   Begin VB.CheckBox chkOP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   " Display... "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   3
      Left            =   135
      TabIndex        =   15
      Top             =   1470
      Width           =   1680
   End
   Begin VB.CheckBox chkOP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   " Sounds... "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      Left            =   135
      TabIndex        =   14
      Top             =   1155
      Width           =   1680
   End
   Begin VB.CheckBox chkOP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   " General... "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   135
      TabIndex        =   13
      Top             =   840
      Width           =   1680
   End
   Begin VB.CheckBox chkOP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   " Connecting... "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   135
      TabIndex        =   11
      Top             =   525
      Width           =   1680
   End
   Begin VB.Label lblConnecting 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   195
      TabIndex        =   12
      Top             =   555
      Width           =   1590
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting..."
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   1950
      TabIndex        =   0
      Top             =   75
      Width           =   3285
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Options..."
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
      TabIndex        =   10
      Top             =   75
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   360
      Left            =   -210
      Top             =   0
      Width           =   5955
   End
   Begin VB.Shape shpBlue 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3930
      Left            =   0
      Top             =   0
      Width           =   1950
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bC As Boolean

Private Sub cbServer_GotFocus()
    With cbServer
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub cbServer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPort.SetFocus
End Sub


Private Sub chkOP_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 0 To 4
        'If i <> Index And chkOP(i).Value = 1 Then
        If i <> Index And chkOP(i).Value = 1 Then chkOP(i).Value = 0
    Next i
    chkOP(Index).Value = 1
End Sub


Private Sub chkOP_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 0 To 4
        'If i <> Index And chkOP(i).Value = 1 Then
        If i <> Index And chkOP(i).Value = 1 Then chkOP(i).Value = 0
    Next i
    chkOP(Index).Value = 1
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    strMyNick = txtNick
    strOtherNick = txtOtherNick
    strFullName = txtFullName
    strIdent = txtIdent
    strServer = cbServer.Text
    strPort = CLng(txtPort)
    Unload Me
End Sub

Private Sub Form_Load()
    txtNick = strMyNick
    txtOtherNick = strOtherNick
    txtFullName = strFullName
    txtIdent = strMyIdent
    cbServer.Text = strServer
End Sub

Private Sub txtFullName_GotFocus()
    With txtFullName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtFullName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtIdent.SetFocus
End Sub


Private Sub txtIdent_GotFocus()
    With txtIdent
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtIdent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cbServer.SetFocus
End Sub


Private Sub txtNick_GotFocus()
    With txtNick
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtNick_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOtherNick.SetFocus
End Sub


Private Sub txtOtherNick_GotFocus()
    With txtOtherNick
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtOtherNick_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFullName.SetFocus
End Sub


Private Sub txtPort_GotFocus()
    With txtPort
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOk_Click
End Sub


Private Sub txtRetry_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) Then Else KeyAscii = 0
End Sub


