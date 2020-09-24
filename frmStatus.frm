VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Status 
   Caption         =   "Status"
   ClientHeight    =   3705
   ClientLeft      =   1320
   ClientTop       =   2040
   ClientWidth     =   6600
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
   ScaleHeight     =   3705
   ScaleWidth      =   6600
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3630
      Top             =   225
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
            Picture         =   "frmStatus.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatus.frx":0454
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   330
      Left            =   5790
      TabIndex        =   6
      Top             =   0
      Width           =   720
      _ExtentX        =   1270
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
   Begin VB.PictureBox picFlat 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   150
      ScaleHeight     =   3195
      ScaleWidth      =   6330
      TabIndex        =   2
      Top             =   375
      Width           =   6330
      Begin VB.PictureBox picDO 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   6360
         TabIndex        =   4
         Top             =   2955
         Width           =   6360
         Begin RichTextLib.RichTextBox DataOut 
            Height          =   390
            Left            =   -45
            TabIndex        =   5
            Top             =   -45
            Width           =   6510
            _ExtentX        =   11483
            _ExtentY        =   688
            _Version        =   393217
            Enabled         =   -1  'True
            MultiLine       =   0   'False
            TextRTF         =   $"frmStatus.frx":08A8
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
         TabIndex        =   3
         Top             =   0
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   5159
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmStatus.frx":09A2
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
   Begin VB.Shape shpDI 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   3240
      Left            =   1635
      Top             =   360
      Width           =   4875
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "not connected"
      Height          =   195
      Left            =   1740
      TabIndex        =   1
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Server Status"
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
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   1425
   End
   Begin VB.Shape shpBlue 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3930
      Left            =   -15
      Top             =   -15
      Width           =   1650
   End
End
Attribute VB_Name = "Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub


Private Sub DataIn_Change()
    DataIn.SelStart = Len(DataIn.Text)
End Sub

Private Sub DataOut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Left(DataOut.Text, 1) = "/" Then
            Client.SendData GetAlias("", RightR(DataOut.Text, 1))
        Else
            Client.SendData GetAlias("", RightR(DataOut.Text, 1))
        End If
        DataOut.Text = ""
    End If
End Sub


Private Sub Form_Activate()
DataOut.SetFocus
End Sub

Private Sub Form_GotFocus()
    If Me.Visible Then DataOut.SetFocus
End Sub

Private Sub Form_Load()
    strBold = Chr(BOLD)
    strUnderline = Chr(UNDERLINE)
    strColor = Chr(Color)
    strReverse = Chr(REVERSE)
    strAction = Chr(ACTION)
    
    PutData Status.DataIn, "Welcome to " & strBold & strColor & "12" & "projectIRC" & strColor & "!"
    PutData Status.DataIn, "projectIRC version " & strColor & "4" & "1" & strColor & " build " & strColor & "4" & App.Revision & strColor
End Sub

Private Sub Form_Resize()
    If Status.WindowState = vbMinimized Then Exit Sub
    If Status.Width < 4500 Then Me.Width = 4500
    If Status.Height < 2500 Then Status.Height = 2500
    
    shpDI.Width = Status.ScaleWidth - 1800
    shpDI.Height = Status.ScaleHeight - 550
    picFlat.Width = Status.ScaleWidth - 330
    picFlat.Height = Status.ScaleHeight - 600
    DataIn.Width = Status.ScaleWidth - 350
    DataIn.Height = Status.ScaleHeight - 870
    DataOut.Width = Status.ScaleWidth - 180
    picDO.Top = Status.ScaleHeight - 840
    picDO.Width = Status.ScaleWidth - 350
    shpBlue.Height = Status.ScaleHeight + 25
    Toolbar.Left = Status.ScaleWidth - 850
End Sub


Private Sub sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            '* Connect
            Client.mnu_File_Connect_Click
        Case 2
            '* Disconnect
            Call Client.mnu_File_Disconnect_Click
    End Select
End Sub


