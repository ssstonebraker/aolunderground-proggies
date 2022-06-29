VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCableChat 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4305
   ClientLeft      =   4950
   ClientTop       =   3675
   ClientWidth     =   6435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClear 
      Interval        =   1
      Left            =   5610
      Top             =   4380
   End
   Begin VB.PictureBox Pic3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4770
      Picture         =   "frmCableChat.frx":0000
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   10
      Top             =   4755
      Width           =   270
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   4485
      Picture         =   "frmCableChat.frx":0BAD
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   9
      Top             =   4740
      Width           =   270
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   5940
      Picture         =   "frmCableChat.frx":1912
      ScaleHeight     =   225
      ScaleWidth      =   450
      TabIndex        =   8
      Top             =   4020
      Width           =   450
   End
   Begin VB.PictureBox b3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1935
      Picture         =   "frmCableChat.frx":24BF
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   4680
      Width           =   225
   End
   Begin VB.PictureBox b2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2205
      Picture         =   "frmCableChat.frx":3155
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   4680
      Width           =   225
   End
   Begin VB.PictureBox b1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6210
      Picture         =   "frmCableChat.frx":3E1D
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   0
      Width           =   225
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   3660
      Left            =   45
      TabIndex        =   3
      Top             =   300
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   6456
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MaxLength       =   9999999
      TextRTF         =   $"frmCableChat.frx":4AB3
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   60
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4005
      Width           =   5850
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Left            =   0
      Picture         =   "frmCableChat.frx":4B7C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblDrag 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   420
      TabIndex        =   7
      Top             =   315
      Width           =   5955
   End
   Begin VB.Image c3 
      Height          =   90
      Left            =   2490
      Picture         =   "frmCableChat.frx":5446
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1065
   End
   Begin VB.Image c2 
      Height          =   90
      Left            =   2490
      Picture         =   "frmCableChat.frx":5DF8
      Stretch         =   -1  'True
      Top             =   4815
      Width           =   1065
   End
   Begin VB.Image c4 
      Height          =   90
      Left            =   2490
      Picture         =   "frmCableChat.frx":6734
      Stretch         =   -1  'True
      Top             =   4740
      Width           =   1065
   End
   Begin VB.Label lblChatRoom 
      BackStyle       =   0  'Transparent
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   2265
      TabIndex        =   2
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Pokémon Center Chat -"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   270
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.Image c1 
      Height          =   225
      Left            =   -15
      Picture         =   "frmCableChat.frx":7039
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6555
   End
End
Attribute VB_Name = "frmCableChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bOn As Boolean, b1On As Boolean
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload frmCableClub
    frmMain.Show
End Sub
Private Sub Form_GotFocus()
    c1.Picture = c3.Picture
End Sub
Private Sub Form_LostFocus()
    c1.Picture = c4.Picture
End Sub
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub lblChatRoom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    c1.Picture = c2.Picture
    SendMessage Me.hwnd, WM_SYSCOMMAND, WM_MOVE, 0
    c1.Picture = c3.Picture
End Sub
Private Sub lblMinimize_Click()
    Unload Me
End Sub
Private Sub lblSend_Click()
    txtSend.Text = ReplaceString(txtSend.Text, vbNewLine, "")
    If Not Len(txtSend.Text) = 0 Then
        If Not Len(TrimSpaces(txtSend.Text)) = 0 Then
            If frmCableClub.sck.State <> sckConnected Then
                txtChat.SelStart = Len(frmCableChat.txtChat.Text)
                txtChat.SelFontName = "Tahoma"
                txtChat.SelFontSize = 8
                txtChat.SelBold = True
                txtChat.SelUnderline = False
                txtChat.SelItalic = False
                txtChat.SelColor = vbGreen + 1000
                txtChat.SelText = vbNewLine + "Previous message was not sent : Connection to Server Lost"
                txtChat.SelStart = Len(txtChat.Text)
                txtSend.Text = Empty
            Else
                frmCableClub.sck.SendData "cht-" + frmMain.Player + "_" + ReplaceString(Left(txtSend.Text, 100), vbNewLine, Empty) + "_" + TrimSpaces(LCase(lblChatRoom.Caption))
                If Left(txtSend.Text, 1) = "/" Then
                    txtSend.Text = Empty
                    TimeOut 0.5
                    Exit Sub
                End If
                txtChat.SelStart = Len(frmCableChat.txtChat.Text)
                txtChat.SelFontName = "Tahoma"
                txtChat.SelFontSize = 8
                txtChat.SelBold = True
                txtChat.SelUnderline = False
                txtChat.SelItalic = False
                txtChat.SelColor = vbBlue
                txtChat.SelText = vbNewLine + "<" + frmMain.Player + "> "
                txtChat.SelStart = Len(txtChat.Text)
                txtChat.SelFontName = "Tahoma"
                txtChat.SelFontSize = 8
                txtChat.SelBold = False
                txtChat.SelUnderline = False
                txtChat.SelItalic = False
                txtChat.SelColor = vbBlack
                txtChat.SelText = txtSend.Text
                txtChat.SelStart = Len(txtChat.Text)
                txtSend.Text = Empty
                TimeOut 0.5
            End If
        Else
            txtChat.SelStart = Len(frmCableChat.txtChat.Text)
            txtChat.SelFontName = "Tahoma"
            txtChat.SelFontSize = 8
            txtChat.SelBold = True
            txtChat.SelUnderline = False
            txtChat.SelItalic = False
            txtChat.SelColor = vbGreen + 1000
            txtChat.SelText = vbNewLine + "Previous message was not sent : Send text is blank."
            txtChat.SelStart = Len(txtChat.Text)
            txtSend.Text = Empty
        End If
    Else
        txtChat.SelStart = Len(frmCableChat.txtChat.Text)
        txtChat.SelFontName = "Tahoma"
        txtChat.SelFontSize = 8
        txtChat.SelBold = True
        txtChat.SelUnderline = False
        txtChat.SelItalic = False
        txtChat.SelColor = vbGreen + 1000
        txtChat.SelText = vbNewLine + "Previous message was not sent : Send text is blank."
        txtChat.SelStart = Len(txtChat.Text)
        txtSend.Text = Empty
    End If
End Sub
Private Sub lblDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    c1.Picture = c2.Picture
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, WM_MOVE, 0
    c1.Picture = c3.Picture
End Sub
Private Sub tmrClear_Timer()
    If LineCount(txtChat.Text) >= 500 Then
        FormLines
    End If
End Sub
Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        lblSend_Click
    End Select
End Sub
Private Sub b1_Click()
    b1On = False
    b1.Picture = b3.Picture
    FormNotOnTop Me
    If MsgBox("Are you sure you want to exit?" + vbNewLine + "(connection to server will be terminated)", vbYesNo) = vbYes Then
        Unload Me
    Else
        FormOnTop Me
    End If
End Sub
Private Sub b1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not b1On Then
        b1On = True
        b1.Picture = b2.Picture
        SetCapture b1.hwnd
    ElseIf X < 0 Or Y < 0 Or X > b1.Width Or Y > b1.Height Then
        b1On = False
        b1.Picture = b3.Picture
        ReleaseCapture
    End If
End Sub
Private Sub Pic1_Click()
    bOn = False
    lblSend_Click
End Sub
Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bOn Then
        bOn = True
        Pic1.Picture = Pic2.Picture
        SetCapture Pic1.hwnd
    ElseIf X < 0 Or Y < 0 Or X > Pic1.Width Or Y > Pic1.Height Then
        bOn = False
        Pic1.Picture = Pic3.Picture
        ReleaseCapture
    End If
End Sub
Function FormLines()
    num& = 100
    strBuffer$ = Empty
str:
    If num& = 501 Then
        txtChat.Text = strBuffer$
        txtChat.SelStart = Len(txtChat.Text)
        Exit Function
    Else
        strBuffer$ = strBuffer$ + vbNewLine + LineFromString(txtChat.Text, num&)
        num& = num& + 1
        GoTo str
    End If
End Function
