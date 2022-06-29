VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vb"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmChat.frx":0000
   ScaleHeight     =   4695
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrChat 
      Interval        =   1000
      Left            =   7080
      Top             =   120
   End
   Begin MSComDlg.CommonDialog cdgColor 
      Left            =   6480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBold 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   3550
      Picture         =   "frmChat.frx":8CED6
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   3810
      Width           =   315
   End
   Begin VB.PictureBox picItalic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   3950
      Picture         =   "frmChat.frx":8D318
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   3810
      Width           =   315
   End
   Begin VB.PictureBox picUnderline 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   4330
      Picture         =   "frmChat.frx":8D75A
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   3810
      Width           =   315
   End
   Begin VB.PictureBox picUnderline 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   4330
      Picture         =   "frmChat.frx":8DB9C
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   3810
      Width           =   315
   End
   Begin VB.PictureBox picItalic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   3950
      Picture         =   "frmChat.frx":8DFDE
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   3810
      Width           =   315
   End
   Begin VB.PictureBox picBold 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   3550
      Picture         =   "frmChat.frx":8E420
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   3810
      Width           =   315
   End
   Begin RichTextLib.RichTextBox txtChatSend 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4130
      Width           =   5320
      _ExtentX        =   9393
      _ExtentY        =   661
      _Version        =   393217
      HideSelection   =   0   'False
      MultiLine       =   0   'False
      MaxLength       =   92
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmChat.frx":8E862
   End
   Begin VB.ComboBox cmbFont 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "Tahoma"
      Top             =   3770
      Width           =   3015
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6376
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmChat.frx":8E923
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      Height          =   1230
      Left            =   6285
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblSend 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   5520
      TabIndex        =   12
      Top             =   4140
      Width           =   640
   End
   Begin VB.Label lblFontColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3150
      TabIndex        =   11
      Top             =   3810
      Width           =   340
   End
   Begin VB.Label lblRoomCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6360
      TabIndex        =   2
      Top             =   840
      Width           =   210
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************************************
'* 2.16.99 - dos - email: xdosx@hotmail.com - aim: xdosx  *
'**********************************************************
'* ok, let me apologize now for this example. its nothing *
'* great, but it does answer several questions i have     *
'* received about the rich text control, fonts in a combo *
'* box, and how aol gets its rich to float text. i was    *
'* just goofing off when i wrote this. if it offends you, *
'* it was probably meant to. have a sense of humor and    *
'* don't come whining to me about it.                     *
'* i didn't document this source because i feel that the  *
'* code is pretty self explanitory. if you can't figure it*
'* out, feel free to email me.                            *
'*                                                        *
'* dos                                                    *
'**********************************************************
Private Sub cmbFont_Click()
    txtChatSend.SelFontName = cmbFont.Text
    txtChatSend.SetFocus
End Sub

Private Sub Form_Load()
    Dim intLoadFonts As Integer
    For intLoadFonts% = 0 To Screen.FontCount - 1
        cmbFont.AddItem Screen.Fonts(intLoadFonts%)
    Next intLoadFonts%
    cmbFont.Text = "Arial"
    Call DoChatStuff("OnlineHost", "*** You are in " & Chr(34) & "vb" & Chr(34) & ". ***", False)
    With lstNames
        .AddItem "Iast joopi"
        .AddItem "IVIR agent"
        .AddItem "Syber"
        .AddItem "VVoodie"
        .AddItem "dos"
        .AddItem "PeaceX101"
        .AddItem "Iamer 101"
        .AddItem "Izekial83"
        .AddItem "WanaBeHkr"
        .AddItem "VB Punk"
        .AddItem "numb"
        .AddItem "It Be Mi"
        .AddItem "Iike TNT"
        .AddItem "MaGuSHaVoK"
        .AddItem "SxMizEviL"
        .AddItem "X AoD X"
        .AddItem "stingy po0"
        .AddItem "BuLLeR"
        .AddItem "MacroBoy"
        .AddItem "ahoy cia"
        .AddItem "beav"
    End With
    lblRoomCount.Caption = lstNames.ListCount
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tmrChat.Enabled = False
End Sub

Private Sub lblFontColor_Click()
    On Error GoTo ErrHandler:
    cdgColor.CancelError = True
    cdgColor.ShowColor
    txtChatSend.SelColor = cdgColor.Color
ErrHandler:
    Exit Sub
End Sub

Private Sub lblSend_Click()
    Dim lngSpot As Long, strChat As String
    If txtChatSend.Text <> "" Then
        strChat$ = LCase(txtChatSend.Text)
        Call DoChatStuff("dos", txtChatSend.TextRTF, True)
        txtChatSend.SelFontName = cmbFont.Text
        txtChatSend.Text = ""
    End If
End Sub

Private Sub picBold_Click(Index As Integer)
    If Index = 1 Then
        picBold(0).Visible = True
        picBold(1).Visible = False
        txtChatSend.SelBold = True
    Else
        picBold(0).Visible = False
        picBold(1).Visible = True
        txtChatSend.SelBold = False
    End If
    txtChatSend.SetFocus
End Sub

Private Sub picItalic_Click(Index As Integer)
    If Index = 1 Then
        picItalic(0).Visible = True
        picItalic(1).Visible = False
        txtChatSend.SelItalic = True
    Else
        picItalic(0).Visible = False
        picItalic(1).Visible = True
        txtChatSend.SelItalic = False
    End If
    txtChatSend.SetFocus
End Sub

Private Sub picUnderline_Click(Index As Integer)
    If Index = 1 Then
        picUnderline(0).Visible = True
        picUnderline(1).Visible = False
        txtChatSend.SelUnderline = True
    Else
        picUnderline(0).Visible = False
        picUnderline(1).Visible = True
        txtChatSend.SelUnderline = False
    End If
    txtChatSend.SetFocus
End Sub

Private Sub tmrChat_Timer()
    Call RandomStuff
End Sub

Private Sub txtChatSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        lblSend_Click
    End If
End Sub

Private Sub txtChatSend_SelChange()
    If txtChatSend.SelBold = False Then
        picBold(0).Visible = False
        picBold(1).Visible = True
    Else
        picBold(0).Visible = True
        picBold(1).Visible = False
    End If
    If txtChatSend.SelItalic = False Then
        picItalic(0).Visible = False
        picItalic(1).Visible = True
        txtChatSend.SelItalic = False
    Else
        picItalic(0).Visible = True
        picItalic(1).Visible = False
        txtChatSend.SelItalic = True
    End If
    If txtChatSend.SelUnderline = False Then
        picUnderline(0).Visible = False
        picUnderline(1).Visible = True
        txtChatSend.SelUnderline = False
    Else
        picUnderline(0).Visible = True
        picUnderline(1).Visible = False
        txtChatSend.SelUnderline = True
    End If
End Sub
