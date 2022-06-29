VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmCableClub 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Pokémon Adventure"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStatus 
      Interval        =   1
      Left            =   4575
      Top             =   1950
   End
   Begin MSWinsockLib.Winsock sck 
      Left            =   4575
      Top             =   1965
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ComboBox lstServer 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCableClub.frx":0000
      Left            =   2490
      List            =   "frmCableClub.frx":000D
      TabIndex        =   1
      Text            =   "Select a Server"
      Top             =   1095
      Width           =   2460
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Left            =   2490
      TabIndex        =   5
      Top             =   180
      Width           =   2505
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Height          =   225
      Left            =   3735
      TabIndex        =   4
      Top             =   1485
      Width           =   720
   End
   Begin VB.Label lblConnect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Connect"
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
      Height          =   225
      Left            =   2775
      TabIndex        =   3
      Top             =   1485
      Width           =   720
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Cable Club Server:"
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
      Height          =   225
      Left            =   2490
      TabIndex        =   2
      Top             =   795
      Width           =   1605
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   2460
      Left            =   0
      Top             =   0
      Width           =   5070
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Height          =   225
      Left            =   195
      TabIndex        =   0
      Top             =   2070
      Width           =   135
   End
   Begin VB.Shape shpBgBorder 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FFFFFF&
      Height          =   2220
      Left            =   120
      Top             =   120
      Width           =   2220
   End
   Begin VB.Image imgBg 
      Height          =   2160
      Left            =   150
      Picture         =   "frmCableClub.frx":0052
      Top             =   150
      Width           =   2160
   End
End
Attribute VB_Name = "frmCableClub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub
Private Sub imgBg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub lblCancel_Click()
    If sck.State > 0 Then
        sck.Close
    Else
        Unload Me
    End If
End Sub
Private Sub lblConnect_Click()
    If lstServer.Text = Empty Then
        MsgBoxA Me, "Please select a server first!"
    Else
        If sck.State > 0 Then
            MsgBoxA Me, "Please cancel the current operation before attempting to connect!"
        Else
            sck.Connect lstServer.Text, 7000
        End If
    End If
End Sub
Private Sub lblExit_Click()
    Unload Me
End Sub
Private Sub sck_Connect()
    sck.SendData "usr-" & frmMain.Player
    Me.Hide
    frmCableChat.Show
End Sub
Private Sub sck_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    sck.GetData strData
    If Left(strData, 4) = "cht-" Then
        strBuffer = Right(strData, Len(strData) - 4)
        strUser = Left(strBuffer, InStr(strBuffer, "_") - 1)
        strMsg = Right(strBuffer, Len(strBuffer) - InStr(strBuffer, "_"))
        If Not LCase(strUser) = LCase(frmMain.Player) Then
            frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
            frmCableChat.txtChat.SelFontName = "Tahoma"
            frmCableChat.txtChat.SelFontSize = 8
            frmCableChat.txtChat.SelBold = True
            frmCableChat.txtChat.SelUnderline = False
            frmCableChat.txtChat.SelItalic = False
            frmCableChat.txtChat.SelColor = vbBlue
            frmCableChat.txtChat.SelText = vbNewLine + "<" + strUser + "> "
            frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
            frmCableChat.txtChat.SelFontName = "Tahoma"
            frmCableChat.txtChat.SelFontSize = 8
            frmCableChat.txtChat.SelBold = False
            frmCableChat.txtChat.SelUnderline = False
            frmCableChat.txtChat.SelItalic = False
            frmCableChat.txtChat.SelColor = vbBlack
            frmCableChat.txtChat.SelText = strMsg
            frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
        End If
    End If
    If Left(strData, 4) = "jrm-" Then
        strRoom = Right(strData, Len(strData) - 4)
        frmCableChat.lblChatRoom = strRoom
        frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
        frmCableChat.txtChat.SelFontName = "Tahoma"
        frmCableChat.txtChat.SelFontSize = 8
        frmCableChat.txtChat.SelBold = True
        frmCableChat.txtChat.SelUnderline = False
        frmCableChat.txtChat.SelItalic = False
        frmCableChat.txtChat.SelColor = &HC000&
        frmCableChat.txtChat.SelText = vbNewLine & "*** Now talking in #" + strRoom
        frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
    End If
    If Left(strData, 4) = "sys-" Then
        strUser = Right(strData, Len(strData) - 4)
        frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
        frmCableChat.txtChat.SelFontName = "Tahoma"
        frmCableChat.txtChat.SelFontSize = 8
        frmCableChat.txtChat.SelBold = True
        frmCableChat.txtChat.SelUnderline = False
        frmCableChat.txtChat.SelItalic = False
        frmCableChat.txtChat.SelColor = &H800000
        frmCableChat.txtChat.SelText = vbNewLine & "*** " + strUser
        frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
    End If
    If Left(strData, 5) = "sysg-" Then
        strUser = Right(strData, Len(strData) - 5)
        frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
        frmCableChat.txtChat.SelFontName = "Tahoma"
        frmCableChat.txtChat.SelFontSize = 8
        frmCableChat.txtChat.SelBold = True
        frmCableChat.txtChat.SelUnderline = False
        frmCableChat.txtChat.SelItalic = False
        frmCableChat.txtChat.SelColor = &HC000&
        frmCableChat.txtChat.SelText = vbNewLine & "*** " + strUser
        frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
    End If
    If Left(strData, 3) = "me-" Then
        strBuffer = Right(strData, Len(strData) - 3)
        strB = Left(strBuffer, InStr(strBuffer, "_") - 1)
        strC = Right(strBuffer, Len(strBuffer) - InStr(strBuffer, "_"))
        frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
        frmCableChat.txtChat.SelFontName = "Tahoma"
        frmCableChat.txtChat.SelFontSize = 8
        frmCableChat.txtChat.SelBold = True
        frmCableChat.txtChat.SelUnderline = False
        frmCableChat.txtChat.SelItalic = False
        frmCableChat.txtChat.SelColor = &H800080
        frmCableChat.txtChat.SelText = vbNewLine & "* " + strB + " " + strC
        frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
    End If
    If Left(strData, 5) = "sysr-" Then
        strUser = Right(strData, Len(strData) - 5)
        frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
        frmCableChat.txtChat.SelFontName = "Tahoma"
        frmCableChat.txtChat.SelFontSize = 8
        frmCableChat.txtChat.SelBold = True
        frmCableChat.txtChat.SelUnderline = False
        frmCableChat.txtChat.SelItalic = False
        frmCableChat.txtChat.SelColor = &HC0&
        frmCableChat.txtChat.SelText = vbNewLine & "*** " + strUser
        frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
    End If
    If Left(strData, 5) = "list_" Then
        strBuffer = Right(strData, Len(strData) - 5)
        frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
        frmCableChat.txtChat.SelFontName = "Tahoma"
        frmCableChat.txtChat.SelFontSize = 8
        frmCableChat.txtChat.SelBold = False
        frmCableChat.txtChat.SelUnderline = False
        frmCableChat.txtChat.SelItalic = False
        frmCableChat.txtChat.SelColor = vbViolet + 1000
        frmCableChat.txtChat.SelText = strBuffer
        frmCableChat.txtChat.SelStart = Len(frmCableChat.txtChat.Text)
    End If
End Sub
Private Sub tmrStatus_Timer()
    If sck.State > 0 Then
        lblCancel.ForeColor = &HFF&
    Else
        lblCancel.ForeColor = &HFFFFFF
    End If
    If sck.State = 0 Then
        lblStatus.Caption = "Not Connected."
    End If
    If sck.State = 1 Then
        lblStatus.Caption = "Socket Open."
    End If
    If sck.State = 3 Then
        lblStatus.Caption = "Connection Pending..."
    End If
    If sck.State = 4 Then
        lblStatus.Caption = "Resolving Host..."
    End If
    If sck.State = 5 Then
        lblStatus.Caption = "Host Resolved."
    End If
    If sck.State = 6 Then
        lblStatus.Caption = "Connecting..."
    End If
    If sck.State = 7 Then
        lblStatus.Caption = "Connected."
    End If
    If sck.State = 8 Then
        lblStatus.Caption = "Connection Terminated by Server."
    End If
    If sck.State = 9 Then
        lblStatus.Caption = "Error!"
    End If
End Sub
