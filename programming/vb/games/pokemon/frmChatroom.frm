VERSION 5.00
Begin VB.Form frmChatroom 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Pokémon Adventure"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChatroom.frx":0000
   ScaleHeight     =   4620
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   210
      Left            =   555
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2700
      Width           =   5730
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   2085
      Left            =   195
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   465
      Width           =   6525
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   5805
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lblMinimize 
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
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   5580
      TabIndex        =   3
      Top             =   120
      Width           =   210
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Pokémon Adventure Chat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   315
      TabIndex        =   2
      Top             =   120
      Width           =   2205
   End
End
Attribute VB_Name = "frmChatroom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub
Private Sub lblMinimize_Click()
    Me.Hide
End Sub
Private Sub lblSend_Click()
    If Not Len(txtSend.Text) = 0 Then
        txtSend.Text = ReplaceString(txtSend.Text, vbNewLine, "")
        If frmBattle.TurnA = 0 Then
            frmInternetConnect.sckConnect.SendData "CHT-" & txtSend.Text
        End If
        If frmBattle.TurnA = 1 Then
            frmInternetListen.sckListen.SendData "CHT-" & txtSend.Text
        End If
        txtChat.Text = txtChat.Text & vbNewLine & frmMain.Player & ":" & Chr(9) & txtSend.Text
        txtSend.Text = Empty
        txtChat.SelStart = Len(txtChat.Text)
    End If
End Sub
Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtSend.Text = ReplaceString(txtSend.Text, vbNewLine, "")
        If Not Len(txtSend.Text) = 0 Then
            If frmBattle.TurnA = 0 Then
                frmInternetConnect.sckConnect.SendData "CHT-" & txtSend.Text
            End If
            If frmBattle.TurnA = 1 Then
                frmInternetListen.sckListen.SendData "CHT-" & txtSend.Text
            End If
            txtChat.Text = txtChat.Text & vbNewLine & frmMain.Player & ":" & Chr(9) & txtSend.Text
            txtSend.Text = Empty
            txtChat.SelStart = Len(txtChat.Text)
        End If
    End If
End Sub
