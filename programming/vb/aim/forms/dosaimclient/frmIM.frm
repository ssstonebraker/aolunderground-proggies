VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmIM 
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   Icon            =   "frmIM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton Command6 
         Caption         =   "&Cancel"
         Height          =   615
         Left            =   2880
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Normal Warning"
         Height          =   615
         Left            =   1560
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Warn &Anonymously"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Would you like to warn this user anonymously? Anonymous warnings are less effective."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Block"
      Height          =   615
      Left            =   960
      Picture         =   "frmIM.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Warn!"
      Height          =   615
      Left            =   120
      Picture         =   "frmIM.frx":180C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Info"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      DownPicture     =   "frmIM.frx":1EC6
      Height          =   615
      Left            =   3600
      Picture         =   "frmIM.frx":2F7C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
   Begin RichTextLib.RichTextBox rtfSend 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1720
      _Version        =   393217
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmIM.frx":4032
   End
   Begin RichTextLib.RichTextBox rtfDisplay 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3625
      _Version        =   393217
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmIM.frx":40E0
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSend_Click()
  If rtfSend.Text <> "" And frmSignOn.wskAIM.state = sckConnected Then
    Call SendProc(2, "toc_send_im " & Me.Tag & " " & Chr(34) & Normalize("<HTML>" & rtfSend.Text & "</HTML>") & Chr(34) & Chr(0))
    Call RTFUpdate(rtfDisplay, "\par\plain\fs16\cf2\b " & m_strFormattedSN$ & ": \plain\fs16\cf0 " & FixRTF(rtfSend.Text))
    rtfSend.Text = ""
    Call PlayWav(strSoundIMOut$)
  End If
End Sub

Private Sub Command1_Click()

        frmInfo.Show
        frmInfo.WhoInfo.Text = Me.Caption
        frmInfo.Caption = "Buddy Info: " & frmInfo.WhoInfo.Text
        Call SendProc(2, "toc_get_info " & Chr(34) & frmInfo.WhoInfo.Text & Chr(34) & Chr(0))

End Sub

Private Sub Command2_Click()
Command3.Enabled = False
Command1.Enabled = False
Command2.Enabled = False

cmdSend.Enabled = False
Frame1.Visible = True

End Sub

Private Sub Command4_Click()
Call SendProc(2, "toc_evil " & Me.Tag & " " & "anon" & Chr(0))
Frame1.Visible = False
Command3.Enabled = True
Command1.Enabled = True
cmdSend.Enabled = True
Command2.Enabled = True

End Sub

Private Sub Command5_Click()
Call SendProc(2, "toc_evil " & Me.Tag & " " & "norm" & Chr(0))
Frame1.Visible = False
Command3.Enabled = True
Command1.Enabled = True
cmdSend.Enabled = True
Command2.Enabled = True
End Sub

Private Sub Command6_Click()
Frame1.Visible = False
Command3.Enabled = True
Command1.Enabled = True
cmdSend.Enabled = True
Command2.Enabled = True
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> 1 Then
    If Me.Width < 4530 Then Me.Width = 4530
    If Me.Height < 4530 Then Me.Height = 4530 '4000
    rtfDisplay.Width = Me.Width - 260
    rtfDisplay.Height = Me.Height - 2205
    rtfSend.Width = Me.Width - 260
    rtfSend.Top = Me.Height - 2100 '1980
    cmdSend.Left = Me.Width - 1095
    cmdSend.Top = Me.Height - 1100 '900
    cmdSend.Left = Me.Width - (cmdSend.Width + 100)
    Command1.Top = cmdSend.Top
    Frame1.Height = rtfSend.Top + rtfSend.Height
    Frame1.Width = rtfSend.Width
    
    Command2.Top = cmdSend.Top
    Command3.Top = cmdSend.Top
    'Command 1 is Get Info feature, added by Steve.
    'Command 2 is Warn feature, added by Tom.
    'Command 3 is Block feature, added by Tom.
    
    
    Command1.Left = cmdSend.Left - 840
    'cmdSend
  End If
End Sub

Private Sub rtfSend_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If rtfSend.Text <> "" And frmSignOn.wskAIM.state = sckConnected Then
      Call SendProc(2, "toc_send_im " & Me.Tag & " " & Chr(34) & Normalize("<HTML>" & rtfSend.Text & "</HTML>") & Chr(34) & Chr(0))
      Call RTFUpdate(rtfDisplay, "\par\plain\fs16\cf2\b " & m_strFormattedSN$ & ": \plain\fs16\cf0 " & FixRTF(rtfSend.Text))
      rtfSend.Text = ""
      Call PlayWav(strSoundIMOut$)
    End If
    KeyAscii = 0
  End If
End Sub
