VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmIM 
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   Icon            =   "frmIM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   3360
      Width           =   855
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
      TextRTF         =   $"frmIM.frx":1272
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
      TextRTF         =   $"frmIM.frx":12F4
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSend_Click()
  If rtfSend.Text <> "" And frmSignOn.wskAIM.State = sckConnected Then
    Call SendProc(2, "toc_send_im " & Me.Tag & " " & Chr(34) & Normalize("<HTML>" & rtfSend.Text & "</HTML>") & Chr(34) & Chr(0))
    Call RTFUpdate(rtfDisplay, "\par\plain\fs16\cf2\b " & m_strFormattedSN$ & ": \plain\fs16\cf0 " & FixRTF(rtfSend.Text))
    rtfSend.Text = ""
    Call PlayWav(strSoundIMOut$)
  End If
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> 1 Then
    If Me.Width < 4000 Then Me.Width = 4000
    If Me.Height < 4000 Then Me.Height = 4000
    rtfDisplay.Width = Me.Width - 260
    rtfDisplay.Height = Me.Height - 2205
    rtfSend.Width = Me.Width - 260
    rtfSend.Top = Me.Height - 1980
    cmdSend.Left = Me.Width - 1095
    cmdSend.Top = Me.Height - 900
  End If
End Sub

Private Sub rtfSend_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If rtfSend.Text <> "" And frmSignOn.wskAIM.State = sckConnected Then
      Call SendProc(2, "toc_send_im " & Me.Tag & " " & Chr(34) & Normalize("<HTML>" & rtfSend.Text & "</HTML>") & Chr(34) & Chr(0))
      Call RTFUpdate(rtfDisplay, "\par\plain\fs16\cf2\b " & m_strFormattedSN$ & ": \plain\fs16\cf0 " & FixRTF(rtfSend.Text))
      rtfSend.Text = ""
      Call PlayWav(strSoundIMOut$)
    End If
    KeyAscii = 0
  End If
End Sub
