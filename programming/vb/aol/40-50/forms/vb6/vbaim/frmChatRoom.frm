VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmChatRoom 
   Caption         =   "Chat Room"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   Icon            =   "frmChatRoom.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstNames 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      IntegralHeight  =   0   'False
      Left            =   5880
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox rtfSend 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmChatRoom.frx":1272
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfDisplay 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmChatRoom.frx":12E9
   End
   Begin VB.Label lblPeople 
      AutoSize        =   -1  'True
      Caption         =   "0 People"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5880
      TabIndex        =   3
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmChatRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
  If Me.WindowState <> 1 Then
    If Me.Width < 6000 Then Me.Width = 6000
    If Me.Height < 2000 Then Me.Height = 4000
    rtfDisplay.Width = Me.Width - 2160
    rtfDisplay.Height = Me.Height - 1500
    lblPeople.Left = Me.Width - 1935
    lstNames.Left = Me.Width - 1935
    lstNames.Height = Me.Height - 1740
    rtfSend.Width = Me.Width - 360
    rtfSend.Top = Me.Height - 1275
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call SendProc(2, "toc_chat_leave " & Chr(34) & Me.Tag & Chr(34) & Chr(0))
End Sub

Private Sub lstNames_DblClick()
  Dim lngFormindex As Long
  If lstNames.ListIndex > -1 Then
    lngFormindex& = FormByTag(LCase(Replace(lstNames.Text, " ", "")))
    If lngFormindex& > -1 Then
      Forms(lngFormindex&).SetFocus
    Else
      Dim frmNewIM As New frmIM
      With frmNewIM
        .Caption = lstNames.Text
        .Tag = LCase(Replace(lstNames.Text, " ", ""))
        .Show
      End With
    End If
  End If
End Sub

Private Sub rtfSend_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If rtfSend.Text <> "" And frmSignOn.wskAIM.State = sckConnected Then
      Call SendProc(2, "toc_chat_send " & Me.Tag & " " & Chr(34) & Normalize("<HTML>" & rtfSend.Text & "</HTML>") & Chr(34) & Chr(0))
      rtfSend.Text = ""
    End If
    KeyAscii = 0
  End If
End Sub
