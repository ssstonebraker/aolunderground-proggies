VERSION 5.00
Begin VB.Form frmInvite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Invite"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmInvite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRoom 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtMessage 
      Height          =   735
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtNames 
      Height          =   735
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Room:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1380
   End
   Begin VB.Label Label2 
      Caption         =   "Invite Message:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Names:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1380
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Chat Invite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1320
   End
End
Attribute VB_Name = "frmInvite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdSend_Click()
  Dim strBuddies As String, lngFormIndex As Long
  If Trim(txtNames.Text) = "" Then
    MsgBox "You must enter names to invite.", vbOKOnly + vbInformation, "Error"
  ElseIf Trim(txtMessage.Text) = "" Then
    MsgBox "You must an invitation message.", vbOKOnly + vbInformation, "Error"
  ElseIf Trim(txtRoom.Text) = "" Then
    MsgBox "You must enter a room name.", vbOKOnly + vbInformation, "Error"
  Else
    lngFormIndex& = FormByCaption(LCase(Replace(txtRoom.Text, " ", "")))
    If lngFormIndex& = -1 Then
      strBuddies$ = Replace(txtNames.Text, " ", "")
      strBuddies$ = Replace(strBuddies$, vbCrLf, " ")
      strInviteBuddies$ = strBuddies$
      strInviteMessage$ = txtMessage.Text
      strInviteRoom$ = txtRoom.Text
      Call SendProc(2, "toc_chat_join 4 " & Chr(34) & txtRoom.Text & Chr(34) & Chr(0))
    Else
      strBuddies$ = Replace(txtNames.Text, " ", "")
      strBuddies$ = Replace(strBuddies$, vbCrLf, " ")
      Call SendProc(2, "toc_chat_invite " & Forms(lngFormIndex&).Tag & " " & Chr(34) & Trim(txtMessage.Text) & Chr(34) & " " & Replace(strBuddies$, vbCrLf, " ") & Chr(0))
    End If
    Unload Me
  End If
End Sub
