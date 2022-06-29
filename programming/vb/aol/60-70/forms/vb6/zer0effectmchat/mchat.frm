VERSION 5.00
Object = "{AAF4C707-6DE7-4418-81C8-24C7D8776A57}#6.0#0"; "Chat.ocx"
Begin VB.Form mchat 
   Caption         =   "Shaggy's AOL 7 Mchat/Ccom Example"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   Icon            =   "mchat.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5160
      Top             =   1680
   End
   Begin Zer0effect.Chat Chat1 
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   5880
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Room Name: None"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "People: 0"
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "mchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' all of this is pretty easy all you
' really need to know is trig=trigger
' so if you want to change it you'll need
' to change the trig string
' trig=text1 or trig=strargument1
' and strargument1=anything after
' you command like pr and before a comma
' where strargument2=everything after the
' comma however this is not ready for a
' starargument 3 which is only really used
' for sendmail which there are no subs for
' thanks for d/ling this and if you need
' help just e-mail me at shaggyze@aol.com

Dim Trig As String
Private Sub Chat1_Scan(ByVal sn As String, ByVal Msg As String)
Dim add As String
add = 20 - Len(sn)
Text1 = Text1.Text & sn & ":" & AddString(add, " ") & Msg & vbCrLf
Text1.SelLength = Len(Text1)
Dim lngSpace As Long, strCommand As String, strArgument1 As String
   Dim strArgument2 As String, lngComma As Long
   If GetUser = sn Then
   If InStr(Msg, Trig) = 1& Then ' tells text wil start with period
      lngSpace& = InStr(Msg, " ") 'gets what said
      If lngSpace& = 0& Then
         strCommand$ = Msg$
      Else
         strCommand$ = Left(Msg$, lngSpace& - 1&)
      End If
strArgument1$ = Trim(Mid(Msg$, lngSpace& + 1&, Len(Msg)))
      Select Case LCase(strCommand)

  Case Is = Trig & "im"
       lngComma& = InStr(Msg$, ",")
  strArgument1$ = Trim(Mid(Msg$, lngSpace& + 1&, lngComma& - lngSpace& - 1&))
  strArgument2$ = Trim(Right(Msg$, Len(Msg$) - lngComma& - 1&))
  ChatSend ("<font color=blue>(¦oº·-»imed-<u>[" & strArgument1 & "]</u>«-·ºo¦)")
  Call SendInstantMessage(strArgument1, strArgument2)

Case Is = Trig & "pr"
ChatSend ("<font color=blue>(¦oº·-»entering pr-<u>[" & strArgument1 & "]</u>«-·ºo¦)")
Call EnterPR(strArgument1)
Pause 2
Call AddRoomToList(List1, True)
Call KillDupes(List1)
End Select
End If
End If

End Sub


Private Sub Form_Load()
List1.Clear
Call EnterPR("Zer0Effect")
Text1.Width = Me.Width - (List1.Width + 100)
Text2.Width = Me.Width - (List1.Width + 100)
List1.Left = Me.Width - List1.Width - 100
Label1.Left = Me.Width - Label1.Width - 100
Text2.Top = Me.Height - Text2.Height - 500
Text1.Height = Me.Height - Text2.Height - 700
List1.Height = Me.Height - 700
Call AddRoomToList(List1, True)
Call KillDupes(List1)
Label1 = "People: " & List1.ListCount
Label2 = "Room Name: " & GetCaption(FindChat)
End Sub

Private Sub Form_Resize()
Text1.Width = Me.Width - (List1.Width + 100)
Text2.Width = Me.Width - (List1.Width + 100)
List1.Left = Me.Width - List1.Width - 100
Label1.Left = Me.Width - Label1.Width - 100
Text2.Top = Me.Height - Text2.Height - 500
Text1.Height = Me.Height - Text2.Height - 700
List1.Height = Me.Height - 700
End Sub

Private Sub Text1_Change()
List1.Clear
Call NewAddRoomToList(List1, True)
Call KillDupes(List1)
Label1 = "People: " & List1.ListCount
Label2 = "Room Name: " & GetCaption(FindChat)
End Sub

Private Sub Text1_GotFocus()
Text2.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendChat (Text2)
Text2 = ""
Text2.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
EnterPR (Text3)
Text3 = ""
Text3.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Call FormOnTop(Me)
If ChatBox = "" Then
Chat1.Scan_Off
List1.Clear
Label1 = "People: 0"
Label2 = "Room Name: 0"
Else
Chat1.Scan_On
Label1 = "People: " & List1.ListCount
Label2 = "Room Name: " & GetCaption(FindChat)
End If
End Sub
