VERSION 2.00
Begin Form options 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   1164
   ClientLeft      =   3168
   ClientTop       =   2616
   ClientWidth     =   3000
   Height          =   1500
   Icon            =   OPTIONS.FRX:0000
   Left            =   3120
   LinkTopic       =   "Form2"
   ScaleHeight     =   1164
   ScaleWidth      =   3000
   Top             =   2328
   Width           =   3096
   Begin TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "(upsscheck@collegemail.com, upsscheck@cyberdude.com)"
      Top             =   1200
      Width           =   1695
   End
   Begin TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin CheckBox Check4 
      BackColor       =   &H00000000&
      Caption         =   "Log The MM"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin DBPush DBPush4 
      Alignment       =   2  'Center
      AutoPictureDisabled=   0  'Off
      BackColor       =   &H00C0C0C0&
      BackColorDown   =   &H00C0C0C0&
      BevelWidth      =   2
      Caption         =   "Exit"
      CornerColor     =   &H00000000&
      FocusRect       =   0   'False
      Font3D          =   0  'None
      Font3DDisabled  =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForceDown       =   0   'False
      ForceUp         =   0   'False
      ForeColor       =   &H00E0E0E0&
      ForeColorDisabled=   &H00000000&
      Height          =   255
      Left            =   1680
      Outline         =   -1  'True
      OutlineSize     =   1
      Picture         =   OPTIONS.FRX:0302
      PictureDown     =   OPTIONS.FRX:126E
      PictureFromEdge =   0
      Placement       =   4  'Text on top of picture
      PredefPicture   =   0  'None
      Style           =   0  'Normal
      TabIndex        =   3
      Top             =   720
      TransColor      =   &H00000000&
      Transparent     =   0   'False
      VertAlignment   =   2  'Center
      Width           =   1095
   End
   Begin CheckBox Check3 
      BackColor       =   &H00000000&
      Caption         =   "Turn IMs Off"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "45 Minute Kill"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Close Chatroom"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "CheckBox:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "MassMail #"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Nickname:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
End
' OPTIONS.FRM
Option Explicit

Sub Check4_Click ()
If  Text1.Enabled = True Then
options.Height = 1590
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Exit Sub
End If
If  Text1.Enabled = False Then
options.Height = 2700
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
MsgBox "This is the MassMail Auto-Logger.  It automatically sends a log to the checkbox. ********WARNING******** THIS CAN NOT BE FORGED IN ANY WAY AT ALL!!! IF YOU ARE CAUGHT SENDING IN A FAKE LOG, YOU WILL BE KICKED OUT OF UPSS!  It is IMPOSSIBLE to duplicate this log, because there is invisible characters that are entered into it, and are invisible in AOL mail, but can be seen by the program I made for them, and in which they know where the characters are, and can tell if its forged.....................Anyways........    Enter your Nickname (ex) " + Chr(34) + "EvicTioN" + Chr(34) + ", and the MassMail #, (ex) " + Chr(34) + "6-10" + Chr(34) + ", and if neccessary, to change the checkbox."
Exit Sub
End If
End Sub

Sub DBPush4_Click ()
Me.Hide
End Sub

Sub Form_Activate ()
sub0193 Me
End Sub

Sub Form_Load ()
Dim l0032 As Variant
sub0193 Me
l0032 = fn0141()
If  l0032 = "" Then Exit Sub
Text2.Text = l0032
End Sub
