VERSION 5.00
Begin VB.Form frmAway 
   Caption         =   "Away Message"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3645
   Icon            =   "frmAway.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "&Hide Windows While I'm Away"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Away"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmAway.frx":030A
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmAway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = False Then
Call ShowForm
Else
Call HideForm
End If

End Sub

Private Sub Command1_Click()
Dim f$
If Command1.Caption = "&Away" Then
Text1.Enabled = False
Command1.Caption = "&I'm Back"
'Call SendProc(2, "toc_set_info " & Chr(34) & Normalize("<HTML>I'm using a VBaim example by Chad Cox, Thomas Grimshaw and Steve Nowiki, with Ragno Web Products. <HTML><BODY BGCOLOR=" & Chr(34) & "#ffffff" & Chr(34) & "><FONT COLOR=" & Chr(34) & "#00ff00" & Chr(34) & " FACE=" & Chr(34) & "Verdana" & Chr(34) & " SIZE=1>[<U></FONT><FONT COLOR=" & Chr(34) & "#ff8000" & Chr(34) & " FACE=" & Chr(34) & "Verdana" & Chr(34) & " SIZE=1>Away</U></FONT><FONT COLOR=" & Chr(34) & "#ff8000" & Chr(34) & " FACE=" & Chr(34) & "Verdana" & Chr(34) & " SIZE=1> <U></FONT><FONT COLOR=" & Chr(34) & "#ff8000" & Chr(34) & " FACE=" & Chr(34) & "Verdana" & Chr(34) & " SIZE=1>Message</U></FONT><FONT COLOR=" & Chr(34) & "#00ff00" & Chr(34) & " FACE=" & Chr(34) & "Verdana" & Chr(34) & " SIZE=1>]</FONT></BODY></HTML>") & Chr(34) & Chr(0))
f$ = "<HTML><BODY BGCOLOR=" & Chr(34) & "#ffffff" & Chr(34) & "><FONT COLOR=" & Chr(34) & "#00ff00" & Chr(34) & " FACE=" & Chr(34) & "Verdana" & Chr(34) & " SIZE=1>[<U></FONT><FONT COLOR=" & Chr(34) & "#ff8000" & Chr(34) & " FACE=" & Chr(34) & "Verdana" & Chr(34) & " SIZE=1>Away</U></FONT><FONT COLOR=" & Chr(34) & "#ff8000" & Chr(34) & " FACE=" & Chr(34) & "Verdana" & Chr(34) & " SIZE=1> <U></FONT><FONT COLOR=" & Chr(34) & "#ff8000" & Chr(34) & " FACE=" & Chr(34) & "Verdana" & Chr(34) & " SIZE=1>Message</U></FONT><FONT COLOR=" & Chr(34) & "#00ff00" & Chr(34) & " FACE=" & Chr(34) & "Verdana" & Chr(34) & " SIZE=1>]</FONT></BODY></HTML><HTML>: " & Text1.Text & "</HTML><HTML><P></P></HTML>"

Call SendProc(2, "toc_set_away " & Chr(34) & Normalize(f$) & Chr(34) & Chr(0))

MeAway = True
Else
Text1.Enabled = True
Call SendProc(2, "toc_set_away " & Chr(34) & Chr(34) & Chr(0))
'Command1.Caption = "&Away"
Call ShowForm
MeAway = False
Unload Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call SendProc(2, "toc_set_away " & Chr(34) & Chr(34) & Chr(0))
'Command1.Caption = "&Away"
Call ShowForm
MeAway = False
Unload frmAway
End Sub
