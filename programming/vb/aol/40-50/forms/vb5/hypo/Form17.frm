VERSION 5.00
Begin VB.Form Form17 
   Caption         =   "Form17"
   ClientHeight    =   780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4710
   LinkTopic       =   "Form17"
   ScaleHeight     =   780
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Tree Macro"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Tea$ = Text1
' fnt$ = "10"
'a = Len(Tea$)
'For w = 1 To a Step 4
'    R$ = Mid$(Tea$, w, 1)
'    u$ = Mid$(Tea$, w + 1, 1)
'    s$ = Mid$(Tea$, w + 2, 1)
'    T$ = Mid$(Tea$, w + 3, 1)
'    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup><FONT SIZE=" + fnt$ + "><b>" & R$ & "</sup></font></b>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
'Next w
'WavYChaTRBb = P$
SendChat "A Partridge in a Pear Tree"
SendChat ""
timeout 0.9
SendChat "       _"
timeout 0.9
SendChat "      ('>"
timeout 0.9
SendChat "      /))@@@@@"
timeout 0.9
SendChat "     /@@@@@@()@"
timeout 0.9
SendChat "   @@()@@()@@@@"
timeout 0.9
SendChat "  @@@O@@@@()@@@"
timeout 0.9
SendChat "   @()@@\@@@()@@"
timeout 0.9
SendChat "     @()@||@@@@@"
timeout 0.9
SendChat "       @@||@@@"
timeout 0.9
SendChat "         ||"
timeout 0.9
SendChat "    ^^^^^^^^^ToaST^^^^^^^^^^^^^^^^^^^^^"


End Sub

