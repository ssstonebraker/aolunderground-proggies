VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "dlutilities"
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "y"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu adv 
         Caption         =   "advertise"
      End
      Begin VB.Menu ccccccccccccccc 
         Caption         =   "continue download when signed back on?"
      End
      Begin VB.Menu sd 
         Caption         =   "go to a specific room when signed back on?"
      End
      Begin VB.Menu ikfsdaolasdffasd 
         Caption         =   "-"
      End
      Begin VB.Menu cc 
         Caption         =   "contact"
      End
      Begin VB.Menu exit 
         Caption         =   "exit"
      End
      Begin VB.Menu help 
         Caption         =   "help"
      End
      Begin VB.Menu minimize 
         Caption         =   "minimize"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adv_Click()
mIRCSendChat ("•.·—› download utilities v2 by koko")
AOL40SendChat ("•.·—› download utilities v2 by koko")
TimeOut (0.3)
mIRCSendChat ("•.·—› current hacker: ") + GetUser
AOL40SendChat ("•.·—› current hacker: ") + GetUser
TimeOut (0.3)
mIRCSendChat ("•.·—› get it at http://www.angelfire.com/yt/koko")
AOL40SendChat ("< A HREF=") + "http://www.angelfire.com/yt/koko" + (">") + ("•.·—›") + (" http://www.angelfire.com/yt/koko") + ("</A>")
End Sub

Private Sub cc_Click()
MsgBox "All mail should go to k0ko@hotmail.com, report any bugs", vbCritical
End Sub

Private Sub ccccccccccccccc_Click()

Dim strAdd As String
strAdd = InputBox("Would you like to continue downloading when you sign back on?.  *NOTE*: Type y for yes and n for no!  Do not type yes or no!", "download utilities v2.0.")

If TrimSpaces(strAdd) = "" Then
Exit Sub
Else

Label1.Caption = (strAdd)

End If
End Sub

Private Sub exit_Click()
mIRCSendChat ("•.·—› download utilities v2 by koko")
AOL40SendChat ("•.·—› download utilities v2 by koko")
TimeOut (0.3)
mIRCSendChat ("•.·—› unloaded by: ") + GetUser
AOL40SendChat ("•.·—› unloaded by: ") + GetUser
TimeOut (0.3)
mIRCSendChat ("•.·—› get it at http://www.angelfire.com/yt/koko")
AOL40SendChat ("< A HREF=") + "http://www.angelfire.com/yt/koko" + (">") + ("•.·—›") + (" http://www.angelfire.com/yt/koko") + ("</A>")
End
End Sub

Private Sub help_Click()
MsgBox "The start auto sign on signs back on when you get logged off lets say at night while your downloading and continues your download on aol, uh thats about it if you can't figure out how to use it your probably a dumbass", vbCritical
End Sub

Private Sub minimize_Click()
Form1.WindowState = vbMinimized
End Sub

Private Sub sd_Click()

Dim strAdd As String
strAdd = InputBox("Would you like to go to a specific private room?  Enter the room name below.", "download utilities v2.0.")

If TrimSpaces(strAdd) = "" Then
Exit Sub
Else

Label2.Caption = (strAdd)

End If
End Sub
