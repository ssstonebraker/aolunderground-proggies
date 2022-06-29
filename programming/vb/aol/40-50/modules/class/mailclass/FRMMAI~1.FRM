VERSION 5.00
Begin VB.Form frmMass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Class Test"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   120
   End
End
Attribute VB_Name = "frmMass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Crisis
'Email Me At: MemorexChris@hotmail.com
'Sample form for the mailclass.
'Features: a simple method to send a list.

'Declares the object
Dim WithEvents Test As MailClass
Attribute Test.VB_VarHelpID = -1

Private Sub Form_Load()
    'makes the object into a new instance of the class.
    Set Test = New MailClass
    Test.SetMailPrefs
    Test.OpenMailBox
    Test.LoadMailList
    Test.SendChat "Type -send list to get a list!~!"
End Sub

Private Sub Test_NewChat(Text As String)
    'The objects event.
    'this is why you have to declare it with
    '"Withevents".
    If InStr(1, Text, "-send list") Then 'normally i don't do this, but this was a quicky :)
        sn$ = Left$(Text, InStr(1, Text, ":") - 1)
            For NextMail = 0 To Test.MailCount - 1
                thelist$ = thelist$ & Test.RetrieveMailString(NextMail) & vbCrLf
            Next NextMail
        Test.SendMail sn$, "The List", thelist$
    End If
End Sub

Private Sub Timer1_Timer()
    'this calls the scanchat method. And no this is not a
    'lastchatline.
    Test.ScanChat
End Sub
