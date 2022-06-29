VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Class Test"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   360
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Crisis
'Email Me At: memorexChris@hotmail.com
'Sample form for the scan class.
'Features: the scanchat code without a ocx.
'Also available in the Massmailer class.

'Declares the object
Dim WithEvents Test As ChatScanClass
Attribute Test.VB_VarHelpID = -1

Private Sub Form_Load()
    'makes the object into a new instance of the class.
    Set Test = New ChatScanClass
End Sub

Private Sub Test_NewChat(Text As String)
    'The objects event.
    'this is why you have to declare it with
    '"Withevents".
List1.AddItem Text
List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub Timer1_Timer()
    'this calls the scanchat method. And no this is not a
    'lastchatline. that's why the timer is set to 5000
    'you can see how it recursively adds things better
    'this way.
    Test.ScanChat
End Sub
