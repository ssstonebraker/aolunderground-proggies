VERSION 2.00
Begin MDIForm MDIForm1 
   Caption         =   "MDI Form - Forms and Menus Project"
   ClientHeight    =   3660
   ClientLeft      =   735
   ClientTop       =   1545
   ClientWidth     =   6465
   Height          =   4095
   Icon            =   CHP04A.FRX:0000
   Left            =   660
   LinkTopic       =   "MDIForm1"
   Top             =   1185
   Width           =   6615
   Begin PictureBox pictStatus 
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C0C0&
      Height          =   372
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   6435
      TabIndex        =   1
      Top             =   3285
      Width           =   6468
   End
   Begin PictureBox pictToolbar 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   0
      Width           =   6465
      Begin PictureBox pictNewChild 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   2100
         Picture         =   CHP04A.FRX:0302
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   120
         Width           =   480
      End
      Begin PictureBox pictModal 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0C0C0&
         Height          =   480
         Left            =   1440
         Picture         =   CHP04A.FRX:0604
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   8
         Top             =   120
         Width           =   480
      End
      Begin PictureBox pictModeless 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   780
         Picture         =   CHP04A.FRX:0906
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   7
         Top             =   120
         Width           =   480
      End
      Begin PictureBox pictQuit 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   120
         Picture         =   CHP04A.FRX:0C08
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   6
         Top             =   120
         Width           =   480
      End
      Begin CommandButton Command4 
         Height          =   615
         Left            =   2040
         TabIndex        =   5
         Top             =   60
         Width           =   615
      End
      Begin CommandButton Command3 
         Height          =   615
         Left            =   1380
         TabIndex        =   4
         Top             =   60
         Width           =   615
      End
      Begin CommandButton Command2 
         Height          =   615
         Left            =   720
         TabIndex        =   3
         Top             =   60
         Width           =   615
      End
      Begin CommandButton Command1 
         Height          =   615
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   615
      End
   End
End

Sub MDIForm_Load ()
     Load formModal        'load the dialog boxes into memory
     Load formModeless     'to save time later
     'create a new document and an entry in the tracking array
     ReDim Document(0)     'this creates a new child form
     ReDim DocState(0)     'this creates the element to track the child form
     Document(0).Tag = 0   'store the document's index to tracking array
     Document(0).Show      'and display the new child
End Sub

Sub MDIForm_Unload (Cancel As Integer)
     Dim msg, answer
     msg = "Do you really want to quit?"
     answer = MsgBox(msg, 36, "Final Confirmation")  ' yes/no
     If answer = 7 Then
          Cancel = True
     Else
          End
     End If
End Sub

Sub pictModal_Click ()
     DisplayModal
End Sub

Sub pictModeless_Click ()
     DisplayModeless
End Sub

Sub pictNewChild_Click ()
     DisplayNewChild
End Sub

Sub pictQuit_Click ()
     Quit
End Sub

