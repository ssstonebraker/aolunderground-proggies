VERSION 2.00
Begin Form formMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reference Project"
   ClientHeight    =   3540
   ClientLeft      =   465
   ClientTop       =   1080
   ClientWidth     =   6420
   Height          =   3975
   Icon            =   CHP06A.FRX:0000
   Left            =   390
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   6420
   Top             =   720
   Width           =   6570
   Begin CommandButton cmndExit 
      Caption         =   "E&xit"
      Height          =   492
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   2892
   End
   Begin CommandButton cmndPrint 
      Caption         =   "&Print Letter"
      Height          =   972
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   2892
   End
   Begin Frame framOptions 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Letter Options"
      Height          =   3132
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   2892
      Begin CheckBox chekLetter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Carbon Copy &List"
         Height          =   192
         Index           =   6
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   2052
      End
      Begin CheckBox chekLetter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Enclosure List"
         Height          =   192
         Index           =   5
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   2532
      End
      Begin CheckBox chekLetter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Closing"
         Height          =   192
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   2532
      End
      Begin CheckBox chekLetter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Body"
         Height          =   192
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   2292
      End
      Begin CheckBox chekLetter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Greeting"
         Height          =   192
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2532
      End
      Begin CheckBox chekLetter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Addressee"
         Height          =   192
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2412
      End
      Begin CheckBox chekLetter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Return Address"
         Height          =   192
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2532
      End
   End
   Begin Frame framType 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Letter Type"
      Height          =   1452
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2892
      Begin OptionButton optnLetter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ge&neral Letter"
         Height          =   192
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   2292
      End
      Begin OptionButton optnLetter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Accep&tance Letter"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   2412
      End
      Begin OptionButton optnLetter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Introduction Letter"
         Height          =   192
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2172
      End
   End
End

Sub chekLetter_Click (Index As Integer)
     If TypeIndex <> -1 Then
          CurrentIndex = Index
          formEntry.Show
     End If
End Sub

Sub cmndExit_Click ()
    End
End Sub

Sub cmndPrint_Click ()
     formMain.Hide
     GenerateLetter
     formLetter.Show MODAL
     formLetter.Hide
     formMain.Show
End Sub

Sub Form_Load ()
     Load formEntry
     Load formLetter
End Sub

Sub optnLetter_Click (Index As Integer)
    HiLite Screen.ActiveControl
    TypeIndex = -1                            'Flag that we're setting type
    Select Case Screen.ActiveControl.Index
        Case 0                                'Introduction letter
          chekLetter(2).Enabled = False
          chekLetter(3).Enabled = False
          chekLetter(4).Enabled = False
          chekLetter(5).Enabled = False
          chekLetter(6).Enabled = False
          UnHiLite optnLetter(1)
          UnHiLite optnLetter(2)
        Case 1                                 'Acceptance letter
          chekLetter(2).Enabled = True
          chekLetter(3).Enabled = True
          chekLetter(4).Enabled = False
          chekLetter(5).Enabled = False
          chekLetter(6).Enabled = False
          UnHiLite optnLetter(0)
          UnHiLite optnLetter(2)
        Case 2                                 'General letter
          chekLetter(2).Enabled = True
          chekLetter(3).Enabled = True
          chekLetter(4).Enabled = True
          chekLetter(5).Enabled = True
          chekLetter(6).Enabled = True
          UnHiLite optnLetter(0)
          UnHiLite optnLetter(1)
     End Select
     For X = RTN_ADDRESS To CARBON             'Gray out checked boxes that are now disabled
          If chekLetter(X).Enabled = False And chekLetter(X).Value = CHECKED Then
               chekLetter(X).Value = GRAYED
          End If                               'Check boxes that are grayed out but now enabled
          If chekLetter(X).Enabled = True And chekLetter(X).Value = GRAYED Then
               chekLetter(X).Value = CHECKED
          End If
          DoEvents
     Next X
     TypeIndex = Screen.ActiveControl.Index
End Sub

