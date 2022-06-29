VERSION 2.00
Begin Form formAppear 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Appearance Project"
   ClientHeight    =   1965
   ClientLeft      =   855
   ClientTop       =   1770
   ClientWidth     =   5520
   Height          =   2400
   Icon            =   CHP05A.FRX:0000
   Left            =   780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5520
   Top             =   1410
   Width           =   5670
   Begin Timer timeDial 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   2280
      Top             =   1080
   End
   Begin CommandButton cmndExit 
      Caption         =   "E&xit"
      Height          =   492
      Left            =   3840
      TabIndex        =   2
      Top             =   1320
      Width           =   1572
   End
   Begin CommandButton cmndDial 
      Caption         =   "&Dial"
      Height          =   852
      Left            =   3840
      MousePointer    =   12  'No Drop
      TabIndex        =   1
      Top             =   240
      Width           =   1572
   End
   Begin ListBox listPhone 
      Height          =   1590
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End

Sub cmndDial_Click ()
     If cmndDial.MousePointer <> NO_DROP Then                'Make sure we're supposed to dial
          Screen.MousePointer = HOURGLASS                    'Tell user we've got to wait
          formWarning.Caption = "Dialing " + listPhone.List(listPhone.ListIndex)
          'you might need to change the path in loadpicture to fit your system
          formAppear.Icon = LoadPicture("\VB\ICONS\COMM\PHONE12.ICO")  'Phone Handset
          formAppear.WindowState = MINIMIZED                 'Minimize Appear form
          timeDial.Enabled = True                            'Turn on dialing timer
     End If
End Sub

Sub cmndExit_Click ()
     End
End Sub

Sub Form_Load ()
    timeDial.Enabled = False                           'Make sure dialing timer is disabled
    listPhone.AddItem "517-555-1212     John Doe"      'Fill list box with fake-o numbers
    listPhone.AddItem "908-789-8901     Jane Mazurski"
    listPhone.AddItem "313-531-7909     Fred Dopfler"
    listPhone.AddItem "206-908-8973     Alan Goodson"
End Sub

Sub listPhone_Click ()
     cmndDial.MousePointer = DEFAULT   'We've selected number to dial; OK to push dial button
End Sub

Sub timeDial_Timer ()
    formWarning.Show MODAL                       'Let the Dialer form take over
    formAppear.Caption = "Appearance Project"    'Reset caption
    Screen.MousePointer = DEFAULT                'Set cursor to normal
    listPhone.ListIndex = -1                     'Deselect list
    cmndDial.MousePointer = NO_DROP              'We can't dial now...
    timeDial.Enabled = False                     'Turn off dialing timer
End Sub

