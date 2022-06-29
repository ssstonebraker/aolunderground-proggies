VERSION 2.00
Begin Form formWarning 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   ClientHeight    =   2670
   ClientLeft      =   1155
   ClientTop       =   4245
   ClientWidth     =   4830
   Height          =   3105
   Left            =   1080
   LinkTopic       =   "Form2"
   ScaleHeight     =   2670
   ScaleWidth      =   4830
   Top             =   3885
   Width           =   4980
   Begin CommandButton cmndCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   492
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
      Width           =   1572
   End
   Begin CommandButton cmndRedial 
      Caption         =   "ReDial"
      Height          =   492
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   1932
   End
   Begin Timer timeResults 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   1560
   End
   Begin Timer timeDialing 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   1560
   End
   Begin Label lablResults 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   372
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4452
   End
   Begin Label lablDial 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4332
   End
End

Sub cmndCancel_Click ()
    lablDial.Caption = ""                 'Blank out Dialing label
    lablResults.Caption = ""              'and Results label
    Unload formWarning                    'Unload the form
    'you might need to change the path in loadpicture to fit your system
    formAppear.Icon = LoadPicture("\VB\ICONS\COMM\PHONE01.ICO")
    formAppear.WindowState = NORMAL       'and put back the main form
End Sub

Sub cmndRedial_Click ()
     Screen.MousePointer = HOURGLASS                 'Tell user we're going to be a while...
     lablResults.Visible = False                     'Hide the results box (to simulate dialing)
     Select Case lablDial.Caption                    'Determine the proper message from what
                                                     '    we've already got
          Case ""                                    'First time through...
               lablDial.Caption = "Redialing Number"
          Case "Redialing Number"                    'Second time through...
               lablDial.Caption = "Redialing Number Again!"
          Case Else                                  'Third time through...
               lablDial.Caption = "Redialing Number one more time!"
               lablDial.BackStyle = SOLID
               lablDial.BackColor = RED
               lablDial.ForeColor = YELLOW
     End Select
     timeDialing.Enabled = True                      'Set the time to simulate a "dialing" delay
End Sub

Sub Form_Load ()
     timeDialing.Enabled = True               'Simulate a "dialing" delay
     timeResults.Enabled = False              'turn off the results loop
End Sub

Sub timeDialing_Timer ()
    timeDialing.Enabled = False              'Done with dialing delay...
    timeResults.Enabled = True               'turn on results loop
End Sub

Sub timeResults_Timer ()
     lablResults.Visible = True                      'Display the results box
     Select Case lablResults.Caption                 'Determine message based on
                                                     '   what we've already done
          Case ""                                    'First time through...
               lablResults.Caption = "No Response"
          Case "No Response"                         'Second time through...
               lablResults.Caption = "Sorry, No Response"
          Case "Sorry, No Response"                  'Third time through...
               lablResults.Caption = "Sorry, Still No Response"
               lablResults.BackStyle = SOLID
               lablResults.BackColor = YELLOW
          Case Else                                  'Pretty persistent, eh?
               lablResults.Caption = "Hey what do you expect? This is a demonstration!"
               lablResults.BackColor = RED
               lablResults.ForeColor = YELLOW
     End Select
     Screen.MousePointer = DEFAULT                    'Done dialing, OK for user to go on
     timeResults.Enabled = False                      'Turn off the results loop
End Sub

