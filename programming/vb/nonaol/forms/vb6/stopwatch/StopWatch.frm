VERSION 5.00
Begin VB.Form frmStopWatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StopWatch"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   1995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Sto&p"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   960
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   " &Start "
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblTimeClock 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'##########################################################
'There's only three lines of code in this whole form.     #
'The first is put in the command button to tell the timer #
'to start. The second is in the timer, which tells it to  #
'run the sub that is in the bas file. Third is put in the #
'other command button to stop the timer.                  #
'                                                         #
'Comments or questions can be sent to me at               #
'NightMare_36@hotmail.com                                 #
'                                                         #
'    by :         NightShade    04/26/1999                #
'##########################################################


 
Private Sub cmdStart_Click()

    Timer1.Enabled = True 'Starts the timer

End Sub

Private Sub cmdStop_Click()

    Timer1.Enabled = False 'Stops the timer

End Sub

Private Sub Timer1_Timer()

    Call StopWatch(lblTimeClock) 'Starts the code in the bas file

End Sub
