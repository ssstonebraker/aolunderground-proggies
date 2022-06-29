VERSION 2.00
Begin Form frmOptions 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "MVaders Game Options"
   ClientHeight    =   2580
   ClientLeft      =   690
   ClientTop       =   1530
   ClientWidth     =   7710
   Height          =   2985
   Icon            =   FRMOPTIO.FRX:0000
   Left            =   630
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   7710
   Top             =   1185
   Width           =   7830
   Begin CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   6660
      TabIndex        =   16
      Top             =   540
      Width           =   975
   End
   Begin CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   315
      Left            =   6660
      TabIndex        =   20
      Top             =   120
      Width           =   975
   End
   Begin HScrollBar hscChange 
      Height          =   240
      Index           =   7
      Left            =   1980
      Max             =   20
      Min             =   10
      TabIndex        =   15
      Top             =   2220
      Value           =   10
      Width           =   3555
   End
   Begin HScrollBar hscChange 
      Height          =   240
      Index           =   6
      Left            =   1980
      Max             =   15
      Min             =   5
      TabIndex        =   14
      Top             =   1920
      Value           =   5
      Width           =   3555
   End
   Begin HScrollBar hscChange 
      Height          =   240
      Index           =   5
      Left            =   1980
      Max             =   25
      Min             =   15
      TabIndex        =   13
      Top             =   1620
      Value           =   15
      Width           =   3555
   End
   Begin HScrollBar hscChange 
      Height          =   240
      Index           =   4
      Left            =   1980
      Max             =   16
      Min             =   6
      TabIndex        =   12
      Top             =   1320
      Value           =   6
      Width           =   3555
   End
   Begin HScrollBar hscChange 
      Height          =   240
      Index           =   3
      LargeChange     =   5
      Left            =   1980
      Max             =   95
      Min             =   70
      SmallChange     =   5
      TabIndex        =   11
      Top             =   1020
      Value           =   70
      Width           =   3555
   End
   Begin HScrollBar hscChange 
      Height          =   240
      Index           =   2
      Left            =   1980
      Max             =   5
      Min             =   3
      TabIndex        =   10
      Top             =   720
      Value           =   3
      Width           =   3555
   End
   Begin HScrollBar hscChange 
      Height          =   240
      Index           =   1
      LargeChange     =   5
      Left            =   1980
      Max             =   60
      Min             =   30
      SmallChange     =   5
      TabIndex        =   9
      Top             =   420
      Value           =   30
      Width           =   3555
   End
   Begin HScrollBar hscChange 
      Height          =   240
      Index           =   0
      LargeChange     =   5
      Left            =   1980
      Max             =   80
      Min             =   20
      SmallChange     =   5
      TabIndex        =   8
      Top             =   120
      Value           =   20
      Width           =   3555
   End
   Begin Label lblVal 
      Height          =   195
      Index           =   7
      Left            =   5700
      TabIndex        =   25
      Top             =   2220
      Width           =   735
   End
   Begin Label lblVal 
      Height          =   195
      Index           =   6
      Left            =   5700
      TabIndex        =   24
      Top             =   1920
      Width           =   735
   End
   Begin Label lblVal 
      Height          =   195
      Index           =   5
      Left            =   5700
      TabIndex        =   23
      Top             =   1620
      Width           =   735
   End
   Begin Label lblVal 
      Height          =   195
      Index           =   4
      Left            =   5700
      TabIndex        =   22
      Top             =   1320
      Width           =   735
   End
   Begin Label lblVal 
      Height          =   195
      Index           =   3
      Left            =   5700
      TabIndex        =   21
      Top             =   1020
      Width           =   735
   End
   Begin Label lblVal 
      Height          =   195
      Index           =   2
      Left            =   5700
      TabIndex        =   19
      Top             =   720
      Width           =   735
   End
   Begin Label lblVal 
      Height          =   195
      Index           =   1
      Left            =   5700
      TabIndex        =   18
      Top             =   420
      Width           =   735
   End
   Begin Label lblVal 
      Height          =   195
      Index           =   0
      Left            =   5700
      TabIndex        =   17
      Top             =   120
      Width           =   735
   End
   Begin Label lblPrompt 
      Alignment       =   1  'Right Justify
      Caption         =   "Player Fire Speed:"
      Height          =   195
      Index           =   7
      Left            =   60
      TabIndex        =   7
      Top             =   2220
      Width           =   1815
   End
   Begin Label lblPrompt 
      Alignment       =   1  'Right Justify
      Caption         =   "Player Speed:"
      Height          =   195
      Index           =   6
      Left            =   60
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin Label lblPrompt 
      Alignment       =   1  'Right Justify
      Caption         =   "Invader Down Step:"
      Height          =   195
      Index           =   5
      Left            =   60
      TabIndex        =   5
      Top             =   1620
      Width           =   1815
   End
   Begin Label lblPrompt 
      Alignment       =   1  'Right Justify
      Caption         =   "Invader Fire Speed:"
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin Label lblPrompt 
      Alignment       =   1  'Right Justify
      Caption         =   "Invader Fire Freq:"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   1020
      Width           =   1815
   End
   Begin Label lblPrompt 
      Alignment       =   1  'Right Justify
      Caption         =   "Invader Speed:"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin Label lblPrompt 
      Alignment       =   1  'Right Justify
      Caption         =   "Invader Spacing:"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   1815
   End
   Begin Label lblPrompt 
      Alignment       =   1  'Right Justify
      Caption         =   "Timer Setting:"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Option Explicit

Sub cmdCancel_Click ()

'Exit without making changes
Unload frmOptions

End Sub

Sub cmdOk_Click ()

'Set the values player has selected
GamePrefs.iTimer = hscChange(0).Value
GamePrefs.iIGap = hscChange(1).Value
GamePrefs.iISpeed = hscChange(2).Value
GamePrefs.fIBFreq = hscChange(3).Value / 100
GamePrefs.iIBSpeed = hscChange(4).Value
GamePrefs.iIDrop = hscChange(5).Value
GamePrefs.iPSpeed = hscChange(6).Value
GamePrefs.iPBSpeed = hscChange(7).Value

'Save the changes
SaveHiScore giHiScore, gsHiName

'And exit
Unload frmOptions

End Sub

Sub Form_Activate ()

'Initialise all the scrollers
hscChange(0).Value = GamePrefs.iTimer
hscChange(1).Value = GamePrefs.iIGap
hscChange(2).Value = GamePrefs.iISpeed
hscChange(3).Value = Int(GamePrefs.fIBFreq * 100)
hscChange(4).Value = GamePrefs.iIBSpeed
hscChange(5).Value = GamePrefs.iIDrop
hscChange(6).Value = GamePrefs.iPSpeed
hscChange(7).Value = GamePrefs.iPBSpeed

End Sub

Sub Form_Load ()

'Center form on the screen
CenterForm Me

End Sub

Sub hscChange_Change (Index As Integer)

lblVal(Index).Caption = Format$(hscChange(Index).Value, "")

End Sub

Sub hscChange_Scroll (Index As Integer)

lblVal(Index).Caption = Format$(hscChange(Index).Value, "")

End Sub

