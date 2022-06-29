VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000B&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Example - Azazel"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3315
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear List"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.ListBox lstItems 
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblDonen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   8
      Top             =   840
      Width           =   90
   End
   Begin VB.Label lblLeftn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   7
      Top             =   480
      Width           =   90
   End
   Begin VB.Label lblTotaln 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   90
   End
   Begin VB.Label lblDone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Done:"
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   840
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Left:"
      Height          =   195
      Left            =   1680
      TabIndex        =   4
      Top             =   480
      Width           =   720
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total: "
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
lblTotaln.Caption = lstItems.ListCount
lblLeftn.Caption = lstItems.ListCount
lblDonen.Caption = "0"
pb1 = 0

' Starts clearing the list as it's using the progressbar:
For i = 0 To lstItems.ListCount - 1
    X = lstItems.ListIndex
    lstItems.RemoveItem (Index)
    lblLeftn.Caption = Val(Str(lblLeftn.Caption - 1))
    lblDonen.Caption = Val(Str(lblDonen.Caption + 1))
    pb1 = Percent(lblDonen, lblTotaln, 100)
If pb1 = 100 Then GoTo Endis
Pause 1
Next i

' Progressbar = 100, and the list is all cleared
Endis:
Pause 0.2
MsgBox "Clear list complete.", vbInformation, "- Progressbar Example by Azazel"
pb1 = 0
Do
DoEvents
Loop
End Sub

Private Sub Form_Load()
' The following are list items that will be removed when the
' Clear List button is clicked.  It will do a timeout of 1 in between
' each name so you can see how to progressbar werks and etc..
lstItems.AddItem "a"
lstItems.AddItem "b"
lstItems.AddItem "c"
lstItems.AddItem "d"
lstItems.AddItem "e"
lstItems.AddItem "f"
lstItems.AddItem "g"
lstItems.AddItem "h"
lstItems.AddItem "i"
lstItems.AddItem "j"
lstItems.AddItem "k"
lstItems.AddItem "l"
lstItems.AddItem "m"
lstItems.AddItem "n"
lstItems.AddItem "o"
lstItems.AddItem "p"
lstItems.AddItem "q"
lstItems.AddItem "r"
lstItems.AddItem "s"
lstItems.AddItem "t"
lstItems.AddItem "u"
lstItems.AddItem "v"
lstItems.AddItem "w"
lstItems.AddItem "x"
lstItems.AddItem "y"
lstItems.AddItem "z"

lblTotaln = lstItems.ListCount
End Sub


