VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pop-up menu - mash"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Menu on other form"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
      Begin VB.CommandButton Command2 
         Caption         =   "Right mouse click"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Left mouse click"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Menu on this form"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "Left mouse click"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Right mouse click"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Menu PopupMain 
      Caption         =   "Popup Example"
      Visible         =   0   'False
      Begin VB.Menu Popup 
         Caption         =   "Popup Example"
      End
      Begin VB.Menu By 
         Caption         =   "by"
      End
      Begin VB.Menu Mash 
         Caption         =   "mash"
      End
      Begin VB.Menu site 
         Caption         =   "http://freez.com/mash"
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
PopupMenu PopupMain, 1, Command1.Left, Command1.Top + Command1.Height, site
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MnuFrm.PopUpNotOnThisForm, 1, 550, 1950, MnuFrm.isnot
    End If
End Sub

Private Sub Form_Load()
'PopUpMenu example by mash.   -    air_ware@hotmail.com
'http://freez.com/mash

'Sup, this is my example on Pop up menu's.
'The main thing you have to know about
'pop up menu's is this.

    'object.PopupMenu menuname, flags, x, y, boldcommand

'Firstly the object, this is optional, its the object you want the menu to
'pop up on.  If it is left out it is assumed that the object is the one
'with focus.

'Next menuname, this is needed and is the name of the menu you
'want to pop up.

'Flags: Optional, this is like the behavior and/or location of the menu.
'I used "1" because the menus i was using were the first in line.

'X: Optional, also known as left/right, its the location of where you want
'the menu to pop up.  If its left out the mouse location is used.

'Y: Optional, this is the same as X exept the location applies to yep
'you guessed it up/down.

'Lastly BoldCommand, this is also optional and makes a selected item
'in the menu displayed as bold text.

'thanks for d/l'ing my example.
'-mash
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu PopupMain, 1, Label1.Left, Label1.Top + Label1.Height, Mash
    End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        PopupMenu MnuFrm.PopUpNotOnThisForm, 1, 550, 2250, MnuFrm.mainform
    End If
End Sub
