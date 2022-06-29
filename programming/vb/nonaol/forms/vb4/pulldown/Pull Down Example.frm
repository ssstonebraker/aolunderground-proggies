VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Pull Down Example by Goblin Jester"
   ClientHeight    =   1530
   ClientLeft      =   3840
   ClientTop       =   1860
   ClientWidth     =   3585
   Height          =   2220
   Left            =   3780
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   3585
   Top             =   1230
   Width           =   3705
   Begin VB.CommandButton Command7 
      Caption         =   "Don't understand the numbers?"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Menus on other Forms"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Bold Menu Items"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Right Click Menu"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "about menus"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Simple Menu"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Better Menu"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu other 
      Caption         =   "Other"
      Begin VB.Menu options 
         Caption         =   "Options"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
'TheFormToPopUpOn.Popupmenu TheFormThatTheMenuIsOn.TheMenuName
'thats the basic part, to get it where you want it is where the numbers come in
'the first number is where it starts out at the bottom of the button, 0 will make it the left
'the second is the left position of the button/label (go to properties and look at the number that says left
'the last is how far from the top it is, this is the only tricky one, you take the position from the top (go to properties:top) and add the button's/lable's height to it (properties: height) and you get the last number
'in this case its top = 120, its height = 255, which equals 365
'even if you hide the menu item so they don't see it on your form, it can still be a pull down menu, so experiment =P
'to hide the menu, goto the menu in the menu editor and uncheck visible
Form1.PopupMenu Form1.file, 0, 120, 365

'there is an easier code than that, its on Simple Menu
End Sub


Private Sub Command2_Click()
'This is a simple version of the first, but it doesn't look as good
'All you have to do is put: PopupMenu TheFormTheMenuIsOn.MenuName
'but this box pops up where ever the user clicks, not at a certain spot, so i don't like it as much =P  its good for simple projects, but if you want your prog to look good, use the one listed on Better Menu
'even if you hide the menu item so they don't see it on your form, it can still be a pull down menu, so experiment =P
'to hide the menu, goto the menu in the menu editor and uncheck visible
PopupMenu Form1.other
End Sub


Private Sub Command3_Click()
'i have a lot of people ask me about making lines in the menu, like there is before exit on this one
'its quite simple, just put:- as the name, just '-' without the ' and you have a line, and just name it somethin like line1 and you set =P
MsgBox "Look at the Coding for the code =P", vbOKOnly, "Coding"
End Sub


Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This is the same as the other menus except it pops up when the user right clicks on the button/label
'all you have to do make this is add: If Button = 2 Then - before the code and: End If - after the code
'and its in the MouseDown section, instead of click
If Button = 2 Then
    Form1.PopupMenu Form1.file, 0, 120, 1095
End If
End Sub


Private Sub Command5_Click()
'This is the same as the Better Menu or the Simple Menu, but you put a comma (,) and then the NAME of the menu item that you want to be bold
Form1.PopupMenu Form1.file, 0, 1680, 1095, about
End Sub


Private Sub Umm_Click()

End Sub


Private Sub Command6_Click()
'The menu doesn't have to be on the form that its poping up on, if you had it on form2 and wanted it to be on this button it would be:
Form1.PopupMenu Form2.help, 0, 1680, 735
'you could also do a simple menu: PopupMenu Form2.help
End Sub


Private Sub Command7_Click()
'On the Better Menu there are three sets of numbers, 0,120, and 165
'these represent the values of the button/label
'leave the first one alone and change the second two
'2nd = the left position, to get that, select the button and go to the properties menu, scroll down to left, and thats the number
'3rd = where the menu stars from the top, so you take the height of the button/label and add it to the top (which are both found on the properties menu)
'got it?  if not e-mail me at goblinjester@dragonzlair.cjb.net
MsgBox "Look at the Coding for the code =P", vbOKOnly, "Coding"
End Sub


