VERSION 2.00
Begin Form frmMain 
   Caption         =   "Toolbars, on the fly"
   ClientHeight    =   450
   ClientLeft      =   1095
   ClientTop       =   3195
   ClientWidth     =   7365
   Height          =   1140
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleWidth      =   7365
   Top             =   2565
   Width           =   7485
   Begin PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   0
      Width           =   7365
      Begin Image Toolbar 
         Height          =   330
         Index           =   0
         Left            =   30
         Tag             =   "1 Save"
         Top             =   30
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin Menu mnuFile 
      Caption         =   "&File"
      Begin Menu mnuFItems 
         Caption         =   "&New"
         Index           =   0
      End
      Begin Menu mnuFItems 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin Menu mnuFItems 
         Caption         =   "&Save"
         Index           =   2
      End
   End
   Begin Menu mnuOptions 
      Caption         =   "&Options"
   End
End
Option Explicit

Sub mnuFItems_Click (index As Integer)

    ' All the menu items are stored in a control array, so when one is clicked we need to check
    ' its Index to determine exactly what the user wants to do (Create a new file, Open, Save and
    ' so on).

    Select Case index
        
        Case 0      ' Case 0, item 0 in the array is New file
            MsgBox "Creating a new file"

        Case 1      ' Case 1, item 1 in the array is open a file (load it up)
            MsgBox "Opening an existing file"

        Case 2      ' Case 2, item 2 in the array is save a file
            MsgBox "Saving the file"
    
    End Select

End Sub

Sub mnuOptions_Click ()
    
    ' When the options menu item is selected, we need to load and show the options
    ' form, which the user can use to determine which buttons are going to appear on the
    ' toolbar.

    Load frmOptions
    frmOptions.Show 1   ' Show 1 shows the options form as a modal form.

End Sub

Sub Toolbar_Click (index As Integer)

    ' When a button on the toolbar is clicked, the number from the tag property should be
    ' copied across to the menu click code

    Call mnuFItems_Click(Val(Left(Toolbar(index).Tag, 1)))

    ' So if the number in Tag was 0 and we pass this to the menu click code, the Select Case
    ' code in that event runs the code to create a new file - take a look to see for yourself.

End Sub

Sub Toolbar_MouseDown (index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Load this button's down image.
    Call LoadImage(Toolbar(index), "dn")

End Sub

Sub Toolbar_MouseUp (index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Load this button's UP image into the button.
    Call LoadImage(Toolbar(index), "up")

End Sub

