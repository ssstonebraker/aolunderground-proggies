VERSION 2.00
Begin Form frmOptions 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Which Buttons"
   ClientHeight    =   2280
   ClientLeft      =   1185
   ClientTop       =   2445
   ClientWidth     =   4770
   ControlBox      =   0   'False
   Height          =   2685
   Left            =   1125
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4770
   Top             =   2100
   Width           =   4890
   Begin CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   420
      Left            =   3270
      TabIndex        =   5
      Top             =   1710
      Width           =   1410
   End
   Begin CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   135
      TabIndex        =   4
      Top             =   1710
      Width           =   1410
   End
   Begin Frame Frame1 
      Caption         =   "Select those buttons you would like on the toolbar"
      Height          =   1530
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   4545
      Begin CheckBox chkOptions 
         Caption         =   "Save the current file"
         Height          =   240
         Index           =   2
         Left            =   1095
         TabIndex        =   3
         Tag             =   "2 Save"
         Top             =   1155
         Width           =   2340
      End
      Begin CheckBox chkOptions 
         Caption         =   "Open an existing file"
         Height          =   240
         Index           =   1
         Left            =   1095
         TabIndex        =   2
         Tag             =   "1 Open"
         Top             =   795
         Width           =   2340
      End
      Begin CheckBox chkOptions 
         Caption         =   "Create a new File"
         Height          =   240
         Index           =   0
         Left            =   1095
         TabIndex        =   1
         Tag             =   "0 New"
         Top             =   435
         Width           =   2340
      End
   End
End
Option Explicit

Sub cmdCancel_Click ()
    
    ' Clicking the Cancel button gets rid of the options form
    Unload frmOptions

End Sub

Sub cmdOK_Click ()

    Dim nButtonNo As Integer

    ' When the OK button is clicked on the Options form, any existing toolbar buttons are
    ' unloaded. A loop is then used to go through the checkbox array on this form and
    ' create the necessary buttons on the toolbar form.


    ' Loop through the existing buttons and delete them
    For nButtonNo = 1 To gnButtons      ' gnButtons is a global variable which tells us how
                                        ' many buttons there currently are. It is declared in
                                        ' toolbar.bas.

        ' Unloading an element of a control array, like this, effectively deletes it.
        Unload frmMain!Toolbar(nButtonNo)

    Next

    ' Now all the buttons have been unloaded, we can set the global button count variable back
    ' down to 0.
    gnButtons = 0

    ' Loop through the checkboxes (which are all in the same control array) to see which are
    ' checked and which are not.
    For nButtonNo = 0 To 2

        ' If the checkbox is checked (value = 1) then ....
        If chkOptions(nButtonNo).Value = 1 Then

            ' Add 1 to the number of buttons on the toolbar.
            gnButtons = gnButtons + 1

            ' Load a new element into the toolbar buttons control array. Since we already have
            ' a global variable which holds the number of buttons on the toolbar, we can use
            ' this variable to tell Visual Basic the index number of the new button.
            Load frmMain!Toolbar(gnButtons)

            ' The tag properties of the option form checkboxes hold the details about which
            ' graphic the buttons should display, and which menu option the buttons emulate.
            ' The Tag property is copied across from the checkbox to the new button.
            frmMain!Toolbar(gnButtons).Tag = chkOptions(nButtonNo).Tag

            ' The Top and Left properties of the new button are now set to position the button
            ' directly to the right of the previous button. Remember, there is always 1
            ' invisible button on the toolbar, from which the new buttons are created. If the
            ' user wants to display only one button on the toolbar then it is positioned to the
            ' right of the invisible button.
            frmMain!Toolbar(gnButtons).Top = frmMain!Toolbar(0).Top
            frmMain!Toolbar(gnButtons).Left = frmMain!Toolbar(gnButtons - 1).Left + frmMain!Toolbar(gnButtons - 1).Width
            
            ' Next, it is important to load the graphic image of the new button. A routine in
            ' Toolbar.Bas does this for us, using the filename in the Tag property to determine
            ' the name of the graphic to load.
            Call LoadImage(frmMain!Toolbar(gnButtons), "up")
            
            ' Now that the new button has been positioned properly and its graphic loaded, we
            ' can at last set the Visible property to True to make it appear.
            frmMain!Toolbar(gnButtons).Visible = True
        End If
    Next
    
    ' The final thing to do, after setting up all the toolbar buttons is to get rid of the
    ' options form and return to the previous form. A simple Unload should suffice!
    Unload frmOptions

End Sub

Sub Form_Load ()

    ' When the options form is loaded up, it is nice to have it positioned in the centre of the
    ' screen. The following two lines set up the form's Left and Top properties so that this
    ' happens.
    frmOptions.Left = (screen.Width - frmOptions.Width) / 2
    frmOptions.Top = (screen.Height - frmOptions.Height) / 2

End Sub

