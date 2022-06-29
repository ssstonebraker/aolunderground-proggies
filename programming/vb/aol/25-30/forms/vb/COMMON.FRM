VERSION 2.00
Begin Form frmcommdlg 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Common Dialogs "
   ClientHeight    =   2115
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   6525
   ControlBox      =   0   'False
   Height          =   2520
   Left            =   1035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6525
   Top             =   1140
   Width           =   6645
   Begin CommonDialog dlgDialog 
      Left            =   3045
      Top             =   1395
   End
   Begin CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5235
      TabIndex        =   5
      Top             =   1515
      Width           =   1215
   End
   Begin CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   5220
      TabIndex        =   4
      Top             =   195
      Width           =   1215
   End
   Begin CommandButton cmdFont 
      Caption         =   "&Font"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   195
      Width           =   1215
   End
   Begin CommandButton cmdColour 
      Caption         =   "&Colour"
      Height          =   495
      Left            =   2685
      TabIndex        =   2
      Top             =   195
      Width           =   1215
   End
   Begin CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   1395
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   195
      Width           =   1215
   End
   Begin CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   1215
   End
End
Option Explicit

Sub cmdColour_Click ()

    ' When the colour button is pressed the common dialog control is set up to display the
    ' colour dialog; 1 is loaded into the Flags property to initialise the control, then the
    ' Action code of 3 actually triggers the dialog itself.

    ' On error is used to trigger the error handling that will kick in if the user hits cancel
    On Error GoTo Colour_Error
    
    ' Set the CancelError property to True causes a trappable error to occur if the user hits
    ' the cancel button on the dialog
    dlgDialog.CancelError = True

    dlgDialog.Flags = 1
    dlgDialog.Action = 3

    Exit Sub

Colour_Error:
    ' This line turns off the error trapping
    On Error GoTo 0

    MsgBox "You hit cancel - no colour was selected."

    Exit Sub

End Sub

Sub cmdExit_Click ()

    ' Unload the form to close the application when the exit button is clicked.

    Unload Me

End Sub

Sub cmdFont_Click ()

    ' Clicking the Font button causes the Fonts common dialog to appear.
    ' In this example we want to see the fonts for both screen and printer
    ' appear so a value of 3 is loaded into the Flags property.
    ' The Action code of 4 actually displays the dialog.

    ' This line tells VB where to go (in the nicest possible way)
    ' when a trappable error occurs. The CancelError property is
    ' set so that if the user clicks on the Cancel button in the
    ' dialog a trappable error occurs

    On Error GoTo Font_Error
    dlgDialog.CancelError = True

    dlgDialog.Flags = 3
    dlgDialog.Action = 4

    Exit Sub

Font_Error:
    
    ' Turn off the error trapping
    On Error GoTo 0

    MsgBox "You hit cancel - no font was selected."

    Exit Sub

End Sub

Sub cmdOpen_Click ()

    ' Clicking the Open command button causes the File open dialog to appear.

    On Error GoTo OpenError

    ' Setting CancelError to True causes a trappable error to occur if the user hits Cancel
    dlgDialog.CancelError = True

    ' The Filter is now set to specify those groups of files that our users will be most
    ' interested in using the dialog to locate
    dlgDialog.Filter = "Executeables (*.exe)|*.exe|Com Files (*.com)|*.com|Batch Files (*.bat)|*.bat"

    ' FilterIndex is set to 1, meaning that the first part of the filter will be used as the
    ' default, in this case *.EXE
    dlgDialog.FilterIndex = 1


    dlgDialog.DialogTitle = "Select a program to open"
    dlgDialog.Action = 1

    MsgBox "You selected " & dlgDialog.Filename

    Exit Sub

OpenError:

    ' Cancel the error trapping
    On Error GoTo 0

    MsgBox "You hit cancel, so no file was selected!"
    
    Exit Sub

End Sub

Sub cmdPrint_Click ()

    ' When the Print command button is clicked the Printing common dialog will appear.
    ' This is set up simply by firing the number 5 into the Action code of the common
    ' dialog control. In addition the Copies, FromPage and ToPage properties can be set
    ' to supply some default values to the user.

    ' The On error line redirects any trappable errors to the label specified, and in
    ' addition, setting the CancelError property to True causes a trappable error to occur
    ' if the user clicks the cancel button on the common dialog.

    On Error GoTo Print_Error

    dlgDialog.CancelError = True

    dlgDialog.Copies = 5
    dlgDialog.FromPage = 2
    dlgDialog.ToPage = 12

    dlgDialog.Action = 5

    Exit Sub


Print_Error:
    
    ' Cancel the error trapping
    On Error GoTo 0

    MsgBox "You hit cancel - under normal circumstances this would mean that you don't want to print"

    Exit Sub

End Sub

Sub cmdSave_Click ()

    ' Clicking the Save command button causes the Save As dialog to appear.

    On Error GoTo Save_Error

    ' Setting CancelError to True causes a trappable error to occur if the user hits Cancel
    dlgDialog.CancelError = True

    ' The Filter is now set to specify those groups of files that our users will be most
    ' interested in using the dialog to locate
    dlgDialog.Filter = "Executeables (*.exe)|*.exe|Com Files (*.com)|*.com|Batch Files (*.bat)|*.bat"

    ' FilterIndex is set to 1, meaning that the first part of the filter will be used as the
    ' default, in this case *.EXE
    dlgDialog.FilterIndex = 1


    dlgDialog.DialogTitle = "Select a program name to save as"

    ' An action code of 2 triggers the save as dialog to appear.
    dlgDialog.Action = 2

    MsgBox "You selected " & dlgDialog.Filename

    Exit Sub

Save_Error:

    ' Cancel the error trapping
    On Error GoTo 0

    MsgBox "You hit cancel, so no file was selected and in an application no data would be saved"
    
    Exit Sub

End Sub

