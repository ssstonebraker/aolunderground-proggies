VERSION 2.00
Begin Form frmMain 
   BackColor       =   &H0080FF80&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Colours and Controls collections"
   ClientHeight    =   3855
   ClientLeft      =   1185
   ClientTop       =   2340
   ClientWidth     =   4515
   ControlBox      =   0   'False
   Height          =   4260
   Left            =   1125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4515
   Top             =   1995
   Width           =   4635
   Begin Frame Frame2 
      Caption         =   "Apply colours to which controls"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1230
      Left            =   135
      TabIndex        =   4
      Top             =   1830
      Width           =   4245
      Begin CheckBox chkLabels 
         Caption         =   "&Labels"
         Height          =   240
         Left            =   2490
         TabIndex        =   8
         Top             =   750
         Width           =   1620
      End
      Begin CheckBox chkTextBoxes 
         Caption         =   "&Text Boxes"
         Height          =   240
         Left            =   2490
         TabIndex        =   7
         Top             =   480
         Width           =   1620
      End
      Begin CheckBox chkCheckBoxes 
         Caption         =   "&Check Boxes"
         Height          =   240
         Left            =   435
         TabIndex        =   6
         Top             =   480
         Width           =   1620
      End
      Begin CheckBox chkFrames 
         Caption         =   "&Frames"
         Height          =   240
         Left            =   435
         TabIndex        =   5
         Top             =   750
         Width           =   1695
      End
   End
   Begin Frame Frame1 
      Caption         =   "Sample data form"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1410
      Left            =   150
      TabIndex        =   3
      Top             =   270
      Width           =   4245
      Begin TextBox Text3 
         Height          =   285
         Left            =   1530
         TabIndex        =   14
         Top             =   945
         Width           =   2535
      End
      Begin TextBox Text2 
         Height          =   285
         Left            =   1530
         TabIndex        =   13
         Top             =   630
         Width           =   2535
      End
      Begin TextBox Text1 
         Height          =   285
         Left            =   1530
         TabIndex        =   12
         Top             =   315
         Width           =   2535
      End
      Begin Label Label3 
         Caption         =   "Phone No."
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   225
         Left            =   225
         TabIndex        =   11
         Top             =   1005
         Width           =   1200
      End
      Begin Label Label2 
         Caption         =   "Forename"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   225
         Left            =   225
         TabIndex        =   10
         Top             =   675
         Width           =   1200
      End
      Begin Label Label1 
         Caption         =   "Surname"
         FontBold        =   -1  'True
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   9.75
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   225
         Left            =   225
         TabIndex        =   9
         Top             =   345
         Width           =   1200
      End
   End
   Begin CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   435
      Left            =   3180
      TabIndex        =   2
      Top             =   3210
      Width           =   1200
   End
   Begin CommandButton cmdForecolour 
      Caption         =   "&Foreground"
      Height          =   435
      Left            =   1455
      TabIndex        =   1
      Top             =   3225
      Width           =   1200
   End
   Begin CommandButton cmdBackGrounds 
      Caption         =   "&Background"
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   3240
      Width           =   1200
   End
   Begin CommonDialog dlgColours 
      Left            =   135
      Top             =   5265
   End
End
Option Explicit

Sub cmdBackGrounds_Click ()

    ' Set up a long number variable to hold the colour value selected in a moment
    Dim nColour As Long

    Dim nControlNo As Integer


    On Error GoTo BackColourError

    ' Run up the colours common dialog to select a colour
    dlgColours.CancelError = True
    dlgColours.DialogTitle = "Select a background colour"
    dlgColours.Flags = &H1&
    dlgColours.Action = 3

    ' Store the selected colour in a variable for use in a moment - its quicker to deal with
    ' variables at run time than it is to deal with properties.
    nColour = dlgColours.Color

    ' Start to loop through the controls on a form
    For nControlNo = 0 To frmMain.Controls.Count - 1
    
        ' If text box check box is set and the current control is textbox, then set the colour
        If TypeOf frmMain.Controls(nControlNo) Is TextBox Then
            If chkTextBoxes.Value = 1 Then frmMain.Controls(nControlNo).BackColor = nColour
        End If


        ' If frame check box set and current control is frame, set the colour
        If TypeOf frmMain.Controls(nControlNo) Is Frame Then
            If chkFrames.Value = 1 Then frmMain.Controls(nControlNo).BackColor = nColour
        End If


        ' If labels check box not set and current control is label, continue loop
        If TypeOf frmMain.Controls(nControlNo) Is Label Then
            If chkLabels.Value = 1 Then frmMain.Controls(nControlNo).BackColor = nColour
        End If


        ' If checkbox check box not set, and current control is checkbox, continue loop
        If TypeOf frmMain.Controls(nControlNo) Is CheckBox Then
            If chkCheckBoxes.Value = 1 Then frmMain.Controls(nControlNo).BackColor = nColour
        End If
        
    ' Continue the loop
    Next nControlNo

BackColourError:

    ' This error handler is hit if the user selects the cancel button in the dialog.
    On Error GoTo 0
    Exit Sub

End Sub

Sub cmdForecolour_Click ()

    ' Set up a long number variable to hold the colour value selected in a moment
    Dim nColour As Long

    Dim nControlNo As Integer


    On Error GoTo ForeColourError

    ' Run up the colours common dialog to select a colour
    dlgColours.CancelError = True
    dlgColours.DialogTitle = "Select a foreground colour"
    dlgColours.Flags = &H1&
    dlgColours.Action = 3

    ' Store the selected colour in a variable for use in a moment - its quicker to deal with
    ' variables at run time than it is to deal with properties.
    nColour = dlgColours.Color

    ' Start to loop through the controls on a form
    For nControlNo = 0 To frmMain.Controls.Count - 1
    
        ' If text box check box is set and the current control is textbox, then set the colour
        If TypeOf frmMain.Controls(nControlNo) Is TextBox Then
            If chkTextBoxes.Value = 1 Then frmMain.Controls(nControlNo).ForeColor = nColour
        End If


        ' If frame check box set and current control is frame, set the colour
        If TypeOf frmMain.Controls(nControlNo) Is Frame Then
            If chkFrames.Value = 1 Then frmMain.Controls(nControlNo).ForeColor = nColour
        End If


        ' If labels check box not set and current control is label, continue loop
        If TypeOf frmMain.Controls(nControlNo) Is Label Then
            If chkLabels.Value = 1 Then frmMain.Controls(nControlNo).ForeColor = nColour
        End If


        ' If checkbox check box not set, and current control is checkbox, continue loop
        If TypeOf frmMain.Controls(nControlNo) Is CheckBox Then
            If chkCheckBoxes.Value = 1 Then frmMain.Controls(nControlNo).ForeColor = nColour
        End If
        
    ' Continue the loop
    Next nControlNo

ForeColourError:

    ' This error handler is hit if the user selects the cancel button in the dialog.
    On Error GoTo 0
    Exit Sub

End Sub

Sub cmdQuit_Click ()

    ' Quit the application by unloading the form
    Unload frmMain

End Sub

