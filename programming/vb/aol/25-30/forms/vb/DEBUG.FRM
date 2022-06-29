VERSION 2.00
Begin Form frmMsgboxes 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "The Beginners Guide To Visual Basic."
   ClientHeight    =   6270
   ClientLeft      =   2400
   ClientTop       =   2265
   ClientWidth     =   8445
   Height          =   6675
   Left            =   2340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   8445
   Top             =   1920
   Width           =   8565
   Begin CommandButton cmdCode 
      Caption         =   "Show me the code!"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   4425
      TabIndex        =   19
      Top             =   5520
      Width           =   3645
   End
   Begin CommandButton cmdGo 
      Caption         =   "Do It!"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   495
      Left            =   405
      TabIndex        =   14
      Top             =   5520
      Width           =   3645
   End
   Begin Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "The text and title of the message box"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFF00&
      Height          =   1560
      Left            =   390
      TabIndex        =   13
      Top             =   3825
      Width           =   7695
      Begin TextBox txtMessage 
         Height          =   285
         Left            =   1965
         TabIndex        =   18
         Text            =   "Wow - a working message box. "
         Top             =   1005
         Width           =   4965
      End
      Begin TextBox txtTitle 
         Height          =   285
         Left            =   1965
         TabIndex        =   17
         Text            =   "Wrox waz ere!"
         Top             =   540
         Width           =   4965
      End
      Begin Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "The message"
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   630
         TabIndex        =   16
         Top             =   1035
         Width           =   1350
      End
      Begin Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "The title "
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   645
         TabIndex        =   15
         Top             =   570
         Width           =   2025
      End
   End
   Begin Frame fraIcons 
      BackColor       =   &H00FF0000&
      Caption         =   "Icon in the message box"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFF00&
      Height          =   2685
      Left            =   4440
      TabIndex        =   2
      Top             =   1065
      Width           =   3630
      Begin OptionButton optIcon 
         BackColor       =   &H00FF0000&
         Caption         =   "Information icon - (64)"
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   3
         Left            =   330
         TabIndex        =   6
         Tag             =   "64"
         Top             =   1845
         Width           =   2775
      End
      Begin OptionButton optIcon 
         BackColor       =   &H00FF0000&
         Caption         =   "Exclamation-mark icon - (48)"
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   2
         Left            =   330
         TabIndex        =   5
         Tag             =   "48"
         Top             =   1410
         Width           =   2775
      End
      Begin OptionButton optIcon 
         BackColor       =   &H00FF0000&
         Caption         =   "Question-mark icon - (32)"
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   1
         Left            =   330
         TabIndex        =   4
         Tag             =   "32"
         Top             =   990
         Width           =   2775
      End
      Begin OptionButton optIcon 
         BackColor       =   &H00FF0000&
         Caption         =   "Stop icon - (16)"
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   0
         Left            =   330
         TabIndex        =   3
         Tag             =   "16"
         Top             =   600
         Width           =   2775
      End
   End
   Begin Frame fraButtons 
      BackColor       =   &H00FF0000&
      Caption         =   "Buttons in the message box"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFF00&
      Height          =   2685
      Left            =   390
      TabIndex        =   1
      Top             =   1065
      Width           =   3630
      Begin OptionButton optButtons 
         BackColor       =   &H00FF0000&
         Caption         =   "Retry and Cancel - (5)"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   5
         Left            =   270
         TabIndex        =   12
         Top             =   2265
         Width           =   3180
      End
      Begin OptionButton optButtons 
         BackColor       =   &H00FF0000&
         Caption         =   "Yes and No - (4)"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   4
         Left            =   270
         TabIndex        =   11
         Top             =   1890
         Width           =   3180
      End
      Begin OptionButton optButtons 
         BackColor       =   &H00FF0000&
         Caption         =   "Yes, No and Cancel - (3)"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   3
         Left            =   270
         TabIndex        =   10
         Top             =   1530
         Width           =   3180
      End
      Begin OptionButton optButtons 
         BackColor       =   &H00FF0000&
         Caption         =   "Abort, Retry and Ignore - (2)"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   9
         Top             =   1155
         Width           =   3180
      End
      Begin OptionButton optButtons 
         BackColor       =   &H00FF0000&
         Caption         =   "OK and Cancel - (1)"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   8
         Top             =   780
         Width           =   3180
      End
      Begin OptionButton optButtons 
         BackColor       =   &H00FF0000&
         Caption         =   "OK only - (0)"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   7
         Top             =   405
         Width           =   3180
      End
   End
   Begin Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Message Boxes - Creating them"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   18
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   870
      TabIndex        =   0
      Top             =   210
      Width           =   5880
   End
End
Option Explicit

' The two variables defined in a moment are used to help us build up the message box command.
' The program could also be written without using any variables at all but I would have to
' use loops and conditions we which cover in more detail in chapter 5. This way the program
' code is also extremely small and easy to follow.

Const QUOTE = 34

' We need a variable to hold the value telling the message box which buttons to use
Dim mnButtonCode As Integer

' ... we also need a variable to tell it which icon to show
Dim mnIconCode As Integer

Sub CentreForm ()
    
    ' Centre the form here
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

End Sub

Sub cmdCode_Click ()

    Dim sCommand As String

    ' When the code command button is clicked we need to display a message box to the user
    ' showing them the actual code for the message box they just built.

    sCommand = "MsgBox " & Chr$(34) & txtMessage.Text & Chr$(34) & ", " & mnButtonCode + mnIconCode
    sCommand = sCommand & ", " & Chr$(34) & txtTitle.Text & Chr$(34)

    MsgBox sCommand, 0, "Command Syntax"

End Sub

Sub cmdGo_Click ()

    ' This displays the final message box based on the user's selections on the form
    Debug.Print "------------------------------------"
    Debug.Print "Go Button Click Event"
    Debug.Print "mnButtonCode = " & mnButtonCode
    Debug.Print "mnIconCode = " & mnIconCode
    Debug.Print "------------------------------------"
    
    MsgBox txtMessage.Text, mnButtonCode + mnIconCode, txtTitle.Text

End Sub

Sub Form_Load ()
    

    CentreForm

    ' Set up the default option buttons
    optButtons(0).Value = True
    optIcon(0).Value = True

End Sub

Sub optButtons_Click (Index As Integer)

    ' This event is triggered when one of the option buttons in the buttons frame is clicked.
    ' The messagebox command needs to have a number sent to it which tells it what buttons
    ' to display in the message. Luckily for us this can be a number between 0 and 5, exactly
    ' the same as the Index number passed to this event. (Remember, we are dealing with a
    ' control array, a collection of controls all with the same name but each with its own
    ' unique number called an Index.

    mnIconCode = Index

    ' Dead easy!

End Sub

Sub optIcon_Click (Index As Integer)

    ' This event is triggered when an option button in the icon frame is clicked (but
    ' you already guessed that from the name right!). Ok, in the TAG property of each
    ' option button I have placed a number that the messagebox command needs to decide
    ' which icon to display. All we need to do here is place that number into the variable
    ' defined in the <declarations> section of this form.

    mnIconCode = optIcon(Index).Tag

    ' Easy huh!

End Sub

