VERSION 2.00
Begin Form formEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   2250
   ClientLeft      =   3255
   ClientTop       =   4425
   ClientWidth     =   5430
   ControlBox      =   0   'False
   Height          =   2685
   Left            =   3180
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5430
   Top             =   4065
   Width           =   5580
   Begin CommandButton cmndClose 
      Caption         =   "&Close"
      Height          =   372
      Left            =   4200
      TabIndex        =   1
      Top             =   1800
      Width           =   972
   End
   Begin TextBox textEntry 
      Height          =   1572
      Index           =   0
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4932
   End
End

Sub cmndClose_Click ()
     formEntry.Hide           'Hide the entry box; back to reference form
End Sub

Sub Form_Activate ()
     textEntry(CurrentIndex).ZOrder         'Place the correct Text box in the front of the "stack"
     textEntry(CurrentIndex).Parent.Caption = textEntry(CurrentIndex).Tag  'Label the form
     textEntry(CurrentIndex).TabIndex = 0   'Make the current text box first in the tab order,
     cmndClose.TabIndex = 1                 'And the close button next
End Sub

Sub Form_Load ()
     For x = ADDRESSEE To CARBON                    'We defined the "template" text box
                                                    'as textEntry(RTN_ADDRESS)
          Load textEntry(x)                         'Now create all the other text boxes
          textEntry(x).Visible = True               'And make sure they're visible
     Next x
     textEntry(RTN_ADDRESS).Tag = "Return Address"  'The tags of each text box hold descriptive
     textEntry(ADDRESSEE).Tag = "Addressee"         'text that we use to label the data entry
     textEntry(GREETING).Tag = "Greeting"           'Title Bar
     textEntry(BODY).Tag = "Body Text"
     textEntry(CLOSING).Tag = "Closing"
     textEntry(ENCLOSURE).Tag = "Enclosures"
     textEntry(CARBON).Tag = "Carbon Copies"
End Sub

Sub Form_Paint ()
     textEntry(CurrentIndex).SetFocus             'Ready for user to enter text
End Sub

