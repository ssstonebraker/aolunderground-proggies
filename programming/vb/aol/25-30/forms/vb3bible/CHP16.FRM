VERSION 2.00
Begin Form formMain 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Keys Project"
   ClientHeight    =   4110
   ClientLeft      =   870
   ClientTop       =   1890
   ClientWidth     =   5145
   Height          =   4545
   Icon            =   CHP16.FRX:0000
   KeyPreview      =   -1  'True
   Left            =   795
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   5145
   Top             =   1530
   Width           =   5295
   Begin CommandButton cmndExit 
      Caption         =   "E&xit"
      Height          =   372
      Left            =   3480
      TabIndex        =   13
      Top             =   3480
      Width           =   1452
   End
   Begin CheckBox chekUpper 
      Caption         =   "&Uppercase"
      Height          =   252
      Left            =   3540
      TabIndex        =   10
      Top             =   2040
      Width           =   1272
   End
   Begin CommandButton cmndClear 
      Cancel          =   -1  'True
      Caption         =   "Clear"
      Height          =   432
      Left            =   4200
      TabIndex        =   6
      Top             =   2400
      Width           =   732
   End
   Begin CommandButton cmndSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   432
      Left            =   3480
      TabIndex        =   5
      Top             =   2400
      Width           =   732
   End
   Begin TextBox textBox 
      Height          =   3612
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   2892
   End
   Begin Label lablKeyCode 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   252
      Left            =   3480
      TabIndex        =   12
      Top             =   1680
      Width           =   1452
   End
   Begin Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Key Code:"
      Height          =   252
      Left            =   3480
      TabIndex        =   11
      Top             =   1440
      Width           =   1452
   End
   Begin Label lablAlt 
      AutoSize        =   -1  'True
      Caption         =   "Alt"
      Height          =   192
      Left            =   4560
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   228
   End
   Begin Label lablCtrl 
      AutoSize        =   -1  'True
      Caption         =   "Ctrl"
      Height          =   192
      Left            =   4080
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   288
   End
   Begin Label lablShift 
      AutoSize        =   -1  'True
      Caption         =   "Shift"
      Height          =   192
      Left            =   3540
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   372
   End
   Begin Label lablKeyUp 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Courier"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   252
      Left            =   3480
      TabIndex        =   4
      Top             =   1080
      Width           =   1452
   End
   Begin Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Keys Released:"
      Height          =   252
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   1452
   End
   Begin Label lablKeyDown 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Courier"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Keys Pressed:"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Const SHIFT_MASK = 1     'for KeyDown and KeyUp
Const CTRL_MASK = 2
Const ALT_MASK = 4
Const CHECKED = 1        'to read status of check box

Sub chekUpper_Click ()
     textBox.SetFocus                'Back to the text box for more entry
End Sub

Sub cmndClear_Click ()
     textBox.Text = ""                    'Blank out Text box
     textBox.SetFocus                     'Back to the text box for more entry
End Sub

Sub cmndExit_Click ()
     End                      'End the program
End Sub

Sub cmndSend_Click ()
    Dim Tries As Integer                'Remember how many times tried to start notepad

    On Error GoTo Load_NotePad          'Can't activate? then load it
    AppActivate "NotePad - (untitled)"  'activate notepad
    On Error GoTo 0                     'Successful activation
    SendKeys "%{ }x"                    'Maximize notepad
    SendKeys textBox.Text, True           'send it contents of text box
Exit Sub

Load_NotePad:                           'Need to load notepad
    If Tries > 0 Then                   'Tried this before and failed, so
        MsgBox "Cannot Load NotePad"    '...tell user they're out of luck
        Exit Sub
    Else                                'We haven't tried before, so
        Tries = Shell("NotePad")        '...load notepad
    End If
Resume                                  'Go back and activate notepad

End Sub

Sub Form_KeyDown (keycode As Integer, Shift As Integer)
     If Shift And ALT_MASK Then lablAlt.Visible = True         'Show shift indicator if shift key pressed
     If Shift And CTRL_MASK Then lablCtrl.Visible = True       'Show control indicator if control key pressed
     If Shift And SHIFT_MASK Then lablShift.Visible = True     'Show alt indicator if alt key pressed
End Sub

Sub Form_KeyUp (keycode As Integer, Shift As Integer)
     If Not (Shift And ALT_MASK) Then lablAlt.Visible = False       'Turn off shift indicator if not pressed
     If Not (Shift And CTRL_MASK) Then lablCtrl.Visible = False     'Turn off ctrl indicator if not pressed
     If Not (Shift And SHIFT_MASK) Then lablShift.Visible = False   'Turn off alt indicator if not pressed
End Sub

Sub textBox_KeyDown (keycode As Integer, Shift As Integer)
     Static KeysPressed As Long                                 'remembers how many keys pressed
                                                        
     KeysPressed = KeysPressed + 1                              'count total of keys pressed
     lablKeyDown.Caption = Format$(KeysPressed, "########0")    'and update the display
     lablKeyCode.Caption = Hex$(keycode)
End Sub

Sub textBox_KeyPress (KeyAscii As Integer)
     If chekUpper.Value = CHECKED Then        'If caps lock is checked,
          Char = UCase$(Chr$(KeyAscii))       '...then convert the keystroke to uppercase
          KeyAscii = Asc(Char)                '...and change it before it gets to the text box
     End If
End Sub

Sub textBox_KeyUp (keycode As Integer, Shift As Integer)
     Static KeysReleased                                        'Remembers number of keys released

     KeysReleased = KeysReleased + 1                            'count keys released
     lablKeyUp.Caption = Format$(KeysReleased, "########0")     'and update display
End Sub

