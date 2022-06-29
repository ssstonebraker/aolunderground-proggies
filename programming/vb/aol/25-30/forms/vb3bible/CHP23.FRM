VERSION 2.00
Begin Form Clip 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clipboard Project"
   ClientHeight    =   4335
   ClientLeft      =   825
   ClientTop       =   2145
   ClientWidth     =   4110
   Height          =   5055
   Icon            =   CHP23.FRX:0000
   Left            =   750
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4110
   Top             =   1500
   Width           =   4260
   Begin CommandButton cmndClearPicture 
      Caption         =   "Clear Picture"
      Height          =   372
      Left            =   2040
      TabIndex        =   3
      Top             =   3840
      Width           =   1932
   End
   Begin CommandButton cmndClearClipboard 
      Caption         =   "Clear Clipboard"
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   1812
   End
   Begin TextBox Text1 
      Height          =   1092
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2640
      Width           =   3852
   End
   Begin PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2412
      Left            =   120
      Picture         =   CHP23.FRX:0302
      ScaleHeight     =   2385
      ScaleWidth      =   3825
      TabIndex        =   0
      Top             =   120
      Width           =   3852
   End
   Begin Shape shapText 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   7
      Height          =   1092
      Left            =   120
      Top             =   2640
      Visible         =   0   'False
      Width           =   3852
   End
   Begin Shape shapPicture 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   7
      Height          =   2412
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   3852
   End
   Begin Menu menuBarFile 
      Caption         =   "&File"
      Begin Menu menuExit 
         Caption         =   "E&xit"
      End
   End
   Begin Menu menuBarEdit 
      Caption         =   "&Edit"
      Begin Menu menuEdit 
         Caption         =   "Cu&t"
         Index           =   0
      End
      Begin Menu menuEdit 
         Caption         =   "&Copy"
         Index           =   1
      End
      Begin Menu menuEdit 
         Caption         =   "&Paste"
         Index           =   2
      End
   End
End
Const CF_LINK = &HBF00
Const CF_TEXT = 1
Const CF_BITMAP = 2
Const CF_METAFILE = 3
Const CF_DIB = 8
Const CF_PALETTE = 9

Sub cmndClearClipboard_Click ()
    Clipboard.Clear           'clear the clipboard of all data
End Sub

Sub cmndClearPicture_Click ()
     If MsgBox("Clear Picture?", 4) = 6 Then    'ok to clear?
          Picture1.Picture = LoadPicture()      'sure, yabetcha, zap the picture
     End If
End Sub

Sub Form_Load ()
    On Error Resume Next           'if an error occurs, it's because the clipboard isn't loaded yet
    AppActivate "Clipboard"        'try to activate clipboard
    If Err > 0 Then A% = Shell("CLIPBRD.EXE", 1)     'if activation failed, then open the clipboad
End Sub

Sub menuBarEdit_Click ()
     menuEdit(0).Enabled = False                       'Cut
     menuEdit(1).Enabled = False                       'Copy
     menuEdit(2).Enabled = False                       'Paste
     If TypeOf Screen.ActiveControl Is TextBox Then    'on the textbox?
          If Text1.SelText <> "" Then                  'is there text selected?
               menuEdit(0).Enabled = True              'there is text to cut
               menuEdit(1).Enabled = True              'there is text to copy
          End If
          If Clipboard.GetFormat(CF_TEXT) Then         'is there text on clipboard?
               menuEdit(2).Enabled = True              'there is text on clipboard to paste
          End If
     ElseIf TypeOf Screen.ActiveControl Is PictureBox Then  'on the picturebox?
          menuEdit(0).Enabled = True                   'ok to cut picture
          menuEdit(1).Enabled = True                   'ok to copy picture
          If Clipboard.GetFormat(CF_BITMAP) Then       'is there a picture on the clipboard?
               menuEdit(2).Enabled = True              'there is a picture to paste
          End If
     End If
End Sub

Sub menuEdit_Click (Index As Integer)
     Select Case Index
          Case 0 '********************* Cut
               If TypeOf Screen.ActiveControl Is TextBox Then         'on textbox?
                    Clipboard.SetText Text1.SelText                   'copy text to clipboard
                    Text1.SelText = ""                                'clear out selected text
               ElseIf TypeOf Screen.ActiveControl Is PictureBox Then  'on picturebox?
                    Clipboard.SetData Picture1.Picture, CF_BITMAP     'copy picture to clipboard
                    Picture1.Picture = LoadPicture("")                'clear out the picture
               End If
          Case 1 '********************* Copy
               If TypeOf Screen.ActiveControl Is TextBox Then         'on textbox?
                    Clipboard.SetText Text1.SelText                   'copy text to clipboard
               ElseIf TypeOf Screen.ActiveControl Is PictureBox Then  'on picturebox?
                    Clipboard.SetData Picture1.Picture, CF_BITMAP     'copy picture to clipboard
               End If
          Case 2 '********************* Paste
               If TypeOf Screen.ActiveControl Is TextBox Then         'on textbox?
                    If Clipboard.GetFormat(CF_TEXT) = True Then       'clipboard has text to paste?
                         Text1.SelText = Clipboard.GetText(CF_TEXT)   'paste in text at cursor
                    End If
               ElseIf TypeOf Screen.ActiveControl Is PictureBox Then  'on picturebox?
                    If Clipboard.GetFormat(CF_BITMAP) = True Then     'clipboard has picture to paste?
                         Picture1.Picture = Clipboard.GetData(CF_BITMAP)   'paste in the picture
                    End If
               End If
     End Select
End Sub

Sub menuExit_Click ()
     End                 'end the program
End Sub

Sub Picture1_DblClick ()
     cmndClearPicture_Click        'clear the picture
End Sub

Sub Picture1_GotFocus ()
     shapPicture.Visible = True    'highlight the picturebox to show focus
End Sub

Sub Picture1_LostFocus ()
     shapPicture.Visible = False        'unhighlight picturebox to show loss of focus
End Sub

Sub Text1_GotFocus ()
     shapText.Visible = True       'highlight text box to show focus
End Sub

Sub Text1_LostFocus ()
     shapText.Visible = False      'unhighlight text box to show loss of focus
End Sub

