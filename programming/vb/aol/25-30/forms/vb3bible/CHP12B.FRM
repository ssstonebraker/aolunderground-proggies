VERSION 2.00
Begin Form formFontList 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Font List"
   ClientHeight    =   4110
   ClientLeft      =   2100
   ClientTop       =   1545
   ClientWidth     =   4155
   Height          =   4545
   Left            =   2025
   LinkTopic       =   "Form2"
   ScaleHeight     =   205.5
   ScaleMode       =   2  'Point
   ScaleWidth      =   207.75
   Top             =   1185
   Width           =   4305
   Begin CommandButton cmndClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   372
      Left            =   2640
      TabIndex        =   3
      Top             =   3600
      Width           =   1212
   End
   Begin VScrollBar scrlUpDown 
      Height          =   3372
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   252
   End
   Begin PictureBox pictContainer 
      Height          =   3372
      Left            =   120
      ScaleHeight     =   167.25
      ScaleMode       =   2  'Point
      ScaleWidth      =   173.25
      TabIndex        =   0
      Top             =   120
      Width           =   3492
      Begin PictureBox pictList 
         Height          =   1812
         Left            =   360
         ScaleHeight     =   89.25
         ScaleMode       =   2  'Point
         ScaleWidth      =   101.25
         TabIndex        =   1
         Top             =   480
         Width           =   2052
      End
   End
   Begin Label lablTransparent 
      Caption         =   "Transparent"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   2052
   End
End

Sub cmndClose_Click ()
    formFontList.Hide         'Hide this form
    formMain.Show             'show the main form again
End Sub

Sub Form_Load ()
    Screen.MousePointer = 11                      'Hourglass
    pictList.FontSize = 18                        '18 points
    fontsLength = Screen.FontCount * pictList.TextHeight("Test")  'height of entire font list
    pictList.Top = 0                              'make list flush with container's top
    pictList.Left = 0                             'make list flush with container's left
    pictList.Width = pictContainer.ScaleWidth     'make list flush with container's right
    pictList.Height = fontsLength                 'make list as long as entire font list
    scrlUpDown.Min = 0                            'display top of list
    scrlUpDown.Max = pictList.Height - pictContainer.Height   'length of entire list
    scrlUpDown.LargeChange = scrlUpDown.Max / 50  '1/50th of total list
    scrlUpDown.SmallChange = scrlUpDown.Max / 250 '1/250th of total list
    pictList.Cls                                  'clear the list
    pictList.FillStyle = 0                        'solid fill
    pictList.FillColor = QBColor(7)               'light gray
    pictList.AutoRedraw = True                    'repaint automatically when uncovered
    pictList.Line (0, 0)-(pictList.ScaleWidth, pictList.ScaleHeight), QBColor(7), BF
    pictList.CurrentX = 0                         'reset coordinates to left
    pictList.CurrentY = 0                         '  ""      ""         top
    For I = 0 To Screen.FontCount - 1             'iterate through all screen fonts
        pictList.FontName = Screen.Fonts(I)       'change list's font
        pictList.Print Screen.Fonts(I)            'and print the name of the font
    Next I
    Screen.MousePointer = 0                       'set hourglass back to normal
End Sub

Sub pictList_Click ()
     pictList.FontTransparent = Not pictList.FontTransparent 'switch transparency setting
     If pictList.FontTransparent = True Then      'and tell user what the setting is
          lablTransparent = "Transparent"
     Else
          lablTransparent = "Opaque"
     End If
     scrlUpDown.Value = 0                         'go to top of the list
     Form_Load                                    'and redraw the list
End Sub

Sub scrlUpDown_Change ()
     pictList.Top = -scrlUpDown.Value        'scroll the list up or down
End Sub

