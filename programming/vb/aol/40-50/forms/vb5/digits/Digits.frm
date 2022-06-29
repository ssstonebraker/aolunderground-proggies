VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digits Demo"
   ClientHeight    =   2190
   ClientLeft      =   1230
   ClientTop       =   1875
   ClientWidth     =   4395
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2190
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin PicClip.PictureClip clpPunctuation 
      Left            =   3840
      Top             =   1200
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   327680
   End
   Begin PicClip.PictureClip clpDigits 
      Left            =   3840
      Top             =   720
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   327680
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   1680
   End
   Begin VB.Image imgPunctuation 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   0
      Left            =   3120
      Picture         =   "Digits.frx":0000
      Top             =   720
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgClock 
      Appearance      =   0  'Flat
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgPunctuation 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   2
      Left            =   3120
      Picture         =   "Digits.frx":029E
      Top             =   1680
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgPunctuation 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   1
      Left            =   3120
      Picture         =   "Digits.frx":053C
      Top             =   1200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgDigits 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   1
      Left            =   120
      Picture         =   "Digits.frx":07DA
      Top             =   1200
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.Image imgDigits 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   0
      Left            =   120
      Picture         =   "Digits.frx":12E8
      Top             =   720
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.Image imgDigits 
      Appearance      =   0  'Flat
      Height          =   405
      Index           =   2
      Left            =   120
      Picture         =   "Digits.frx":1DF6
      Top             =   1680
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptColor 
         Caption         =   "&Red"
         Index           =   0
      End
      Begin VB.Menu mnuOptColor 
         Caption         =   "&Green"
         Index           =   1
      End
      Begin VB.Menu mnuOptColor 
         Caption         =   "&Blue"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Digits - VB5 PicClip demo program
'Copyright (c) 1994-97 SoftCircuits Programming (R)
'Redistributed by Permission.
'
'This VB5 demo shows how you can create a digital display that appears
'like some sort of electronic LCD/LED display and includes several
'bitmaps that can be used to create such displays. You will need the
'PicClip control that ships with the professional edition of Visual
'Basic in order to run the demo program.
'
'Note: At the time this demo was put together, SoftCircuits was
'putting together an enhanced version of this code called LCD, which
'is implemented as a control that is much easier to use and doesn't
'require the PicClip control. At this time, this enhanced version will
'not be distributed for free. If you find this code useful but it
'doesn't do exactly what you want, you might want to drop by our web
'site to see what else we have available.
'
'This program may be distributed on the condition that it is
'distributed in full and unchanged, and that no fee is charged for
'such distribution with the exception of reasonable shipping and media
'charged. In addition, the code in this program may be incorporated
'into your own programs and the resulting programs may be distributed
'without payment of royalties.
'
'This example program was provided by:
' SoftCircuits Programming
' http://www.softcircuits.com
' P.O. Box 16262
' Irvine, CA 92623
Option Explicit

Private Sub Form_Load()
    Dim i As Integer, xFactor As Integer, yFactor As Integer

    'Init PicClip controls
    clpDigits.Cols = 11
    clpPunctuation.Cols = 8
    'Load PicClip bitmaps so can know cell sizes
    Call SetColor(0, False)
    'Load image controls to hold digits
    imgClock(0) = clpDigits.GraphicCell(0)
    For i = 1 To 7
        Load imgClock(i)
        'Digit cells are a different size than colon (punctuation) cells
        If i = 2 Or i = 5 Then
            imgClock(i) = clpPunctuation.GraphicCell(0)
        Else
            imgClock(i) = clpDigits.GraphicCell(0)
        End If
        imgClock(i).Left = imgClock(i - 1).Left + imgClock(i - 1).Width
        imgClock(i).Visible = True
    Next i
    'Size window to fit time display
    xFactor = Width - ScaleWidth: yFactor = Height - ScaleHeight
    Move Left, Top, imgClock(7).Left + imgClock(7).Width + 120 + xFactor, imgClock(0).Height + 240 + yFactor
    'Show initial time display
    Call ShowCurrTime
End Sub

Private Sub Form_Resize()
    'Since we display the time in the caption when the form is
    'minimized, restore caption if we are no longer minimized
    If WindowState <> 1 Then
        Caption = "Digits Demo"
    Else
        'If window has just been minimized, update time
        Call ShowCurrTime
    End If
End Sub

Private Sub mnuFileExit_Click()
    'Unload form to terminate program
    Unload Me
End Sub

Private Sub mnuOptColor_Click(Index As Integer)
    'Update display color
    Call SetColor(Index, True)
End Sub

Private Sub SetColor(clr As Integer, updateTime As Integer)
    Static currColor As Integer
    Dim i As Integer

    'Set new color index
    currColor = clr
    'Load PicClip controls with bitmap for selected color
    clpDigits = imgDigits(currColor)
    clpPunctuation = imgPunctuation(currColor)
    'Check/uncheck menu items to indicate current color
    For i = 0 To 2
        mnuOptColor(i).Checked = (i = currColor)
    Next i
    'Update time display if requested
    If updateTime Then Call ShowCurrTime
End Sub

Private Sub ShowCurrTime()
    Static showColon As Integer
    Dim i As Integer, buff As String, aChar As String

    'If window is minimized, show time in caption
    If WindowState = 1 Then
        Caption = Format$(Now, "h:mm:ss am/pm")
    Else
        'Get current time in buff
        buff = Format$(Now, "hh:mm:ss am/pm")
        'Hide first character if it is "0"
        aChar = Mid$(buff, 1, 1)
        If aChar = "0" Then
            imgClock(0) = clpDigits.GraphicCell(10)
        Else
            imgClock(0) = clpDigits.GraphicCell(Asc(aChar) - Asc("0"))
        End If
        'Display remaining digits
        For i = 2 To 8
            aChar = Mid$(buff, i, 1)
            If aChar = ":" Then
                If showColon Then
                    imgClock(i - 1) = clpPunctuation.GraphicCell(2)
                Else
                    imgClock(i - 1) = clpPunctuation.GraphicCell(3)
                End If
            Else
                imgClock(i - 1) = clpDigits.GraphicCell(Asc(aChar) - Asc("0"))
            End If
        Next i
    End If
    'Toggle display of colon
    showColon = Not showColon
End Sub

Private Sub Timer1_Timer()
    'Update time display
    Call ShowCurrTime
End Sub

