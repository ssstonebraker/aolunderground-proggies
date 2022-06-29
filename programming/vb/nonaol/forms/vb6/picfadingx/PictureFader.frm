VERSION 5.00
Begin VB.Form PictureFader 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Picture Fading Example"
   ClientHeight    =   5490
   ClientLeft      =   1590
   ClientTop       =   1545
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   4335
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   1635
   End
   Begin VB.PictureBox PictureHolder 
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   60
      ScaleHeight     =   1995
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   60
      Width           =   2475
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fade The Picture"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   60
      Width           =   1635
   End
   Begin VB.PictureBox HoldNewPic 
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   60
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   4260
      Width           =   1395
   End
   Begin VB.PictureBox endpic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1950
      Left            =   2100
      Picture         =   "PictureFader.frx":0000
      ScaleHeight     =   1950
      ScaleWidth      =   1950
      TabIndex        =   1
      Top             =   2220
      Width           =   1950
   End
   Begin VB.PictureBox startpic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1950
      Left            =   60
      Picture         =   "PictureFader.frx":4512
      ScaleHeight     =   1950
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   2220
      Width           =   1995
   End
End
Attribute VB_Name = "PictureFader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Fade one picture into another picture example
' by PAT or JK
' 11.6.99
' my webpage: http://www.patorjk.com/   (check it out)
' email: patorjk@aol.com

' This examples shows you one way on how to fade one picture into
' another picture. On slow computers the fading may take awhile,
' but for the most part it should fade at an ok speed.

' I picked the radiohead and eminem cd covers as the pictures to fade
' because they are almost the same size (only a little different).
' (and they're also my two favorite cds)

' Anyway, when using this code you should only try to fade images
' that are the same size or almost the same size. Larger images will
' take longer to fade.

Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Sub Command1_Click()
    Dim rows As Integer, cols As Integer
    Dim Color1 As Long, Color2 As Long
    Dim FadeNum As Integer, TheColor As Long
    Dim FadeSize As Integer
    ' just in case something goes wrong
    On Error Resume Next
    
    ' set up all the picture boxes so everything works right
    Me.ScaleMode = 3 ' pixel mode
    startpic.ScaleMode = 3
    endpic.ScaleMode = 3
    ' IMPORTANT - this controls the number of pics it takes to get
    ' to the final one
    FadeSize = 7
    ' make sure all the pictures are the right size
    endpic.Height = startpic.Height
    endpic.Width = startpic.Width
    HoldNewPic.Height = startpic.Height
    HoldNewPic.Width = startpic.Width
    ' make sure we'll be able to get the pixels from the pictures
    ' even when they're not shown
    startpic.AutoRedraw = True
    endpic.AutoRedraw = True
    HoldNewPic.AutoRedraw = True
    ' set our main picture to the starting picture
    PictureHolder.Picture = startpic.Picture
    
    DoEvents
    For FadeNum = 1 To FadeSize
        For cols = 0 To startpic.ScaleWidth
            For rows = 0 To startpic.ScaleHeight
                Color1 = GetPixel(startpic.hdc, rows, cols)
                Color2 = GetPixel(endpic.hdc, rows, cols)
                If Color1 <> Color2 Then
                    ' get the color pixel we want
                    TheColor = GetFadedColor(Color1, Color2, FadeNum, FadeSize)
                Else
                    ' the two colors are the same so we don't
                    ' need to waste time with the fading sub
                    TheColor = Color1
                End If
                ' set the pixel in our holding picture box
                Call SetPixel(HoldNewPic.hdc, rows, cols, TheColor)
            Next
        Next
        ' set the created picture in the holding picture box to the
        ' picture box we see on screen
        Set PictureHolder.Picture = HoldNewPic.Image
    Next
End Sub

Public Function GetFadedColor(c1 As Long, c2 As Long, FN As Integer, FS As Integer) As Long
    Dim i%, red1%, green1%, blue1%, red2%, green2%, blue2%, pat1!, pat2!, pat3!, cx1!, cx2!, cx3!
    
    ' get the red, green, and blue values out of the different
    ' colors
    red1% = (c1 And 255)
    green1% = (c1 \ 256 And 255)
    blue1% = (c1 \ 65536 And 255)
    red2% = (c2 And 255)
    green2% = (c2 \ 256 And 255)
    blue2% = (c2 \ 65536 And 255)
    
    ' get the step of the color changing
    pat1 = (red2% - red1%) / FS
    pat2 = (green2% - green1%) / FS
    pat3 = (blue2% - blue1%) / FS

    ' set the cx variables at the starting colors
    cx1 = red1%
    cx2 = green1%
    cx3 = blue1%

    ' loop till you reach the faze you are at in the fading
    For i% = 1 To FN
        cx1 = cx1 + pat1
        cx2 = cx2 + pat2
        cx3 = cx3 + pat3
    Next
    GetFadedColor = RGB(cx1, cx2, cx3)
End Function

Private Sub Command2_Click()
    ' End the program
    End
End Sub

Private Sub Form_Load()
    Set startpic.Picture = startpic.Image
    PictureHolder.Picture = startpic.Picture
    Me.Height = PictureHolder.Height + (120 * 4)
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub
