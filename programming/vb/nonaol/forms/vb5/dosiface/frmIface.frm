VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIface 
   BorderStyle     =   0  'None
   Caption         =   "dos's iface example"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8445
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MouseIcon       =   "frmIface.frx":0000
   ScaleHeight     =   5655
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdgIface 
      Left            =   5000
      Top             =   500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "dos's iface example"
      Filter          =   "bitmap and dat files (*.bmp, *.dat)|*.bmp|*.dat"
   End
   Begin VB.PictureBox picExit 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   3280
      ScaleHeight     =   345
      ScaleWidth      =   1050
      TabIndex        =   11
      Top             =   1680
      Width           =   1050
   End
   Begin VB.PictureBox picLoadSkin 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   2160
      ScaleHeight     =   345
      ScaleWidth      =   1050
      TabIndex        =   10
      Top             =   1680
      Width           =   1050
   End
   Begin VB.PictureBox picMsg3 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   2160
      ScaleHeight     =   345
      ScaleWidth      =   2175
      TabIndex        =   9
      Top             =   1240
      Width           =   2175
   End
   Begin VB.PictureBox picMsg2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   2160
      ScaleHeight     =   345
      ScaleWidth      =   2175
      TabIndex        =   8
      Top             =   840
      Width           =   2175
   End
   Begin VB.PictureBox picMsg1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   2160
      ScaleHeight     =   345
      ScaleWidth      =   2175
      TabIndex        =   7
      Top             =   440
      Width           =   2175
   End
   Begin VB.PictureBox picControlExit 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   60
      Width           =   195
   End
   Begin VB.PictureBox picControlMin 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   60
      Width           =   195
   End
   Begin VB.PictureBox picInfo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   330
      ScaleHeight     =   1185
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   610
      Width           =   1515
   End
   Begin VB.PictureBox picWindow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   140
      ScaleHeight     =   1575
      ScaleWidth      =   1935
      TabIndex        =   3
      Top             =   430
      Width           =   1935
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   4500
      TabIndex        =   2
      Top             =   300
      Width           =   4500
   End
   Begin VB.PictureBox picTitleBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   4500
      TabIndex        =   1
      Top             =   0
      Width           =   4500
   End
   Begin VB.PictureBox picSourceImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3360
      Left            =   -360
      Picture         =   "frmIface.frx":030A
      ScaleHeight     =   3360
      ScaleWidth      =   8775
      TabIndex        =   0
      Top             =   2280
      Width           =   8775
   End
End
Attribute VB_Name = "frmIface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'*         i-face example by dos                          *
'*         1/25/99                                        *
'*         email: xdosx@hotmail.com                       *
'*         aim:   xdosx                                   *
'**********************************************************
'* the major focus of this example is the use of the      *
'* bitblt api. for those of you who aren't familiar with  *
'* bitblt, it basically allows you to transfer an image or*
'* a section of an image to another image. bitblt has the *
'* following arguments....                                *
'*                                                        *
'* ByVal hDestDC As Long,                                 *
'*    hDestCD is the handle of the destination of the     *
'*    image. for example "picture1.hdc".                  *
'*                                                        *
'* ByVal X As Long,                                       *
'*    X is the x-coordinate of destination rectangle's    *
'*    upper-left corner.                                  *
'*                                                        *
'* ByVal Y As Long,                                       *
'*    Y is the y-coordinate of destination rectangle's    *
'*    upper-left corner.                                  *
'*                                                        *
'* ByVal nWidth As Long,                                  *
'*    nWidth is the width of destination rectangle.       *
'*                                                        *
'* ByVal nHeight As Long,                                 *
'*    nHeight is the height of the destination rectangle. *
'*                                                        *
'* ByVal hSrcDC As Long,                                  *
'*    hSrcDC is the handle of the source image. for       *
'*    example "picture2.hdc".                             *
'*                                                        *
'* ByVal xSrc As Long,                                    *
'*    xSrc is the x-coordinate of source rectangle's      *
'*    upper-left corner.                                  *
'*                                                        *
'* ByVal ySrc As Long,                                    *
'*    ySrc is the y-coordinate of the source rectangle's  *
'*    upper-left corner.                                  *
'*                                                        *
'* ByVal dwRop As Long                                    *
'*    dwRop is the raster operation code. we use SRCCOPY  *
'*    since it copies the source rectangle directly to the*
'*    destination rectangle.                              *
'**********************************************************
'* the next thing you will see is the use of capture api  *
'* and the cursor position. these should be pretty self-  *
'* explanitory in the code. they are used to give the     *
'* java-script like mouse-overs.                          *
'*                                                        *
'* also, you may notice that with some of the buttons, i  *
'* check to make sure the cursor is still over the button *
'* in the mouseup event. i do this because just like with *
'* regular windows, when you press a button, should you   *
'* move it away from the button before releasing the mouse*
'* button, it will not "click" the button.                *
'**********************************************************
'  tips....                                               *
'*   1. make sure to set your picture box's autoredraw    *
'*      property to true. this will not work if you don't *
'*   2. in your vb menu, go to format/lock controls. the  *
'*      picture boxes are hard enough to line up as it is *
'*      and you won't want to go through that again just  *
'*      because you accidently knocked them out of their  *
'*      positions.                                        *
'*   3. vary your picture box's background colors. this   *
'*      will make them easier to keep track of.           *
'**********************************************************
'* on the images, they can be changed to a ".dat" file or *
'* whatever. i have seen several people ask about this.   *
'* the file extension isn't important as long as it is    *
'* still a valid bitmap (granted, vb6 can load other image*
'* types that aren't available in earlier versions.       *
'*                                                        *
'* also, don't critisize my artwork. i am not an artist by*
'* any means. this is just what i could do with photoshop *
'* in a few hours. i don't actually use these images, they*
'* were made for this example only. hell, i don't even    *
'* like this ieet0 iface stuff. =)                        *
'*                                                        *
'*                                    dos                 *
'**********************************************************
'*                                                        *
Dim bMoveFrom As Boolean, LastPoint As POINTAPI

Private Sub LoadIface()
    Call BitBlt(picTitleBar.hDC, 0, 0, 300, 20, picSourceImage.hDC, 0, 0, SRCCOPY)
    picTitleBar.Refresh
    Call BitBlt(picMain.hDC, 0, 0, 300, 125, picSourceImage.hDC, 0, 20, SRCCOPY)
    picMain.Refresh
    Call BitBlt(picWindow.hDC, 0, 0, 129, 105, picSourceImage.hDC, 445, 0, SRCCOPY)
    picWindow.Refresh
    Call BitBlt(picInfo.hDC, 0, 0, 101, 79, picSourceImage.hDC, 0, 145, SRCCOPY)
    picInfo.Refresh
    Call BitBlt(picControlMin.hDC, 0, 0, 13, 13, picSourceImage.hDC, 407, 145, SRCCOPY)
    picControlMin.Refresh
    Call BitBlt(picControlExit.hDC, 0, 0, 13, 13, picSourceImage.hDC, 422, 145, SRCCOPY)
    picControlExit.Refresh
    Call BitBlt(picMsg1.hDC, 0, 0, 145, 23, picSourceImage.hDC, 300, 0, SRCCOPY)
    picMsg1.Refresh
    Call BitBlt(picMsg2.hDC, 0, 0, 145, 23, picSourceImage.hDC, 300, 23, SRCCOPY)
    picMsg2.Refresh
    Call BitBlt(picMsg3.hDC, 0, 0, 145, 23, picSourceImage.hDC, 300, 46, SRCCOPY)
    picMsg3.Refresh
    Call BitBlt(picLoadSkin.hDC, 0, 0, 70, 23, picSourceImage.hDC, 515, 105, SRCCOPY)
    picLoadSkin.Refresh
    Call BitBlt(picExit.hDC, 0, 0, 70, 23, picSourceImage.hDC, 445, 105, SRCCOPY)
    picExit.Refresh
End Sub

Private Sub Form_Load()
    Me.Width = 4500
    Me.Height = 2175
    iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
    Call LoadIface
End Sub

Private Sub picControlExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(picControlExit.hDC, 0, 0, 13, 13, picSourceImage.hDC, 422, 159, SRCCOPY)
    picControlExit.Refresh
End Sub

Private Sub picControlExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(picControlExit.hDC, 0, 0, 13, 13, picSourceImage.hDC, 422, 145, SRCCOPY)
    picControlExit.Refresh
    If X > 0 And X < picControlExit.Width And Y > 0 And Y < picControlExit.Height Then
        Unload Me
    End If
End Sub

Private Sub picControlMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(picControlMin.hDC, 0, 0, 13, 13, picSourceImage.hDC, 407, 159, SRCCOPY)
    picControlMin.Refresh
End Sub

Private Sub picControlMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(picControlMin.hDC, 0, 0, 13, 13, picSourceImage.hDC, 407, 145, SRCCOPY)
    picControlMin.Refresh
    If X > 0 And X < picControlMin.Width And Y > 0 And Y < picControlMin.Height Then
        Me.WindowState = 1
    End If
End Sub

Private Sub picExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(picExit.hDC, 0, 0, 70, 23, picSourceImage.hDC, 445, 105, SRCCOPY)
    picExit.Refresh
    If X > 0 And X < picExit.Width And Y > 0 And Y < picExit.Height Then
        Unload Me
    End If
End Sub

Private Sub picExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(picExit.hDC, 0, 0, 70, 23, picSourceImage.hDC, 445, 151, SRCCOPY)
    picExit.Refresh
End Sub

Private Sub picExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long
    If GetCapture() <> picExit.hwnd Then
        Ret = SetCapture(picExit.hwnd)
        Call BitBlt(picExit.hDC, 0, 0, 70, 23, picSourceImage.hDC, 445, 128, SRCCOPY)
        picExit.Refresh
        Me.MousePointer = 99
    End If
    If X > 0 And X < picExit.Width And Y > 0 And Y < picExit.Height Then
        CurrentX = X
        CurrentY = Y
    Else
        If GetCapture() = picExit.hwnd Then
            Ret = ReleaseCapture()
            Call BitBlt(picExit.hDC, 0, 0, 70, 23, picSourceImage.hDC, 445, 105, SRCCOPY)
            picExit.Refresh
            Me.MousePointer = 0
        End If
    End If
End Sub

Private Sub picLoadSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call BitBlt(picLoadSkin.hDC, 0, 0, 70, 23, picSourceImage.hDC, 515, 151, SRCCOPY)
    picLoadSkin.Refresh
End Sub

Private Sub picLoadSkin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long
    If GetCapture() <> picLoadSkin.hwnd Then
        Ret = SetCapture(picLoadSkin.hwnd)
        Call BitBlt(picLoadSkin.hDC, 0, 0, 70, 23, picSourceImage.hDC, 515, 128, SRCCOPY)
        picLoadSkin.Refresh
        Me.MousePointer = 99
    End If
    If X > 0 And X < picLoadSkin.Width And Y > 0 And Y < picLoadSkin.Height Then
        CurrentX = X
        CurrentY = Y
    Else
        If GetCapture() = picLoadSkin.hwnd Then
            Ret = ReleaseCapture()
            Call BitBlt(picLoadSkin.hDC, 0, 0, 70, 23, picSourceImage.hDC, 515, 105, SRCCOPY)
            picLoadSkin.Refresh
            Me.MousePointer = 0
        End If
    End If
End Sub

Private Sub picLoadSkin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Error_Event:
    Call BitBlt(picLoadSkin.hDC, 0, 0, 70, 23, picSourceImage.hDC, 515, 105, SRCCOPY)
    picLoadSkin.Refresh
    If X > 0 And X < picLoadSkin.Width And Y > 0 And Y < picLoadSkin.Height Then
        cdgIface.CancelError = True
        cdgIface.ShowOpen
        picSourceImage.Picture = LoadPicture(cdgIface.FileName)
        Call LoadIface
    End If
Error_Event:
    Exit Sub
End Sub

Private Sub picMsg1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long
    If GetCapture() <> picMsg1.hwnd Then
        Ret = SetCapture(picMsg1.hwnd)
        Call BitBlt(picMsg1.hDC, 0, 0, 145, 23, picSourceImage.hDC, 300, 69, SRCCOPY)
        picMsg1.Refresh
        Call BitBlt(picInfo.hDC, 0, 0, 101, 79, picSourceImage.hDC, 101, 145, SRCCOPY)
        picInfo.Refresh
        Me.MousePointer = 99
    End If
    If X > 0 And X < picMsg1.Width And Y > 0 And Y < picMsg1.Height Then
        CurrentX = X
        CurrentY = Y
    Else
        If GetCapture() = picMsg1.hwnd Then
            Ret = ReleaseCapture()
            Call BitBlt(picMsg1.hDC, 0, 0, 145, 23, picSourceImage.hDC, 300, 0, SRCCOPY)
            picMsg1.Refresh
            Call BitBlt(picInfo.hDC, 0, 0, 101, 79, picSourceImage.hDC, 0, 145, SRCCOPY)
            picInfo.Refresh
            Me.MousePointer = 0
        End If
    End If
End Sub

Private Sub picMsg2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long
    If GetCapture() <> picMsg2.hwnd Then
        Ret = SetCapture(picMsg2.hwnd)
        Call BitBlt(picMsg2.hDC, 0, 0, 145, 23, picSourceImage.hDC, 300, 92, SRCCOPY)
        picMsg2.Refresh
        Call BitBlt(picInfo.hDC, 0, 0, 101, 79, picSourceImage.hDC, 202, 145, SRCCOPY)
        picInfo.Refresh
        Me.MousePointer = 99
    End If
    If X > 0 And X < picMsg2.Width And Y > 0 And Y < picMsg2.Height Then
        CurrentX = X
        CurrentY = Y
    Else
        If GetCapture() = picMsg2.hwnd Then
            Ret = ReleaseCapture()
            Call BitBlt(picMsg2.hDC, 0, 0, 145, 23, picSourceImage.hDC, 300, 23, SRCCOPY)
            picMsg2.Refresh
            Call BitBlt(picInfo.hDC, 0, 0, 101, 79, picSourceImage.hDC, 0, 145, SRCCOPY)
            picInfo.Refresh
            Me.MousePointer = 0
        End If
    End If
End Sub

Private Sub picMsg3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret As Long
    If GetCapture() <> picMsg3.hwnd Then
        Ret = SetCapture(picMsg3.hwnd)
        Call BitBlt(picMsg3.hDC, 0, 0, 145, 23, picSourceImage.hDC, 300, 115, SRCCOPY)
        picMsg3.Refresh
        Call BitBlt(picInfo.hDC, 0, 0, 101, 79, picSourceImage.hDC, 303, 145, SRCCOPY)
        picInfo.Refresh
        Me.MousePointer = 99
    End If
    If X > 0 And X < picMsg3.Width And Y > 0 And Y < picMsg3.Height Then
        CurrentX = X
        CurrentY = Y
    Else
        If GetCapture() = picMsg3.hwnd Then
            Ret = ReleaseCapture()
            Call BitBlt(picMsg3.hDC, 0, 0, 145, 23, picSourceImage.hDC, 300, 46, SRCCOPY)
            picMsg3.Refresh
            Call BitBlt(picInfo.hDC, 0, 0, 101, 79, picSourceImage.hDC, 0, 145, SRCCOPY)
            picInfo.Refresh
            Me.MousePointer = 0
        End If
    End If
End Sub

Private Sub picTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    bMoveFrom = True
End Sub

Private Sub picTitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iDX As Long, iDY As Long
    Dim POINT As POINTAPI
    If Not bMoveFrom Then Exit Sub
    GetCursorPos POINT
    iDX& = (POINT.X - LastPoint.X) * iTPPX&
    iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    Me.Move Me.Left + iDX&, Me.Top + iDY&
End Sub

Private Sub picTitleBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMoveFrom = False
End Sub
