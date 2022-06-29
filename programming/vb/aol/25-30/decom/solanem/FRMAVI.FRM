VERSION 2.00
Begin Form frmAVIShow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Welcome to MMer..."
   ClientHeight    =   2385
   ClientLeft      =   3450
   ClientTop       =   3390
   ClientWidth     =   3180
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FontBold        =   0   'False
   FontItalic      =   0   'False
   FontName        =   "MS Sans Serif"
   FontSize        =   8.25
   FontStrikethru  =   -1  'True
   FontUnderline   =   0   'False
   Height          =   2790
   Left            =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3180
   Top             =   3045
   Width           =   3300
   Begin Timer Timer1 
      Interval        =   3400
      Left            =   990
      Top             =   945
   End
   Begin CheckBox chkTargeting 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Targeting On"
      Height          =   255
      Left            =   2220
      TabIndex        =   9
      Top             =   5220
      Visible         =   0   'False
      Width           =   1515
   End
   Begin HScrollBar hscrSpeed 
      Height          =   255
      LargeChange     =   5
      Left            =   1260
      Max             =   1
      Min             =   200
      TabIndex        =   5
      Top             =   4800
      Value           =   20
      Visible         =   0   'False
      Width           =   3435
   End
   Begin PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   675
      Left            =   180
      Picture         =   FRMAVI.FRX:0000
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   4260
      Visible         =   0   'False
      Width           =   675
   End
   Begin PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   675
      Left            =   180
      Picture         =   FRMAVI.FRX:04B2
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   3540
      Visible         =   0   'False
      Width           =   675
   End
   Begin CommandButton btnStart 
      Caption         =   "&Start"
      Height          =   435
      Left            =   3300
      TabIndex        =   2
      Top             =   3270
      Width           =   1335
   End
   Begin CommonDialog CMDialog1 
      Left            =   4860
      Top             =   3660
   End
   Begin MMControl MMControl1 
      Height          =   375
      Left            =   1140
      TabIndex        =   1
      Top             =   3540
      UpdateInterval  =   0
      Visible         =   0   'False
      Width           =   3540
   End
   Begin PictureBox picAVIScreen 
      BackColor       =   &H00000000&
      Height          =   2775
      Left            =   90
      ScaleHeight     =   183
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   206
      TabIndex        =   0
      Top             =   120
      Width           =   3120
   End
   Begin Label lblAveFrames 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Top             =   3480
      Width           =   3375
   End
   Begin Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Faster"
      Height          =   195
      Left            =   3840
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
   End
   Begin Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Slower"
      Height          =   195
      Left            =   1260
      TabIndex        =   6
      Top             =   4560
      Width           =   795
   End
   Begin Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Option Explicit
'--------------------------------------------------
' AVI_ANIM.FRM
'--------------------------------------------------

' Constants used in Windows API Call
Const SRCCOPY = &HCC0020
Const SRCAND = &H8800C6
Const SRCPAINT = &HEE0086

' Windows API call used to render bitmaps
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal DestX As Integer, ByVal DestY As Integer, ByVal DestWidth As Integer, ByVal DestHeight As Integer, ByVal hSrcDC As Integer, ByVal SrcX As Integer, ByVal SrcY As Integer, ByVal RasterOp As Long) As Integer

' The x-y coordinates at which to draw the "targeting" sprite.
Dim DrawX As Integer
Dim DrawY As Integer

' Store the last frame of the AVI file.
Dim LastFrameNum As Integer

Sub btnstart_click ()
'--------------------------------------------------
' When the Start/Stop button is pushed, we run the
' animation sequence.  Rather than use the MCI Play
' command, we use the MCI Step command to move
' frame-by-frame through the video clip.
'--------------------------------------------------
Static Running As Integer
Static FrameNum As Integer
Static StartTime As Single
Static TotalFrames As Long
Dim AveFrames As Single
Dim rc As Integer

    ' Exit if they haven't chosen an AVI file.
    If MMControl1.FileName = "" Then Exit Sub

    ' Toggle the animation on and off
    Running = Not Running

    ' Set the button caption to reflect what state
    ' we're in.
    If Running Then
	btnStart.Caption = "&Stop"
	StartTime = Timer
	TotalFrames = 0
    Else
	btnStart.Caption = "&Start"
	AveFrames = TotalFrames / (Timer - StartTime)
	lblAveFrames = "Average Frames: " & Format$(AveFrames, "#0.0") & " per second"
    End If

    ' Display the video clip, frame by frame.
    Do While Running
	' When we get to the end of the animation, we go back to
	' the beginning and start over, creating a continuous loop.
	If FrameNum >= (LastFrameNum - 1) Then MMControl1.Command = "Play"

	' Keep Track of the current frame number so we know
	'If LastFrameNum = 0 Then LastFrameNum = 1
	' when we reach the end.
	'FrameNum = (FrameNum + 1) Mod LastFrameNum
	'MMControl1.Command = "Step"

	' Only display the targeting sprite if chkTargeting
	' is checked.
	If chkTargeting = 1 Then
	    ' Draw the Mask onto the background
	    rc = BitBlt(picAVIScreen.hDC, DrawX, DrawY, picMask.ScaleWidth, picMask.ScaleHeight, picMask.hDC, 0, 0, SRCAND)
	    ' Draw the Sprite onto the backgorund
	    rc = BitBlt(picAVIScreen.hDC, DrawX, DrawY, picSprite.ScaleWidth, picSprite.ScaleHeight, picSprite.hDC, 0, 0, SRCPAINT)
	End If

	' Add a delay specified by the value of the Speed scroll bar.
	Delay hscrSpeed
	TotalFrames = TotalFrames + 1
    Loop

End Sub

Sub Delay (ByVal ms As Single)
'---------------------------------------------------------
' Delay the program for a specified number of milliseconds.
'---------------------------------------------------------
Dim StartTimer As Single

    StartTimer = Timer
    Do
	DoEvents
    Loop While (Timer - StartTimer) < (ms / 1000)
    
End Sub

Sub Form_Load ()
'--------------------------------------------------
' Arrange some of the controls on the form.
'--------------------------------------------------
Dim OK As Integer
Dim FileExt As String
Dim afilename As String
       Show
    OK = False
'    GetNewFileName AFileName
 'AFileName = "D:\Stuff\Ultra\Gate2A.AVI"
 afilename = app.Path & "\mmer.AVI"
 ' Extract the file's extension
    If Len(afilename) > 4 Then FileExt = Right$(afilename, 4)

    ' Based on the extension, load the media file
    Select Case FileExt
	Case ".AVI"
	    'ResetScreenSize
	    OK = SetupAVI(afilename, MMControl1, picAVIScreen)
	    'Me.Caption = "AVI Animation - [" & LCase$(MMControl1.FileName) & "]"
    End Select
    
    ' Set the AVI viewing screen to its maximum size.
    'ResetScreenSize
 btnstart_click
    ' Position the Average Frames label.
    'lblAveFrames.Move MMControl1.Left, MMControl1.Top, MMControl1.Width
End Sub

Sub Form_Unload (Cancel As Integer)
'--------------------------------------------------
' End the program when the form is unloaded.
'--------------------------------------------------

    ' Make sure any AVI file is closed before we
    ' exit.
    MMControl1.Command = "CLOSE"
timer1.Enabled = False
'    End
End Sub

Sub GetNewFileName (AFile As String)
'--------------------------------------------------
' Use the Common File Dialog to retrieve a file name.
'--------------------------------------------------
Dim Canceled As Integer

    Canceled = False
    On Error GoTo GetNewFileName_Error

    ' Set up the Common File Dialog
    CMDialog1.InitDir = app.Path
    CMDialog1.Filter = "AVI Video Files|*.avi"
    CMDialog1.DialogTitle = "Select AVI Video File"
    CMDialog1.CancelError = True

    ' Open the Common File Dialog
    CMDialog1.Action = 1

    If Not Canceled Then
	AFile = CMDialog1.Filename
    Else
	AFile = ""
    End If
Exit Sub
GetNewFileName_Error:
    Canceled = True
    Resume Next
End Sub

Sub mnuFileExit_Click ()
'--------------------------------------------------
' Exit the program by unloading the form.
'--------------------------------------------------

    Unload Me
End Sub

Sub mnuFileOpen_Click ()
'--------------------------------------------------
' Retrieve a media file name and attempt to load
' it into the MCI control.
'--------------------------------------------------
Dim OK As Integer
Dim FileExt As String
Dim afilename As String

    OK = False
    GetNewFileName afilename

    ' Extract the file's extension
    If Len(afilename) > 4 Then FileExt = Right$(afilename, 4)

    ' Based on the extension, load the media file
    Select Case FileExt
	Case ".AVI"
	    'ResetScreenSize
	    OK = SetupAVI(afilename, MMControl1, picAVIScreen)
	    Me.Caption = "AVI Animation - [" & LCase$(MMControl1.FileName) & "]"
    End Select

End Sub

Sub ResetScreenSize ()
'--------------------------------------------------
' Adjust the AVI viewing screen (a PictureBox
' control) to its maximum allowable size.
'--------------------------------------------------
Dim margin As Integer

    margin = 10 * Screen.TwipsPerPixelY
    picAVIScreen.Move margin, margin, Me.ScaleWidth - (margin * 2), MMControl1.Top - (margin * 2)

End Sub

Sub SetAVIDimensions ()
'--------------------------------------------------
' Readjust the AVI viewing screen PictureBox to the
' actual dimensions of the AVI file.
'--------------------------------------------------
Dim x As Integer
Dim y As Integer
Dim AColor As Long

    ' Look for the first non-white pixel in the y direction.
    For y = (picAVIScreen.ScaleHeight) - 2 To 0 Step -1
	AColor = picAVIScreen.Point(1, y)
	If AColor <> RGB(255, 255, 255) Then Exit For
    Next

    ' Look for the first non-white pixel in the x direction.
    For x = (picAVIScreen.ScaleWidth) - 2 To 0 Step -1
	AColor = picAVIScreen.Point(x, 1)
	If AColor <> RGB(255, 255, 255) Then Exit For
    Next

    ' Readjust the picture box.
    picAVIScreen.Width = x * Screen.TwipsPerPixelX
    picAVIScreen.Height = y * Screen.TwipsPerPixelY

    ' Center the picture box.
    picAVIScreen.Left = (Me.ScaleWidth - picAVIScreen.Width) / 2
    picAVIScreen.Top = (MMControl1.Top - picAVIScreen.Height) / 2
    
    ' Calclaute the new center position of the targeting sprite.
    DrawX = (picAVIScreen.ScaleWidth - picMask.ScaleWidth) / 2
    DrawY = (picAVIScreen.ScaleHeight - picMask.ScaleHeight) / 2
    
    ' Save the last frame number of this AVI video clip.
    LastFrameNum = MMControl1.Length
End Sub

Function SetupAVI (ByVal afilename As String, MMControl1 As MMControl, picAVIScreen As PictureBox) As Integer
'--------------------------------------------------
' Attempt to load an Audio-Video Interleave (AVI)
' file into a picture box (picAVIScreen) for
' display.
'--------------------------------------------------

    On Error GoTo SetupAVI_Error
    SetupAVI = False

    ' Load the file into the MCI control.
    MMControl1.Command = "Close"
    If MMControl1.FileName <> afilename Then
	MMControl1.Notify = False
	MMControl1.Wait = True
	MMControl1.Shareable = False
	MMControl1.DeviceType = "AVIVideo"
	MMControl1.FileName = afilename
    End If

    ' Display the picture control.
    picAVIScreen.Visible = True
    
    ' This causes the file to be displayed in the picture
    ' box, instead of the default popup window.
    MMControl1.hWndDisplay = picAVIScreen.hWnd

    ' Open the MCI AVI device.
    MMControl1.Command = "Open"
    MMControl1.Command = "Step"

    ' Position the picAVIScreen control for the new AVI file.
    'SetAVIDimensions
    MMControl1.Refresh

    SetupAVI = True
Exit Function
SetupAVI_Error:
    Exit Function
End Function

Sub Timer1_Timer ()
MMControl1.Command = "STOP"
frmbuster.Show
Me.Hide
End Sub

