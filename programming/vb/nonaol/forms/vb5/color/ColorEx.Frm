VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1935
   ClientLeft      =   2130
   ClientTop       =   2235
   ClientWidth     =   4410
   Height          =   2340
   Icon            =   "ColorEx.frx":0000
   Left            =   2070
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   4410
   Top             =   1890
   Width           =   4530
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Text            =   "Click here to change the color of the background"
      Top             =   2880
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Text            =   "Click here to change the color of the text"
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Background"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Text"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "ColorEx.frx":030A
      Top             =   480
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   3240
   End
   Begin VB.Label Label1 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin DwsbcLib.SubClass SubClass1 
      Left            =   1080
      Top             =   5880
      _Version        =   262144
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      CtlParam        =   ""
      Persist         =   0
      RegMessage1     =   ""
      RegMessage2     =   ""
      RegMessage3     =   ""
      RegMessage4     =   ""
      RegMessage5     =   ""
      Type            =   1
      Messages        =   "ColorEx.frx":037D
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' *** Title bar Manipulation Example

' This program demonstrates how to modify the caption bar of a
' window.  It creates a gradient of the current caption color
' and left justifies the text like the current Microsoft Office
' applicaitons do.  It correctly handles the transition between
' being in focus and not in focus, and it correctly replaces the
' caption when the window is minimized.
'
' It would not be hard to remove the part of the code that does
' the drawing and replace it with whatever drawing functions you
' want.  However, here are a few pointers to help you create a
' useful and attractive window:
' 1) Be sure the Caption property of the window is cleared when
'    the window is not minimized.  On systems with real-time
'    window resizing (like Windows NT and Windows 95 with the
'    Plus pack) the caption will be redrawn frequently, causing
'    the text underneath to blink visibly.
' 2) Keep the color of the caption similar to the one set by the
'    user.  It helps prevent noticable blinking, and keeps the
'    interface unified.  Do not create a garish window that
'    distracts the user.
' 3) Make sure a focused window is differentiated from an
'    unfocused one.
' 4) Keep the drawing time down.  This also help prevent blinking.
' 5) Be prepared for changing colors and window caption sizes.  You
'    must re-aquire your position, breadth and colors every time
'    just before you draw.

Option Explicit

' Setting this higher than 16 on 256 color systems, 32 on high color
' systems or 64 on true color systems does not improve the quality.
Private Const NumOfSteps As Long = 32

#If Win32 Then
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
#Else
Private Type RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type
#End If

' This constant is for making brushes.
Private Const PS_NULL As Long = 5&

' These constants are for getting the size of the window.
Private Const SM_CYCAPTION As Long = 4&
Private Const SM_CXBORDER As Long = 5&
Private Const SM_CYBORDER As Long = 6&
Private Const SM_CXSIZE As Long = 30&
Private Const SM_CYSIZE As Long = 31&
Private Const SM_CXFRAME As Long = 32&
Private Const SM_CYFRAME As Long = 33&

' These are for geting the system colors.
Private Const COLOR_ACTIVECAPTION As Long = 2&
Private Const COLOR_INACTIVECAPTION As Long = 3&
Private Const COLOR_CAPTIONTEXT As Long = 9&

' These are for text placement.
Private Const DT_LEFT As Long = &H0&
Private Const DT_CENTER As Long = &H1&
Private Const DT_RIGHT As Long = &H2&

' This is for drawing the text.
Private Const TRANSPARENT As Long = 1

#If Win32 Then
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDc As Long, ByVal mode As Long) As Long
Private Declare Function GetVersion& Lib "kernel32" ()
#Else
Private Declare Function CreatePen Lib "gdi" (ByVal nPenStyle As Integer, ByVal nWidth As Integer, ByVal crColor As Long) As Integer
Private Declare Function CreateSolidBrush Lib "gdi" (ByVal crColor As Long) As Integer
Private Declare Function DrawText Lib "User" (ByVal hDc As Integer, ByVal lpStr As String, ByVal nCount As Integer, lpRect As RECT, ByVal wFormat As Integer) As Integer
Private Declare Function DeleteObject Lib "gdi" (ByVal hObject As Integer) As Integer
Private Declare Function GetSysColor Lib "User" (ByVal nIndex As Integer) As Long
Private Declare Function GetSystemMetrics Lib "User" (ByVal nIndex As Integer) As Integer
Private Declare Function GetWindowDC Lib "User" (ByVal hWnd As Integer) As Integer
Private Declare Function GetWindowLong Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Long
Private Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, lpRect As RECT)
Private Declare Function GetVersion& Lib "kernel" ()
Private Declare Function Rectangle Lib "gdi" (ByVal hDc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
Private Declare Function ReleaseDC Lib "User" (ByVal hWnd As Integer, ByVal hDc As Integer) As Integer
Private Declare Function SelectObject Lib "gdi" (ByVal hDc As Integer, ByVal hObject As Integer) As Integer
Private Declare Function SetBkColor Lib "gdi" (ByVal hDc As Integer, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi" (ByVal hDc As Integer, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi" (ByVal hDc As Integer, ByVal mode As Integer) As Integer
#End If

' These are messages we intercept or submessages of WM_SYSCOMMAND.
Private Const WM_NCPAINT As Long = &H85&
Private Const WM_NCACTIVATE As Long = &H86&
Private Const WM_SETTEXT As Long = &HC&
Private Const WM_SYSCOMMAND As Long = &H112&
Private Const SC_MINIMIZE As Long = &HF020&
Private Const SC_RESTORE As Long = &HF120&

' Global variables
Dim CurrentCaption As String    ' The text to be printed in the caption bar.
Dim CurrentlyActive As Boolean  ' "True" is the form has focus, "False" if not.
Dim InWin95 As Boolean            ' "True" if in Windows 95


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
label1.Caption = (Text2.Text)
End Sub

Private Sub Command2_Click()
CommonDialog1.Flags = cdlCCRGBInit
  CommonDialog1.ShowColor
  Text1.BackColor = CommonDialog1.Color
End Sub


Private Sub btrack_Click()
Dim from As String
If (min = 0 And sec = 0) Then
    If (track > 1) Then
        from = CStr(track - 1)
    Else
        from = CStr(numTracks)
    End If
Else
    from = CStr(track)
End If
If (fPlaying) Then
    cmd = "play cd from " & from
    SendMCIString cmd, True
Else
    cmd = "seek cd to " & from
    SendMCIString cmd, True
End If
Update
End Sub

Private Sub eject_Click()
SendMCIString "set cd door open", True
Update
End Sub

Private Sub ff_Click()
Dim s As String * 40

SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", s, Len(s), 0
If (fPlaying) Then
    cmd = "play cd from " & CStr(CLng(s) + fastForwardSpeed * 1000)
Else
    cmd = "seek cd to " & CStr(CLng(s) + fastForwardSpeed * 1000)
End If
mciSendString cmd, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
End Sub

Private Sub Command1_Click()
CommonDialog1.Flags = cdlCCRGBInit
  CommonDialog1.ShowColor
  Text1.ForeColor = CommonDialog1.Color

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
label1.Caption = (Text3.Text)
End Sub

Private Sub Form_Load()

If (App.PrevInstance = True) Then
    End
End If


Timer1.Enabled = False




 



Timer1.Enabled = True
    Dim version As Long
    InWin95 = False
#If Win32 Then
    version = GetVersion()
    If (version And &H80000000) Then InWin95 = True
#Else
    version = GetVersion() And &HFFFF&
    version = version / 256
    If version = 95 Then InWin95 = True
#End If

    SubClass1.HwndParam = Me.hWnd
    CurrentCaption = "color example by stunt"
End Sub


Private Sub ftrack_Click()
If (track < numTracks) Then
    If (fPlaying) Then
        cmd = "play cd from " & track + 1
        SendMCIString cmd, True
    Else
        cmd = "seek cd to " & track + 1
        SendMCIString cmd, True
    End If
Else
    SendMCIString "seek cd to 1", True
End If
Update
End Sub

Private Sub pause_Click()
SendMCIString "pause cd", True
fPlaying = False
Update
End Sub

Private Sub play_Click()
SendMCIString "play cd", True
fPlaying = True
End Sub

Private Sub rew_Click()
Dim s As String * 40

SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", s, Len(s), 0
If (fPlaying) Then
    cmd = "play cd from " & CStr(CLng(s) - fastForwardSpeed * 1000)
Else
    cmd = "seek cd to " & CStr(CLng(s) - fastForwardSpeed * 1000)
End If
mciSendString cmd, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
End Sub

Private Sub stopbtn_Click()
SendMCIString "stop cd wait", True
cmd = "seek cd to " & track
SendMCIString cmd, True
fPlaying = False
Update
End Sub

#If Win32 Then
Private Sub SubClass1_WndMessage(hWnd As Long, msg As Long, wp As Long, lp As Long, RetVal As Long, nodef As Integer)
#Else
Private Sub SubClass1_WndMessage(hWnd As Integer, msg As Integer, wp As Integer, lp As Long, RetVal As Long, nodef As Integer)
#End If
    Dim ActiveColor As Long
    Dim InactiveColor As Long
    Dim TextColor As Long
    ActiveColor = GetSysColor(COLOR_ACTIVECAPTION)
    InactiveColor = GetSysColor(COLOR_INACTIVECAPTION)
    TextColor = GetSysColor(COLOR_CAPTIONTEXT)
    
    Select Case msg
        Case WM_NCACTIVATE
            If wp = 0 Then
                CurrentlyActive = False
                Me.Caption = CurrentCaption
            Else
                CurrentlyActive = True
                Me.Caption = ""
                DrawTitleBar Me.hWnd, CurrentCaption, ActiveColor, TextColor
            End If
        Case WM_NCPAINT:
            If CurrentlyActive = False Then
            Else
                DrawTitleBar Me.hWnd, CurrentCaption, ActiveColor, TextColor
            End If
        Case WM_SYSCOMMAND:
            If (wp And &HFFF0) = SC_MINIMIZE Then
                Me.Caption = CurrentCaption
            ElseIf (wp And &HFFF0) = SC_RESTORE Then
                Me.Caption = ""
                DrawTitleBar Me.hWnd, CurrentCaption, ActiveColor, TextColor
            End If
    End Select
End Sub

' This is the actual function that does the drawing.
#If Win32 Then
Public Sub DrawTitleBar(hWnd As Long, Caption As String, DrawColor As Long, TextColor As Long)
#Else
Public Sub DrawTitleBar(hWnd As Integer, Caption As String, DrawColor As Long, TextColor As Long)
#End If
    Dim i As Integer        ' for counting
    Dim lTop As Long        ' top of the caption
    Dim lBottom As Long     ' bottom    "   "
    Dim lLeft As Long       ' left edge "   "
    Dim lWidth As Long      ' width     "   "
    Dim hDc As Long
    Dim hBrush As Long
    Dim hOldBrush As Long
    Dim hPen As Long
    Dim hOldPen As Long
    Dim WindowRect As RECT  ' bounds of the window
    Dim CaptionRect As RECT ' bounds of the caption
    Dim DrawColorR As Long  ' Red component of the drawing color
    Dim DrawColorG As Long  ' Green "       "
    Dim DrawColorB As Long  ' Blue  "       "
    Dim RetVal As Long      ' for holding the return value
    Dim div(NumOfSteps) As Double

    ' First we pre-compute all the fractions of 16.  We use this to
    ' make 16 blocks in the caption bar with 16 different shades.
    ' Precomputing the numbers makes this routine about twice as
    ' fast as using " * (i/16)" in all the math routines.
    ' To change it from 16 to another number of steps, change the
    ' constant "NumOfSteps" to some other number.
    
    For i = 0 To NumOfSteps
        div(i) = i / NumOfSteps
    Next i

    hDc = GetWindowDC(hWnd)
    If hDc = 0 Then GoTo ErrorCondition

    RetVal = SetTextColor(hDc, TextColor)

#If Win32 Then
    RetVal = GetWindowRect(hWnd, WindowRect)
    If RetVal = 0 Then GoTo ErrorCondition
#Else
    GetWindowRect hWnd, WindowRect
#End If
    
    ' Find the caption bar's origin. This window has a system box
    ' a minimize box, a maximize box, and has a resizeable frame
    lLeft = GetSystemMetrics(SM_CXSIZE) + GetSystemMetrics(SM_CXBORDER) + GetSystemMetrics(SM_CXFRAME)
    lTop = GetSystemMetrics(SM_CYFRAME)
    
    ' (WindowRect.Right - WindowRect.Left) is the width of the window.
    ' lLeft currently holds the width of the frame, border and one button.
    ' Double it to get the width of the objects on both sides of the window
    ' and subtract another SM_CXSIZE for the minimize button. If the window
    ' does not have min or max buttons, be sure to add the proper length to
    ' the width.  If you are using the Windows 95 shell, then you must
    ' subtract enough space for another button.
    lWidth = (WindowRect.Right - WindowRect.Left) - (2 * lLeft) - GetSystemMetrics(SM_CXSIZE) - 1
    If (InWin95 = True) Then lWidth = lWidth - GetSystemMetrics(SM_CXSIZE) + 4

    lBottom = GetSystemMetrics(SM_CYSIZE) + GetSystemMetrics(SM_CYFRAME) + 1
    
    hPen = CreatePen(PS_NULL, 0, 0)
    If hPen = 0 Then GoTo ErrorCondition

    ' Split the color to its red, green and blue components so that
    ' we can calculate the brightness evenly.
    DrawColorB = DrawColor \ 65536
    If (DrawColor >= 65536) Then DrawColor = DrawColor Mod 65536
    DrawColorG = DrawColor \ 256
    If (DrawColor >= 256) Then DrawColor = DrawColor Mod 256
    DrawColorR = DrawColor

    ' Now draw brightening effect.
    For i = 0 To (NumOfSteps / 2) - 1
        hBrush = CreateSolidBrush(RGB(DrawColorR * div(NumOfSteps - (i * 2)), DrawColorG * div(NumOfSteps - (i * 2)), DrawColorB * div(NumOfSteps - (i * 2))))
        If hBrush = 0 Then GoTo ErrorCondition

        hOldBrush = SelectObject(hDc, hBrush)
        hOldPen = SelectObject(hDc, hPen)
        ' These rectangles are for the left side of the caption.
        ' The integer division used here could result in gaps between blocks.
        ' This is solved by the "+2" to expand the width to overlap the blocks
        ' on either side without overdrawing the black border of the Command Button.
        RetVal = Rectangle(hDc, lLeft + (lWidth * div(i)), lTop, lLeft + (lWidth * div(i + 1)) + 2, lBottom)
        ' These rectanges are for the right side of the caption.
        ' The integer division used here could result in gaps between blocks.
        ' This is solved by the "-1" and "+1" to expand the width to overlap
        ' the blocks on either side.
        RetVal = Rectangle(hDc, lLeft + (lWidth * div(NumOfSteps - i - 1)) - 1, lTop, lLeft + (lWidth * div(NumOfSteps - i)) + 1, lBottom)
        RetVal = SelectObject(hDc, hOldBrush)
        RetVal = SelectObject(hDc, hOldPen)
        RetVal = DeleteObject(hBrush)
    Next i
    
    ' Now create the rectangle to draw the text in.
    CaptionRect.Top = lTop
    CaptionRect.Left = lLeft + 2 ' the "+2" is to give the text a little space
    CaptionRect.Bottom = lBottom
    CaptionRect.Right = lLeft + lWidth
    RetVal = SetBkMode(hDc, TRANSPARENT)
    RetVal = DrawText(hDc, Caption, -1, CaptionRect, DT_LEFT)
    RetVal = DeleteObject(hPen)
Exit Sub
    
ErrorCondition:
    MsgBox "An error has occured in the drawing routine.  This is probably due to low memory"
    Exit Sub
End Sub

