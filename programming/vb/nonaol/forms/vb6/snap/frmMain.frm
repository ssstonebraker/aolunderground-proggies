VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   2295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "10"
      Top             =   240
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "# of pixels to snap at"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   1470
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''LOTS OF THANKS TO AARON
''''''''''''''''''''''''YOUNG AT REDWING FOR
''''''''''''''''''''''''TEACHING ME THIS CODE
''''''''''''''''''''''''
Private Type RECT      '
        Left As Long   '
        Top As Long    '
        Right As Long  '''''''''RECTANGLE TYPE
        Bottom As Long '
End Type               '
''''''''''''''''''''''''
''''''''''''''''''''''''GETS WINDOW RECTANGULAR DIMENSIONS
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

''''''''''''''''''''''''GETS A WINDOW'S HANDLE (TASKBAR)
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

''''''''''''''''''''''''OUR RECT TYPE (PUBLIC/PRIVATE/GLOBAL)
Private tRect As RECT
Private Sub Form_Load()
''''''''''''''''''''''''CALL FUNCTION TO GET DIMENSTIONS FOR tRect
Call GetWindowRect(FindWindowEx(0&, 0&, "Shell_TrayWnd", vbNullString), tRect)
End Sub

Private Sub Timer1_Timer()
''''''''''''''''''''''''DETERMINES WHERE WINDOW IS AND WHERE IT SHOULD
''''''''''''''''''''''''SNAP TO
DoEvents

If (Left >= -ScaleX(Val(Text1.Text), vbPixels, vbTwips) And Left <= ScaleX(Val(Text1.Text), vbPixels, vbTwips)) And (Top >= -ScaleY(Val(Text1.Text), vbPixels, vbTwips) And Top <= ScaleY(Val(Text1.Text), vbPixels, vbTwips)) Then
'Topleft snap
    Top = 0
    Left = 0

ElseIf (Top + Height <= ScaleY(tRect.Top, vbPixels, vbTwips) + ScaleY(Val(Text1.Text), vbPixels, vbTwips) And Top + Height >= ScaleY(tRect.Top, vbPixels, vbTwips) - ScaleY(Val(Text1.Text), vbPixels, vbTwips)) And (Left >= -ScaleX(Val(Text1.Text), vbPixels, vbTwips) And Left <= ScaleX(Val(Text1.Text), vbPixels, vbTwips)) Then
'Bottomleft snap
    Top = ScaleY(tRect.Top, vbPixels, vbTwips) - Height
    Left = 0

ElseIf (Top + Height <= ScaleY(tRect.Top, vbPixels, vbTwips) + ScaleY(Val(Text1.Text), vbPixels, vbTwips) And Top + Height >= ScaleY(tRect.Top, vbPixels, vbTwips) - ScaleY(Val(Text1.Text), vbPixels, vbTwips)) And (Left + Width <= Screen.Width + ScaleX(Val(Text1.Text), vbPixels, vbTwips) And Left + Width >= Screen.Width - ScaleX(Val(Text1.Text), vbPixels, vbTwips)) Then
'Bottomright snap
    Top = ScaleY(tRect.Top, vbPixels, vbTwips) - Height
    Left = Screen.Width - Width

ElseIf (Top >= -ScaleY(Val(Text1.Text), vbPixels, vbTwips) And Top <= ScaleY(Val(Text1.Text), vbPixels, vbTwips)) And (Left + Width <= Screen.Width + ScaleX(Val(Text1.Text), vbPixels, vbTwips) And Left + Width >= Screen.Width - ScaleX(Val(Text1.Text), vbPixels, vbTwips)) Then
'Topright snap
    Top = 0
    Left = Screen.Width - Width

ElseIf Top >= -ScaleY(Val(Text1.Text), vbPixels, vbTwips) And Top <= ScaleY(Val(Text1.Text), vbPixels, vbTwips) Then
'Top snap
    Top = 0

ElseIf Left >= -ScaleX(Val(Text1.Text), vbPixels, vbTwips) And Left <= ScaleX(Val(Text1.Text), vbPixels, vbTwips) Then
'Left snap
    Left = 0

ElseIf Top + Height <= ScaleY(tRect.Top, vbPixels, vbTwips) + ScaleY(Val(Text1.Text), vbPixels, vbTwips) And Top + Height >= ScaleY(tRect.Top, vbPixels, vbTwips) - ScaleY(Val(Text1.Text), vbPixels, vbTwips) Then
'Bottom snap
    Top = ScaleY(tRect.Top, vbPixels, vbTwips) - Height

ElseIf Left + Width <= Screen.Width + ScaleX(Val(Text1.Text), vbPixels, vbTwips) And Left + Width >= Screen.Width - ScaleX(Val(Text1.Text), vbPixels, vbTwips) Then
'Right snap
    Left = Screen.Width - Width

End If

End Sub


