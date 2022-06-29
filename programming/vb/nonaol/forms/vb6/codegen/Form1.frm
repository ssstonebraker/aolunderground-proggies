VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code Gen. Example by PAT or JK"
   ClientHeight    =   3135
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4275
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2100
      TabIndex        =   2
      Text            =   "Window Handle"
      Top             =   2580
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Code Generator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4035
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Text            =   "Form1.frx":0000
         Top             =   300
         Width           =   3675
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   720
         MouseIcon       =   "Form1.frx":0014
         Picture         =   "Form1.frx":031E
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   4
         Top             =   2100
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Create Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   3
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Drag and Drop The Ring"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   2220
         Width           =   2115
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' FindWindow Code Generator Example by PAT or JK
' This is an example on how to make a code generator
' program to generate the code to find windows.
' If you use any part of this example to make a program,
' make sure and give some credit to patorjk

' contact:
' patorjk@aol.com
' http://www.patorjk.com/

Dim PictureThis As StdPicture

Private Sub Command1_Click()
Dim WinHandle As Long, X As Long, WinClassName As String
Dim WinParent As Long, i As Integer, TheCN As Long
Dim i2 As Integer, i3 As Integer, WindowHandle As Long
Dim TheParentClassName As String * 100
ReDim HoldWinClassNames(0) As String ' array to hold window class names
ReDim HoldVarNames(0) As String ' array to hold variable names
ReDim HoldWinHanles(0) As Long ' array to hold window handles

On Error Resume Next ' Just in case we bump into some kind of error

Text1.Text = "" ' clear the code box
WindowHandle = CLng(Text2.Text) ' set the handle of the window to find

' set the first value of the HoldWinHanles array to the window
' you want to create code for.
HoldWinHanles(0) = WindowHandle

' set the first value of the HoldWinClassNames array to the
' class name of the window you want to create code for.
WinClassName = String(100, Chr$(0))
X = GetClassName(WindowHandle, WinClassName, 100)
HoldWinClassNames(0) = Chr$(34) & Left(WinClassName, X) & Chr$(34)

' Get handle of parent window
WinParent = getparent(WindowHandle)
         
i = 1
Do While WinParent <> 0
    ' in this loop we go up and find the handles and
    ' class names of all the parent windows
    ReDim Preserve HoldWinClassNames(i) ' make the array a little bigger
    ReDim Preserve HoldWinHanles(i) ' make the array a little bigger
    TheCN = GetClassName(WinParent, TheParentClassName, 100)
    HoldWinClassNames(i) = Chr$(34) & Left(TheParentClassName, TheCN) & Chr$(34)
    HoldWinHanles(i) = WinParent
    WinParent = getparent(WinParent)
    i = i + 1
Loop

For i2 = 0 To UBound(HoldWinClassNames, 1)
    ' here we're creating the variable names we'll use
    ReDim Preserve HoldVarNames(i2) ' make the array a little bigger
    HoldVarNames(i2) = ExHandle(HoldWinClassNames(i2))
    If HoldVarNames(i2) = "" Then
        HoldVarNames(i2) = "x"
    End If
Next

' Here's where we loop through and create the code

' We start at the window we want to make the code for
' and loop up till there isn't any more windows
For i2 = 0 To UBound(HoldWinClassNames, 1)
    ' get the handle of the parent window of the window we're looking at
    X = FindWindowEx(getparent(HoldWinHanles(i2)), 0&, Mid$(HoldWinClassNames(i2), 2, Len(HoldWinClassNames(i2)) - 2), vbNullString)
    ' If we're not looking at the last window, and this
    ' window's handle isn't zero then...
    If i2 <> UBound(HoldWinClassNames, 1) And X <> 0 Then
        If HoldWinHanles(i2) <> X And X <> 0 Then
            Do While HoldWinHanles(i2) <> X
                ' loop through the different sibling windows
                X = FindWindowEx(getparent(HoldWinHanles(i2)), X, Mid$(HoldWinClassNames(i2), 2, Len(HoldWinClassNames(i2)) - 2), vbNullString)
                Text1.Text = HoldVarNames(i2) & "& = FindWindowEx(" & HoldVarNames(i2 + 1) & "&, " & HoldVarNames(i2) & "&, " & HoldWinClassNames(i2) & ", vbNullString)" & Chr(13) & Chr(10) & Text1.Text
            Loop
            Text1.Text = HoldVarNames(i2) & "& = FindWindowEx(" & HoldVarNames(i2 + 1) & "&, 0&, " & HoldWinClassNames(i2) & ", vbNullString)" & Chr(13) & Chr(10) & Text1.Text
        Else
            Text1.Text = HoldVarNames(i2) & "& = FindWindowEx(" & HoldVarNames(i2 + 1) & "&, 0&, " & HoldWinClassNames(i2) & ", vbNullString)" & Chr(13) & Chr(10) & Text1.Text
        End If
    Else
        ' This is either the last window or
        ' a window where the parent of this
        ' window's handle is zero
        Text1.Text = HoldVarNames(i2) & "& = FindWindow(" & HoldWinClassNames(i2) & ", vbNullString)" & Chr(13) & Chr(10) & Text1.Text
    End If
Next

' We've created all the code so it's
' time to declare the variables

' in these two loops we check to make sure we don't
' declare the same variable twice
For i2 = 0 To UBound(HoldWinClassNames, 1)
    For i3 = 1 To UBound(HoldWinClassNames, 1)
        If i2 <> i3 Then
            If HoldVarNames(i2) = HoldVarNames(i3) Then
            HoldVarNames(i2) = ""
            End If
        End If
    Next
Next

' Now we loop through and declare the variables
For i2 = 0 To UBound(HoldWinClassNames, 1)
    If Trim(HoldVarNames(i2)) <> "" Then
        Text1.Text = "Dim " & HoldVarNames(i2) & "&" & Chr(13) & Chr(10) & Text1.Text
    End If
Next
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Here we change the mouse cursor into a picture
' of a ring. To make it look more realistic the
' ring in picture1 will disappear while the mouse
' button is down and then reappear when the mouse
' button is let up on.

' save the picture of the ring in the PictureThis variable
Set PictureThis = Picture1.Picture
' set the cursor equal to the ring picture
Picture1.MousePointer = 99
' set picture1.picture equal to nothing
Picture1.Picture = Me.Picture
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TheWindow As Long, OldWindow As Long, CPos As POINTAPI

' If the left mouse button is down then...
If Button = 1 Then
    Call GetCursorPos(CPos)
    ' set TheWindow equal to the handle of the window
    ' the mouse is over
    TheWindow = WindowFromPoint(CPos.X, CPos.Y)
    ' make sure we're not over the same window
    If TheWindow <> OldWindow Then
        Text2.Text = CStr(TheWindow)
    End If
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' set the cursor equal to it's default
Picture1.MousePointer = 0
' set picture1.picture to the picture of the ring
Set Picture1.Picture = PictureThis
' Click the Command1 button
Command1_Click
End Sub
