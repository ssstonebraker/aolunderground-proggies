VERSION 5.00
Begin VB.Form findit 
   BorderStyle     =   0  'None
   Caption         =   "Find Window Like..."
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_exit 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   0
      Left            =   4010
      TabIndex        =   24
      Top             =   65
      Width           =   255
   End
   Begin VB.CommandButton cmd_mini 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   1
      Left            =   3720
      TabIndex        =   25
      Top             =   65
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Winow Spy"
      Height          =   1095
      Left            =   0
      TabIndex        =   14
      Top             =   4320
      Width           =   4335
      Begin VB.PictureBox picspy 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   3720
         MouseIcon       =   "findit.frx":0000
         Picture         =   "findit.frx":08CA
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   16
         ToolTipText     =   "Drag this to the window you want to spy on..."
         Top             =   360
         Width           =   480
      End
      Begin VB.Label clas 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label hndl 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label ttl 
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "Class"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00800000&
         Caption         =   "Title:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Handle:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmd_get_child 
      Caption         =   "Get Children"
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txt_child 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmd_Get_Parents 
      Caption         =   "Get Parents"
      Height          =   285
      Left            =   3120
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txt_parents 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmd_find_by_class 
      Caption         =   "Find Class"
      Height          =   315
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txt_Class 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Text            =   "SysTabControl32"
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmd_Find_Window_Like 
      Caption         =   "Find Title"
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txt_Title 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "*Window Like*"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   2175
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1800
      Width           =   4335
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   2760
      Picture         =   "findit.frx":1194
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbl_title_bar 
      BackColor       =   &H80000002&
      Caption         =   " Find Window Like..."
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   30
      TabIndex        =   23
      Top             =   30
      Width           =   4310
   End
   Begin VB.Label Label4 
      Caption         =   "Window Handle"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Window Handle"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Window Class"
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "This Winows Handle: "
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   4200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Window Title"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1170
   End
End
Attribute VB_Name = "findit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I Wrote this example to show you how to FindWindowByTitle and FindWindowByClass
'then I added a bunch of other stuff to it..
'the API spy is easy, you can add other stuff to it,
'like getting the Parents info, not hard at alL!!
    Option Explicit
    'make sure everything is declared with this..
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    'used to set the form on top of all other windows
    Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    'gets windows such as NEXT or CHILD
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    'duh
    Private Declare Function GetWindowLW Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    'not sure but it works...
    Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    'get the parent handle
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    'get the class name
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    'get the windows text
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    'get cursor pos, used for API spy.
    Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    'gets the handle of the window from the point of the cursor
    'also for api spy..
    Private Type POINTAPI
        X   As Long
        Y   As Long
    End Type
    'used for getcursorpos
    Private Const HWND_NOTOPMOST = -2
    Private Const HWND_TOPMOST = -1
    'these 2 are used for SetWindowPos, can you figure
    'out what they do?
    Private Const SWP_NOMOVE = &H2
    Private Const SWP_NOSIZE = &H1
    Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    'these 3 are used for the FLAGS in setting the form ontop
    Private Const GWL_ID = (-12)
    Private Const GW_HWNDNEXT = 2
    Private Const GW_CHILD = 5
    'constants for GetWindow
    
    Private Spying      As Boolean
    'determines if we are spying or not.
    Private Cursor      As POINTAPI 'the cursor pos
    'put all that in a bas file.. also the function
    'make them all public if you do use in bas file, must be private within form..
Public Function FindWindowLike(hWndArray() As Long, ByVal hWndStart As Long, WindowText As String, Classname As String, ID) As Long
    Dim hwnd        As Long
    Dim R           As Long
    Static Level    As Long ' Hold the level of recursion:
    Static iFound   As Long ' Hold the number of matching windows:
        Dim sWindowText     As String
    Dim sClassName      As String
    Dim sID
    
    ' Initialize if necessary:
    If Level = 0 Then
        iFound = 0
        ReDim hWndArray(0 To 0)
        If hWndStart = 0 Then hWndStart = GetDesktopWindow()
    End If
    ' Increase recursion counter:
    Level = Level + 1
    ' Get first child window:
    hwnd = GetWindow(hWndStart, GW_CHILD)
    Do Until hwnd = 0
        DoEvents
        ' Search children by recursion:
        R = FindWindowLike(hWndArray(), hwnd, WindowText, Classname, ID)
        ' Get the window text and class name:
        sWindowText = Space(255)
        R = GetWindowText(hwnd, sWindowText, 255)
        sWindowText = Left(sWindowText, R)
        sClassName = Space(255)
        R = GetClassName(hwnd, sClassName, 255)
        sClassName = Left(sClassName, R)
        ' If window is a child get the ID:
        If GetParent(hwnd) <> 0 Then
            R = GetWindowLW(hwnd, GWL_ID)
            sID = CLng("&H" & Hex(R))
        Else
            sID = Null
        End If
        ' Check that window matches the search parameters:
        If LCase(sWindowText) Like LCase(WindowText) And LCase(sClassName) Like LCase(Classname) Then
        'you can remove the Lcase for a non case sensitive search..
            If IsNull(ID) Then
                ' If find a match, increment counter and
                '  add handle to array:
                iFound = iFound + 1
                ReDim Preserve hWndArray(0 To iFound)
                hWndArray(iFound) = hwnd
            ElseIf Not IsNull(sID) Then
                If CLng(sID) = CLng(ID) Then
                    ' If find a match increment counter and
                    '  add handle to array:
                    iFound = iFound + 1
                    ReDim Preserve hWndArray(0 To iFound)
                    hWndArray(iFound) = hwnd
                End If
            End If
        End If
        ' Get next child window:
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop
    ' Decrement recursion counter:
    Level = Level - 1
    ' Return the number of windows found:
    FindWindowLike = iFound
End Function


Private Sub cmd_exit_Click(Index As Integer)
    End
End Sub

Private Sub cmd_find_by_class_Click()
    Static hWnds()  As Long
    Dim R           As Integer
    Dim hW          As Integer
    Dim Class       As String
    Dim Cur         As Long
    Text2.Text = ""
    'find any window with the class of txt_class.text!
    R = FindWindowLike(hWnds(), 0, "*", txt_Class.Text, Null)
    'find it...
    'the "*" is specifying that ANY title is avaliable
    If R = 1 Then
        Text2.Text = Text2.Text & "Found Handle: " & hWnds(1) & vbCrLf
        'if it was found then return the handle
    ElseIf R > 1 Then
        Text2.Text = Text2.Text & "Found " & R & " Windows" & vbCrLf
        'if more than one was found then you need to be more specific..
        For Cur = 1 To R
            Text2.Text = Text2.Text & "Window " & Cur & " handle: " & hWnds(Cur) & vbCrLf
            'this loop will return each handle that was found..
        Next Cur
    ElseIf R = 0 Then
        Text2.Text = Text2.Text & "Did Not Find any window with the class of " & txt_Class.Text
        'wasn't found..
    End If
End Sub

Private Sub cmd_Find_Window_Like_Click()
    Static hWnds()  As Long
    Dim R           As Integer
    Dim hW          As Integer
    Dim Title       As String
    Dim Cur         As Long
    Text2.Text = ""
    ' Find window with title of txt_Title.text, this doesn't have to be case sensitive..
    'if you want to be less accurate put something like
    '"*Form*" in the text box..
    'that will find any window that has Form in it, if you put
    '"Form" then it will have to be Form with no other text
    Title = txt_Title.Text
    'set the variable
    txt_Title.Text = ""
    'set text1 to null so we don't find 2 windows...
    R = FindWindowLike(hWnds(), 0, Title, "*", Null)
    'find it...
    'the "*" is specifying that ANY class is avaliable, use whatever you want
    'such as AOL CHILD... or RICHCNTL...get it?
    If R = 1 Then
        Text2.Text = Text2.Text & "Found Handle: " & hWnds(1) & vbCrLf
        'if it was found then return the handle
    ElseIf R > 1 Then
        Text2.Text = Text2.Text & "Found " & R & " Windows" & vbCrLf
        'if more than one was found then you need to be more specific..
        For Cur = 1 To R
            Text2.Text = Text2.Text & "Window " & Cur & " handle: " & hWnds(Cur) & vbCrLf
            'this loop will return each handle that was found..
        Next Cur
    ElseIf R = 0 Then
        Text2.Text = Text2.Text & "Did Not Find Window!"
        'wasn't found..
    End If
    txt_Title.Text = Title
    'restet the text
End Sub

Private Sub cmd_get_child_Click()
    If Not IsNumeric(txt_child.Text) Then MsgBox "must be a number!", vbCritical, "error": Exit Sub
    'make sure the text is numeric!
    Dim Child           As Long
    Dim OldChild        As Long
    Text2.Text = ""
    Child = GetWindow(txt_child.Text, GW_CHILD)
    'get the first child..
    If Child = 0 Then
        'there is no childern
        Text2.Text = txt_child.Text & " has no children!"
    Else
    'found children so search for more.
        Text2.Text = Text2.Text & txt_child.Text & "'s children are " & Child & vbCrLf
        Do
            DoEvents
            Child = GetWindow(Child, GW_HWNDNEXT)
            'get the next window.
            If Child <> 0 Then
                Text2.Text = Text2.Text & Child & vbCrLf
                'if it was found then let the user know
            Else
                'no more children so we can leave..
                Exit Sub
            End If
        Loop Until Child = 0
    End If
End Sub

Private Sub cmd_Get_Parents_Click()
    Dim Parent          As Long
    Dim OldParent       As Long
    If Not IsNumeric(txt_parents.Text) Then MsgBox "handles are always numbers...", vbCritical, "error": Exit Sub
    'make sure it is numeric!
    Text2.Text = ""
    Parent = GetParent(txt_parents.Text)
    'get the parent window
    If Parent = 0 Then Text2.Text = txt_parents.Text & " has no parent window.": Exit Sub
    'no parents found, that means you have the highest window
    Text2.Text = txt_parents.Text & "'s parent is " & Parent & vbCrLf
    OldParent = Parent
    'set up a variable
    Do
        DoEvents
        Parent = GetParent(OldParent)
        'get the parent of the last parent
        If Parent <> 0 Then
            'if there is a parent
            Text2.Text = Text2.Text & OldParent & "'s parent is " & Parent & vbCrLf
            'let the user know
        Else
            'if not
            Text2.Text = Text2.Text & OldParent & " has no parent window."
            'tell them there are no more parents
            Exit Sub
            'exit the sub because there is nothing else to do
        End If
        OldParent = Parent
        'set the variable for the next time around
    Loop Until Parent = 0
End Sub

Private Sub cmd_mini_Click(Index As Integer)
MsgBox Picture1.hwnd
'this just hides the form and shows it..
    With lbl_title_bar
        If Height = 5430 Then
            Height = .Top + .Height + 35
        Else
            Height = 5430
        End If
    End With
End Sub

Private Sub Form_Load()
    Label2.Caption = Label2.Caption & Me.hwnd
    'set the caption that says This Windows Handle:
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
    'form is now always on top
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Spying = False Then picspy.Picture = Picture1.Picture
    'if you are not spying anymore then set the picture back to default
End Sub

Private Sub picspy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'the clicked the mouse on the spy picture so start the spy
    Spying = True
    With lbl_title_bar
        Height = .Top + .Height + 35
        'hide the form
    End With
End Sub

Private Sub picspy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this is the window spy, not to hard...
    Static hWndLast     As Long
    'a static will stay the same the next time the event is triggerd
    'kind of like putting the Private in declarations
    Dim hwnd            As Long
    Dim hWndTitle       As String
    Dim hWndClass       As String
    Dim R               As String
    If Spying = False Then Exit Sub
    'if we are not spying then we don't need to continue
    picspy.Picture = Picture
    'clear the picture
    picspy.MousePointer = 99
    'set the mouse pointer
    Call GetCursorPos(Cursor)
    'get cursor position
    hwnd = WindowFromPointXY(Cursor.X, Cursor.Y)
    'you can see how the POINTAPI works, i hope..
    'that would be Cursor.X and Cursor.Y
    If hwnd <> hWndLast Then
        'if the mouse moved to a different window then get the
        'window info..
        hWndLast = hwnd
        'set the variable so we don't get the same window again.
        hndl.Caption = hwnd
        'let the user know the handle..
        hWndTitle = Space(255)
        'a string used for getting text..
        R = GetWindowText(hwnd, hWndTitle, 255)
        'get the text and store it to R a variable..
        ttl.Caption = Left(hWndTitle, R)
        'let the user know whats up, use the Left function to remove spaces..
        hWndClass = Space(255)
        'variable for class..
        R = GetClassName(hwnd, hWndClass, 255)
        'same as with title but with the class now..
        clas.Caption = Left(hWndClass, R)
        'let the user know whats up..
        lbl_title_bar.Caption = hndl.Caption
        'this is just to put the window handle in the
        'title bar label's caption, so the user knows
        'what is going on...
    End If
End Sub

Private Sub picspy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Spying = False
    'no longer spying
    Height = 5430
    'show the form
    lbl_title_bar.Caption = " Find Window Like..."
    'set the title back to default
    picspy.MousePointer = 0
    'set the mouse pointer to default
    picspy.Picture = Picture1.Picture
    'show the picture again..
End Sub

Private Sub txt_child_KeyPress(KeyAscii As Integer)
    'if they hit Enter then click the button
    'set keyascii to 0 to avoid the BEEP
    If KeyAscii = 13 Then
        cmd_get_child_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txt_Class_KeyPress(KeyAscii As Integer)
    'if they hit Enter then click the button
    'set keyascii to 0 to avoid the BEEP
   If KeyAscii = 13 Then
        cmd_find_by_class_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txt_parents_KeyPress(KeyAscii As Integer)
    'if they hit Enter then click the button
    'set keyascii to 0 to avoid the BEEP
    If KeyAscii = 13 Then
        cmd_Get_Parents_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txt_Title_KeyPress(KeyAscii As Integer)
    'if they hit Enter then click the button
    'set keyascii to 0 to avoid the BEEP
    If KeyAscii = 13 Then
        cmd_Find_Window_Like_Click
        KeyAscii = 0
    End If
End Sub
