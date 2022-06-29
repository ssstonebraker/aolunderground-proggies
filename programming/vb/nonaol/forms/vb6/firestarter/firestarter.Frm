VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fire STARTer by NILESH P KURHADE - bluenile5@hotmail.com"
   ClientHeight    =   3705
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3705
   ScaleWidth      =   7335
   Visible         =   0   'False
   Begin VB.PictureBox picFire 
      Height          =   525
      Left            =   3510
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   12
      Top             =   810
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Index           =   11
      Left            =   2340
      Picture         =   "firestarter.frx":0000
      ScaleHeight     =   285
      ScaleWidth      =   795
      TabIndex        =   11
      Top             =   810
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Index           =   10
      Left            =   4950
      Picture         =   "firestarter.frx":0E5A
      ScaleHeight     =   285
      ScaleWidth      =   795
      TabIndex        =   10
      Top             =   2580
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Index           =   9
      Left            =   6150
      Picture         =   "firestarter.frx":1CB4
      ScaleHeight     =   285
      ScaleWidth      =   795
      TabIndex        =   9
      Top             =   2790
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Index           =   8
      Left            =   5370
      Picture         =   "firestarter.frx":2B0E
      ScaleHeight     =   285
      ScaleWidth      =   795
      TabIndex        =   8
      Top             =   3240
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Index           =   7
      Left            =   4440
      Picture         =   "firestarter.frx":3968
      ScaleHeight     =   285
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   3210
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Index           =   6
      Left            =   3480
      Picture         =   "firestarter.frx":47C2
      ScaleHeight     =   285
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   3210
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Index           =   5
      Left            =   3000
      Picture         =   "firestarter.frx":561C
      ScaleHeight     =   285
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   2730
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Index           =   4
      Left            =   2280
      Picture         =   "firestarter.frx":6476
      ScaleHeight     =   285
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   3180
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Index           =   3
      Left            =   1020
      Picture         =   "firestarter.frx":72D0
      ScaleHeight     =   285
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   3180
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Index           =   2
      Left            =   1740
      Picture         =   "firestarter.frx":812A
      ScaleHeight     =   285
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   2580
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Index           =   1
      Left            =   1590
      Picture         =   "firestarter.frx":8F84
      ScaleHeight     =   285
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   2130
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   3090
      Top             =   1620
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Index           =   0
      Left            =   120
      Picture         =   "firestarter.frx":9DDE
      ScaleHeight     =   285
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   390
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Const GW_NEXT = 2
Const GW_CHILD = 5
Const BM_SETIMAGE = &HF7

Dim Stbutt As Long

Private Sub Form_Load()
'Fire STARTer is a app to animate the start button
'and set it on fire.
'project undertaken in a hurry on user's request therefore
'sorry about the poor commenting.
'exe included.
'Things to do
' 1) detect the change in system colors terminate the app and
'    restart it. REASON: When the system colors are changed
'    animation is not proper ( If anyone has a better alternative
'    PLEASE let me know
' 2) register the app as a service so that it is invisble
'    to the user.
' CODE BY - NILESH P KURHADE
' EMAIL : bluenile5@hotmail.com
' very soon i will be posting the code to animate the desktop icons

'find the start button
Dim dt, hwnda As Long, Clstring As String * 14
'get the desktop window
'remember desktop is the parent of all the windows
dt = GetDesktopWindow()
'get desktop's first child
hwnda = GetWindow(dt, GW_CHILD)
Do While hwnda <> 0
    'check if the handle is of Start Button's parent window
    GetClassName hwnda, Clstring, 14
    If Left(Clstring, 13) = "Shell_TrayWnd" Then
        'get the child of Shell_TrayWnd
        hwnda = GetWindow(hwnda, GW_CHILD)
        Do While hwnda <> 0
            'check if the handle is of Start Button
            GetClassName hwnda, Clstring, 7
            If Left(Clstring, 6) = "Button" Then
                Stbutt = hwnda
                'found the start button now animate it in
                'the timer
                Timer1.Enabled = True
                Exit Sub
            End If
            'get the next window
            hwnda = GetWindow(hwnda, GW_NEXT)
        Loop
    End If
    'get the next window
    hwnda = GetWindow(hwnda, GW_NEXT)
Loop
End Sub
Private Sub Timer1_Timer()
    Static i As Integer
    picFire.Picture = Picture1(i).Picture
    'change the start button picture
    PostMessage Stbutt, BM_SETIMAGE, 0, picFire.Picture.Handle
    'increment the picture counter
    i = i + 1
    If i = 12 Then i = 0
End Sub
