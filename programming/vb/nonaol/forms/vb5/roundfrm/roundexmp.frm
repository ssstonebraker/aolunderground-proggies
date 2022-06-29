VERSION 5.00
Begin VB.Form rndform 
   BorderStyle     =   0  'None
   Caption         =   "Round form example"
   ClientHeight    =   2340
   ClientLeft      =   4920
   ClientTop       =   3360
   ClientWidth     =   2310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "roundexmp.frx":0000
   ScaleHeight     =   2340
   ScaleWidth      =   2310
   Begin VB.CommandButton min 
      Caption         =   "_"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton unload 
      Caption         =   "X"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "rndform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long 'used in form mouse down
Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_DIFF = 4
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Function CreateFormRegion() As Long
    Dim circlearea As Long, holdingarea As Long, objct As Long, nRet As Long
    Dim PolyPoints() As POINTAPI
    circlearea = CreateRectRgn(0, 0, 0, 0)
    holdingarea = CreateRectRgn(0, 0, 0, 0)
    objct = CreateEllipticRgn(0, 14, 144, 152)
    nRet = CombineRgn(circlearea, objct, objct, RGN_COPY)
    DeleteObject objct
    CreateFormRegion = circlearea
End Function
Private Sub Form_Load()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion, True)
'good example of how to make a shaped form with out
'formshaper.ocx
'made by: Þútõ² vb6@att.net
'e-mail me if you need help, not so shure of how much help
' i can be (i still dont understand 100%, just well enough
'to make this work) but i will try
'i am very proud of submitting something to KNK
' they taught me to program and i owe all this to them
' now im just gonna bullshit so you dont have to read this:
' no one knows all of VB, as KNK said, it is limitless
' if you are in a chat and someone claims to know every
'thing about vb ask them 2 things
'      1. do you use vb3?
'      2. do you know API?
'if the answer to number 1 is yes, they know jack shit
' about programming
' if they know API ask them for examples. not stolen subs,
' simple examples
' And the most important thing any programmer can do:
'                     HELP A BEGINER!!
'Think, at one time dident we all use vb3 and go
'looking for free code to steal? i shure as hell did
'and ill admit that. I dont want nor like to steal code,
'i try and code it myself before looking for help.
'but help is always a good thing. if you need any help, no matter
'how dumb hard or strange the question is im me (AIM:fishbumjr)
'or ICQ:35981084 or email vb6@att.net
'i cant promise ill be of help but i will try
'This was made on VB5 ent.(got VB6, most dont so i used VB5)
'if you use this or any varation e-mail me
'if you put me in the greets ill be your best friend (lol)
'last: i aint gonna bullshit anymore, your probally about
'ready to kill your computer now so later all and if you
'like this e-mail me, id really appreciate it
'Þútõ²
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'enables dragging of form, remove and see what happens
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub min_Click()
WindowState = vbMinimized
'if you dont do show in taskbar it gets wierd, try it
End Sub

Private Sub unload_Click()
End
End Sub
