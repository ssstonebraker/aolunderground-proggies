VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "©1999 RL Collision Detection"
   ClientHeight    =   4665
   ClientLeft      =   4170
   ClientTop       =   3210
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   4800
      Picture         =   "Main.frx":0000
      ScaleHeight     =   645
      ScaleWidth      =   1710
      TabIndex        =   14
      Top             =   120
      Width           =   1710
   End
   Begin VB.PictureBox picBlank 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   6720
      Picture         =   "Main.frx":0B6C
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   13
      Top             =   3300
      Width           =   975
   End
   Begin VB.PictureBox picCD 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   5160
      Picture         =   "Main.frx":0EDC
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox PicBakGnd1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2925
      Left            =   6720
      Picture         =   "Main.frx":124C
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   302
      TabIndex        =   6
      Top             =   60
      Width           =   4590
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   6840
      TabIndex        =   2
      Top             =   4500
      Width           =   2835
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   1440
         Picture         =   "Main.frx":1867
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.PictureBox picCharacter 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   540
         Index           =   0
         Left            =   120
         Picture         =   "Main.frx":1C31
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   7860
      TabIndex        =   1
      Top             =   3180
      Width           =   2955
      Begin VB.PictureBox picMask1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   1380
         Picture         =   "Main.frx":2247
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   61
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.PictureBox picCharacter 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   540
         Index           =   2
         Left            =   180
         Picture         =   "Main.frx":2610
         ScaleHeight     =   480
         ScaleWidth      =   915
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2925
      Left            =   120
      Picture         =   "Main.frx":2C27
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   302
      TabIndex        =   0
      Top             =   60
      Width           =   4590
   End
   Begin VB.Label Label7 
      Caption         =   $"Main.frx":3242
      Height          =   675
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   6435
   End
   Begin VB.Label Label6 
      Caption         =   $"Main.frx":332A
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   3420
      Width           =   6435
   End
   Begin VB.Label Label5 
      Caption         =   "Hold down the right mouse button inside the pisture and move it around. "
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   3060
      Width           =   6375
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Pixel"
      Height          =   315
      Left            =   5160
      TabIndex        =   5
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Extend"
      Height          =   195
      Left            =   5160
      TabIndex        =   4
      Top             =   840
      Width           =   840
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'============================================================
'== Author  : Richard Lowe
'== Date    : July 99
'== Contact : riklowe@hotmail.com
'============================================================
'== Desciption
'==
'== This program demonstrates how to performs accurate pixel
'== Collisions
'==
'============================================================
'== Version History
'============================================================
'== 1.0  28-July-99  RL  Initial Release.
'============================================================

'------------------------------------------------------------
'Dimension variables
'------------------------------------------------------------

Dim lTask As Long
Dim MDown As Boolean
Dim bkg(64, 64) As Byte
Dim dwn%
Dim Iwidth As Integer
Dim IHeight As Integer

Dim iMPx As Integer
Dim iMPy As Integer

Private Sub Form_Load()
'------------------------------------------------------------
'Initialise variables
'------------------------------------------------------------
    Iwidth = picCharacter(0).ScaleWidth
    IHeight = picCharacter(0).ScaleHeight

    iMPx = 100
    iMPy = 50
    
'------------------------------------------------------------
'Initialise display
'------------------------------------------------------------
    BitBlt picBack.hdc, 0, 0, PicBakGnd1.Width, PicBakGnd1.Height, PicBakGnd1.hdc, 0, 0, vbSrcCopy
    
    BitBlt picBack.hdc, iMPx, iMPy, Iwidth, IHeight, picMask1.hdc, 0, 0, vbSrcAnd
    BitBlt picBack.hdc, iMPx, iMPy, Iwidth, IHeight, picCharacter(2).hdc, 0, 0, vbSrcPaint
    
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MDown = True
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'------------------------------------------------------------
'This sub perform the crude BUT FAST extent collision detection.
'Once this has been detected, the Pixel colision detection
'function is called.
'------------------------------------------------------------
    
    If MDown Then
    
'------------------------------------------------------------
'Create display
'------------------------------------------------------------
        BitBlt picBack.hdc, 0, 0, PicBakGnd1.Width, PicBakGnd1.Height, PicBakGnd1.hdc, 0, 0, vbSrcCopy
        
        BitBlt picBack.hdc, iMPx, iMPy, Iwidth, IHeight, picMask1.hdc, 0, 0, vbSrcAnd
        BitBlt picBack.hdc, iMPx, iMPy, Iwidth, IHeight, picCharacter(2).hdc, 0, 0, vbSrcPaint
                
        BitBlt picBack.hdc, x, y, Iwidth, IHeight, picMask.hdc, 0, 0, vbSrcAnd
        BitBlt picBack.hdc, x, y, Iwidth, IHeight, picCharacter(0).hdc, 0, 0, vbSrcPaint
        
'------------------------------------------------------------
'Detect Extent collisions, and call pixel collision detect function
'------------------------------------------------------------
        If (x + Iwidth > iMPx) And (x < iMPx + Iwidth) And (y + IHeight > iMPy) And (y < iMPy + IHeight) Then
            Label3 = "Collision"
            If CollisionDetect(x, y, picMask, iMPx, iMPy, picMask1, picBlank) Then
                Label4 = "Collision"
            Else
                Label4 = ""
            End If
        Else
            BitBlt picCD.hdc, 0, 0, picBlank.ScaleWidth, picBlank.ScaleHeight, picBlank.hdc, 0, 0, vbNotSrcCopy
            Label3 = ""
            Label4 = ""
        End If
        
        picBack.Refresh

    End If
End Sub

Private Sub picBack_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MDown = False
End Sub

Private Sub Picture1_Click()
    frmAbout.Show vbModal
End Sub
