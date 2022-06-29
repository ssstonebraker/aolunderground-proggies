VERSION 2.00
Begin Form frmSignOn 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SoLuTioN Sign On MaGiK"
   ClientHeight    =   1305
   ClientLeft      =   2910
   ClientTop       =   3240
   ClientWidth     =   3330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   1710
   HelpContextID   =   40
   Left            =   2850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3330
   Top             =   2895
   Width           =   3450
   Begin CommandButton cmdExit 
      Caption         =   "&X"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontName        =   "Arial"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   360
      HelpContextID   =   40
      Left            =   2820
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin CommandButton cmdHelp 
      Caption         =   "&Help"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   360
      HelpContextID   =   40
      Left            =   2010
      TabIndex        =   4
      Top             =   840
      Width           =   825
   End
   Begin CommandButton cmdNewUser 
      Caption         =   "&New User"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   360
      HelpContextID   =   40
      Left            =   1050
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin CommandButton cmdSignOn 
      Caption         =   "&Sign On"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   360
      HelpContextID   =   40
      Left            =   90
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin TextBox txtPW 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      HelpContextID   =   40
      Left            =   105
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "sheep sex"
      Top             =   465
      Width           =   3075
   End
   Begin TextBox txtSN 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      HelpContextID   =   40
      Left            =   105
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "new-user"
      Top             =   105
      Width           =   3075
   End
End

Sub cmdExit_Click ()
Unload Me
End Sub

Sub cmdHelp_Click ()
MsgBox "SoLuTioN Sign On TooL." & Chr(13) & Chr(13) & "Just input the screen name into the top text, and the password into the bottom text, and click 'Sign On' and boom, yer online w/ that name (if its a working account).  To signon with new user use the syntax already provided, or click 'New User'.  Thats it!" & Chr(13) & Chr(13) & "btw, no i will not mail you the code on how to make this tool, please stop e-mailing/IMing me about it, thanx", 32, "SiGN oN MaGiK!"
End Sub

Sub cmdNewUser_Click ()
Call signon("new-user", "sheep sex"): DoEvents


End Sub

Sub cmdSignOn_Click ()
If Len(txtsn.Text) < 2 Then MsgBox "You have not entered a working screen name", 32, "SoLuTioN SiGN oN MaGiK": Exit Sub
If Len(txtsn.Text) < 3 Then MsgBox "You have not entered a working password", 32, "SoLuTioN SiGN oN MaGiK": Exit Sub
Call signon((txtsn.Text), (txtpw.Text)): DoEvents
End Sub

Sub fadeform (TheForm As Form)
Dim hBrush%
    Dim FormHeight%, red%, StepInterval%, x%, retVal%, OldMode%
    Dim FillArea As RECT
    OldMode = TheForm.ScaleMode
    TheForm.ScaleMode = 3  'Pixel
    FormHeight = TheForm.ScaleHeight
' Divide the form into 63 regions
    StepInterval = FormHeight \ 63
    red = 255
    FillArea.Left = 0
    FillArea.Right = TheForm.ScaleWidth
    FillArea.Top = 0
    FillArea.Bottom = StepInterval
    For x = 1 To 63
        hBrush% = CreateSolidBrush(RGB(90, 45, red))
        retVal% = FillRect(TheForm.hDC, FillArea, hBrush)
        retVal% = DeleteObject(hBrush)
        red = red - 2
        FillArea.Top = FillArea.Bottom
        FillArea.Bottom = FillArea.Bottom + StepInterval
    Next
' Fill the remainder of the form with black
    FillArea.Bottom = FillArea.Bottom + 63
    hBrush% = CreateSolidBrush(RGB(0, 0, 0))
    retVal% = FillRect(TheForm.hDC, FillArea, hBrush)
    retVal% = DeleteObject(hBrush)
    TheForm.ScaleMode = OldMode
End Sub

Sub Form_Load ()
stayontop Me
Show
Call fadeform(Me)
End Sub

Sub Form_Paint ()
Call fadeform(Me)
Call apply3d(Me, txtsn)
Call apply3d(Me, txtpw)
End Sub

