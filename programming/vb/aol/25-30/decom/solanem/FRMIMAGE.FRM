VERSION 2.00
Begin Form frmImageMailer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mailing Options"
   ClientHeight    =   2565
   ClientLeft      =   1620
   ClientTop       =   2175
   ClientWidth     =   5835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   3255
   Left            =   1560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5835
   Top             =   1545
   Width           =   5955
   Begin ComboBox cmbRoom 
      FontBold        =   0   'False
      FontItalic      =   -1  'True
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   330
      Left            =   195
      TabIndex        =   11
      Top             =   1950
      Width           =   1350
   End
   Begin SSFrame fraRoomBuster 
      Caption         =   "Room Buster:"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      Height          =   825
      Left            =   75
      TabIndex        =   10
      Top             =   1605
      Width           =   3000
   End
   Begin TextBox txtMessage 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1935
      Left            =   3285
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "Free iMaGe Mass Mail!  Enjoy your warez! =)"
      Top             =   405
      Width           =   2295
   End
   Begin SSFrame fraMessage 
      Caption         =   "Mailing Message:"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      Height          =   2340
      Left            =   3165
      TabIndex        =   8
      Top             =   90
      Width           =   2550
   End
   Begin SSOption optRange 
      Caption         =   "Select Range"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin SSOption optAll 
      Caption         =   "Send All"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   675
      Width           =   1335
   End
   Begin SSOption optCurrent 
      Caption         =   "Send Current"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   405
      Width           =   1335
   End
   Begin SSCheck chkSignOffAfter 
      Caption         =   "Sign-Off After"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1215
      Width           =   1335
   End
   Begin SSCheck chkIMFull 
      Caption         =   "IM Full Boxes"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   945
      Width           =   1335
   End
   Begin SSCheck chkStartatBottom 
      Caption         =   "Start at Bottom"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   675
      Width           =   1335
   End
   Begin SSCheck chkKeepasNew 
      Caption         =   "Keep as New"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   375
      Width           =   1335
   End
   Begin SSFrame fraMailOptions 
      Caption         =   "Mailing Options:"
      Font3D          =   0  'None
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      Height          =   1485
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   3000
   End
   Begin Menu mnuMain 
      Caption         =   "&Main"
      Begin Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin Menu mnuLogg 
         Caption         =   "&Logg?"
      End
      Begin Menu mnuMainSep 
         Caption         =   "-"
      End
      Begin Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Sub cmbRoom_KeyPress (keyascii As Integer)
If keyascii = 13 Then
    If Len(cmbroom.Text) > 20 Then Exit Sub
        For x = 0 To cmbroom.ListCount - 1
            If cmbroom.List(x) = Trim$(cmbroom.Text) Then Exit Sub
        Next x
cmbroom.AddItem Trim$(cmbroom.Text): cmbroom.Text = ""
End If
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
        hBrush% = CreateSolidBrush(RGB(0, 0, red))
        retVal% = FillRect(TheForm.hDC, FillArea, hBrush)
        retVal% = DeleteObject(hBrush)
        red = red - 4
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
x = setfocusapi(cmbroom.hWnd)
End Sub

Sub make3d (Frm As Form, Ctl As Control)
Frm.ScaleMode = 3
Frm.CurrentX = Ctl.Left - 1
Frm.CurrentY = Ctl.Top + Ctl.Height
Frm.Line -Step(0, -(Ctl.Height + 1)), RGB(92, 92, 92)
Frm.Line -Step(Ctl.Width + 1, 0), RGB(92, 92, 92)
Frm.Line -Step(0, Ctl.Height + 1), RGB(255, 255, 255)
Frm.Line -Step(-(Ctl.Width + 1), 0), RGB(255, 255, 255)
End Sub

