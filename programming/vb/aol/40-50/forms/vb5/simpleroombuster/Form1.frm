VERSION 5.00
Begin VB.Form main 
   BorderStyle     =   0  'None
   Caption         =   "Illusion For AOL 4.0"
   ClientHeight    =   2805
   ClientLeft      =   3675
   ClientTop       =   1860
   ClientWidth     =   1665
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   187
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   111
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tries:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "For AOL 4.0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Illusion RoomBuster"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   645
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image Image9 
      Height          =   270
      Left            =   840
      Picture         =   "Form1.frx":0000
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image Image8 
      Height          =   270
      Left            =   120
      Picture         =   "Form1.frx":0E26
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image Image7 
      Height          =   270
      Left            =   360
      Picture         =   "Form1.frx":1C4C
      Top             =   720
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   645
   End
   Begin VB.Image Image6 
      Height          =   270
      Left            =   840
      Picture         =   "Form1.frx":2BC3
      Top             =   1800
      Width           =   645
   End
   Begin VB.Image Image5 
      Height          =   270
      Left            =   120
      Picture         =   "Form1.frx":39E9
      Top             =   1800
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   840
      Picture         =   "Form1.frx":480F
      Top             =   1320
      Width           =   645
   End
   Begin VB.Image Image3 
      Height          =   270
      Left            =   120
      Picture         =   "Form1.frx":5635
      Top             =   1320
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   0
      Picture         =   "Form1.frx":645B
      Top             =   0
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   0
      Picture         =   "Form1.frx":7F56
      Top             =   600
      Width           =   1620
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Public Sub GlassifyForm(frm As Form)

Const RGN_DIFF = 4
Const RGN_OR = 2

Dim outer_rgn As Long
Dim inner_rgn As Long
Dim wid As Single
Dim hgt As Single
Dim border_width As Single
Dim title_height As Single
Dim ctl_left As Single
Dim ctl_top As Single
Dim ctl_right As Single
Dim ctl_bottom As Single
Dim control_rgn As Long
Dim combined_rgn As Long
Dim ctl As Control

    If WindowState = vbMinimized Then Exit Sub

    ' Create the main form region.
    wid = ScaleX(Width, vbTwips, vbPixels)
    hgt = ScaleY(Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wid, hgt)

    border_width = (wid - ScaleWidth) / 2
    title_height = hgt - border_width - ScaleHeight
    inner_rgn = CreateRectRgn( _
        border_width, _
        title_height, _
        wid - border_width, _
        hgt - border_width)

    ' Subtract the inner region from the outer.
    combined_rgn = CreateRectRgn(0, 0, 0, 0)
    CombineRgn combined_rgn, outer_rgn, _
        inner_rgn, RGN_DIFF

    ' Create the control regions.
    For Each ctl In Controls
        If ctl.Container Is frm Then
            ctl_left = ScaleX(ctl.Left, frm.ScaleMode, vbPixels) _
                + border_width
            ctl_top = ScaleX(ctl.Top, frm.ScaleMode, vbPixels) _
                + title_height
            ctl_right = ScaleX(ctl.Width, frm.ScaleMode, vbPixels) _
                + ctl_left
            ctl_bottom = ScaleX(ctl.Height, frm.ScaleMode, vbPixels) _
                + ctl_top
            control_rgn = CreateRectRgn( _
                ctl_left, ctl_top, _
                ctl_right, ctl_bottom)
            CombineRgn combined_rgn, combined_rgn, _
                control_rgn, RGN_OR
        End If
    Next ctl

    ' Restrict the window to the region.
    SetWindowRgn hwnd, combined_rgn, True
End Sub











Private Sub Form_Load()

GlassifyForm Me
Call LoadComboBox("combo.xt", Combo1)
End Sub

Private Sub Form_Resize()
FormOnTop Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
Call SaveComboBox("combo.xt", Combo1)
End Sub

Private Sub Image2_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Label1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Top = Image3.Top
Image7.Left = Image3.Left
Image7.Visible = True
Image3.Visible = False
Label10.Caption = "0"
Call CloseWindow(FindRoom&)
time.Timer1.Enabled = True
End Sub


Private Sub Label1_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = False
Image3.Visible = True
End Sub


Private Sub Label2_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Top = Image5.Top
Image7.Left = Image5.Left
Combo1.AddItem Combo1.text
Image7.Visible = True
Image5.Visible = False

End Sub


Private Sub Label2_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = False
Image5.Visible = True
End Sub


Private Sub Label3_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Top = Image4.Top
Image7.Left = Image4.Left
Image7.Visible = True
Image4.Visible = False
time.Timer1.Enabled = False
End Sub


Private Sub Label3_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = False
Image4.Visible = True

End Sub


Private Sub Label4_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image7.Top = Image6.Top
Image7.Left = Image6.Left
Combo1.RemoveItem Combo1.ListIndex
Image7.Visible = True
Image6.Visible = False
End Sub


Private Sub Label4_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = False
Image6.Visible = True
End Sub


Private Sub Label5_Click()

Do
DoEvents

Me.Height = Trim(Str(Val(Me.Height) - 10))
Loop Until Me.Height < 600

Do
DoEvents

Me.Width = Trim(Str(Val(Me.Width) - 10))
Loop Until Me.Width < 10
If Me.Width < 10 Then Unload Me
End Sub

Private Sub Label6_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
If button = 2 Then
Me.Height = 600
Exit Sub
End If
Me.WindowState = vbMinimized
End Sub

Private Sub Label7_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
If button = 2 Then
Me.Height = 2800
Exit Sub
End If
FormDrag Me
End Sub


Private Sub Label8_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
If button = 2 Then
Me.Height = 2800
Exit Sub
End If
FormDrag Me
End Sub


