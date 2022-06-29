VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1041"
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   6120
      Top             =   4080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mike Altmanshofer  Timeline Studios Software."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   -375
      TabIndex        =   1
      Top             =   4440
      Width           =   3540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "©1999  All rights reserved."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4110
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   600
      Picture         =   "frmSplash.frx":0000
      Top             =   2175
      Width           =   6000
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   120
      Picture         =   "frmSplash.frx":339E
      Top             =   300
      Width           =   5550
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
    Const CCDEVICENAME = 32
    Const CCFORMNAME = 32
    Const DM_PELSWIDTH = &H80000
    Const DM_PELSHEIGHT = &H100000

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

    Dim DevM As DEVMODE
Private Sub form_load()
autokenomode = 0
delaytime = 50
normalwidth = Screen.Width \ Screen.TwipsPerPixelX
normalheight = Screen.Height \ Screen.TwipsPerPixelY
fullmode = GetSetting(App.Title, "options", "fullmode", 0)
If fullmode = 1 Then
Call ChangeRes(640, 480)
options.fullscreenmode.Enabled = True
options.fullscreenmode.Checked = True
End If

End Sub

Private Sub Timer1_Timer()
options.Show
Call formcircle(Me, 40)
Unload Me
End Sub
Sub ChangeRes(iWidth As Single, iHeight As Single)

    Dim a As Boolean
    Dim i&
    i = 0

    Do
        a = EnumDisplaySettings(0&, i&, DevM)
        i = i + 1
    Loop Until (a = False)

        Dim b&
        DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        DevM.dmPelsWidth = iWidth
        DevM.dmPelsHeight = iHeight
        b = ChangeDisplaySettings(DevM, 0)
End Sub

Sub formcircle(frm As Form, Size As Integer)



    For e% = Size% - 1 To 0 Step -1
        frm.Left = frm.Left - e%
        frm.Top = frm.Top + (Size% - e%)
    Next e%



    For e% = Size% - 1 To 0 Step -1
        frm.Left = frm.Left + (Size% - e%)
        frm.Top = frm.Top + e%
    Next e%



    For e% = Size% - 1 To 0 Step -1
        frm.Left = frm.Left + e%
        frm.Top = frm.Top - (Size% - e%)
    Next e%



    For e% = Size% - 1 To 0 Step -1
        frm.Left = frm.Left - (Size% - e%)
        frm.Top = frm.Top - e%
    Next e%


End Sub
