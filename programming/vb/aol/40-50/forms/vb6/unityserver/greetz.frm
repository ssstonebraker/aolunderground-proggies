VERSION 5.00
Object = "{1EF6BBE0-244A-11CF-840E-444553540000}#2.0#0"; "ROTEXT32.OCX"
Begin VB.Form greetz 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Greetz"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ROTEXTLib.Rotext Rotext1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _Version        =   131072
      _ExtentX        =   8281
      _ExtentY        =   5741
      _StockProps     =   79
      Caption         =   "Rotext1"
      ForeColor       =   -2147483624
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   480
   End
End
Attribute VB_Name = "greetz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'THIS FORM IS UN IMPORTANT!
Dim angle
Dim fonts
Dim nextname

Private Sub Form_Load()
On Error Resume Next
Me.Show
nextname = 0
FormOnTop Me
Timer1.Interval = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Interval = 0
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If nextname = 0 Then Rotext1.Caption = "ChiP"
If nextname = 1 Then Rotext1.Caption = "NooB"
If nextname = 2 Then Rotext1.Caption = "Scoob"
If nextname = 3 Then Rotext1.Caption = "Blu"
If nextname = 4 Then Rotext1.Caption = "Spook"
If nextname = 5 Then Rotext1.Caption = "Azar"
If nextname = 6 Then Rotext1.Caption = "Diox"
If nextname = 7 Then Rotext1.Caption = "Goob"
If nextname = 8 Then Rotext1.Caption = "Icee"
If nextname = 9 Then Rotext1.Caption = "Perolta"
If nextname = 10 Then Rotext1.Caption = "Comet"
If nextname = 11 Then Rotext1.Caption = "FeW"
If nextname = 12 Then Rotext1.Caption = "CrackLyn"
If nextname = 13 Then Rotext1.Caption = "General"
If nextname = 14 Then Rotext1.Caption = "Phreaker"
If nextname = 15 Then Rotext1.Caption = "BaD"
If nextname = 16 Then Rotext1.Caption = "Biorn"
If nextname = 17 Then Rotext1.Caption = "SOO"
If nextname = 18 Then
    Rotext1.Caption = ""
    Timer1.Interval = 0
    Me.Hide
End If
If Rotext1.Left <> 0 Then Rotext1.Left = 0
For x = 1 To 45
Rotext1.angle = angle + 1
Next x
Rotext1.FontSize = fonts + 1
fonts = fonts + 1
angle = angle + 45
If angle >= 360 Then angle = 0
If fonts = 50 Then
Rotext1.angle = 0
Pause 1
Do: DoEvents
Rotext1.Left = Rotext1.Left + 20
Loop Until Rotext1.Left > Me.Left
Rotext1.FontSize = 0
fonts = 0
nextname = nextname + 1
End If
End Sub
