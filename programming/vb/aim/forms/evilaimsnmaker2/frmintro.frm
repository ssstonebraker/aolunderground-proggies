VERSION 5.00
Begin VB.Form frmintro 
   BorderStyle     =   0  'None
   Caption         =   "Evil Aim Sn Maker 2 By Source"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   Icon            =   "frmintro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   0
      Picture         =   "frmintro.frx":030A
      ScaleHeight     =   3075
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1320
         Top             =   2280
      End
   End
End
Attribute VB_Name = "frmintro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'ontop sub from the module: aimsnmake.bas
OnTOP Me

'loadformstate sub from the module,selected frmmain to
'load the state of
LoadFormState frmmain

'loadformstate sub from the module,selected menu to
'load the state of
LoadFormState menu

'if check1(load intro?) isnt check then, it turns off
'the timer, and hides intro before it loads, then
'shows the main form (frmmain).

If menu.Check1.Value = 0 Then

Timer1.Enabled = False

frmintro.Hide

frmmain.Show

Else
'if its not off..its on !
Timer1.Enabled = True

End If

End Sub

Private Sub Picture1_Click()
'quicky bypass...
Timer1.Enabled = False
frmintro.Hide
frmmain.Show
End Sub

Private Sub Timer1_Timer()
If menu.Check1.Value = 0 Then Exit Sub

'calls 'pause' sub from module, and pauses for 1.2 seconds

Pause 1.5

'hides intro form
frmintro.Hide

'shows main form
frmmain.Show

'turns off timer so that it doesnt effect the rest
'of the run time
Timer1.Enabled = False
End Sub
