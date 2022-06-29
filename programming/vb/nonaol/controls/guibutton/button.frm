VERSION 5.00
Object = "{2FE877A9-0482-11D3-B240-44455354616F}#40.0#0"; "BUTTON.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   1950
   StartUpPosition =   3  'Windows Default
   Begin GUIButton.Button Button1 
      Height          =   555
      Left            =   405
      TabIndex        =   0
      Top             =   120
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   979
      Pic_down        =   "button.frx":0000
      Pic_over        =   "button.frx":1EF8
      Pic_off         =   "button.frx":3DF0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Button1_Click()
Button1.AfterClick
'this sub is for when the mouse isn't
'over the button and my OCX doesn't read
'it going off. Its really only for 1 reason.
'Message box's! They mess up the button a little
End Sub

Private Sub Button1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'mouse button down
End Sub

Private Sub Button1_MouseOVER(Button As Integer, Shift As Integer, X As Single, Y As Single)
'when mouse is over
End Sub

Private Sub Button1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'when mouse button is up
End Sub

Private Sub Form_Load()
Button1.Pic_off = LoadPicture("C:\Button1.jpg") 'now this is ONLY an exmple of how to use this
Button1.Pic_over = LoadPicture("C:\Button2.jpg") 'basicly you get three pictures...1 for when the mouse
Button1.Pic_down = LoadPicture("C:\Button3.jpg") 'ins't on your button, when it is, and when the mouse is pushing down
'on your button.....Thats all!
End Sub
