VERSION 5.00
Begin VB.Form frmMulti 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Azazel - Multi Form - Team Crystal"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "About This"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make another form, the same as me!"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Azazel As New frmMulti
    Azazel.Show
    Azazel.Move Left + (Width \ 10), Top + (Height \ 10)
    Azazel.BackColor = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
End Sub


Private Sub Command2_Click()
MsgBox "I have found this good for something like a Macro Maker or some kind of werd-pad of some kind.  If you have any questions, IM me at ''Azazel 6X9'' or Email me at ''Azazel_666@juno.com''.  If you would like to be apart of Crystal 2000, which is a programming team I am on, then you can IM me about that also.  It's just a small team of ICP, MasTaDogg and I and we need maybe 2 more members.", vbInformation, "- Multi Form - Team Crystal"
End Sub


