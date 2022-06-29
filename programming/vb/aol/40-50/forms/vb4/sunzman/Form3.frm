VERSION 4.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   600
   ClientLeft      =   2445
   ClientTop       =   1290
   ClientWidth     =   4710
   Height          =   1290
   Left            =   2385
   LinkTopic       =   "Form3"
   ScaleHeight     =   600
   ScaleWidth      =   4710
   Top             =   660
   Width           =   4830
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu info 
         Caption         =   "Info About Prog."
      End
      Begin VB.Menu Contact 
         Caption         =   "Contact PoOH"
      End
      Begin VB.Menu verfy 
         Caption         =   "Verrifier"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mini 
         Caption         =   "Minimize"
      End
      Begin VB.Menu ex 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu opt 
      Caption         =   "Options"
      Begin VB.Menu chat 
         Caption         =   "Chat Room Toolz"
      End
      Begin VB.Menu imt 
         Caption         =   "Instant Message Toolz"
      End
      Begin VB.Menu buster 
         Caption         =   "RooM Buster"
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu ma 
         Caption         =   "Mail"
         Begin VB.Menu coma 
            Caption         =   "Count Mail ( Mail Must be Open )"
         End
         Begin VB.Menu ma1 
            Caption         =   "Mail Me"
         End
      End
   End
   Begin VB.Menu adver 
      Caption         =   "Advertise"
      Begin VB.Menu one 
         Caption         =   "¹"
      End
      Begin VB.Menu two 
         Caption         =   "²"
      End
      Begin VB.Menu three 
         Caption         =   "³"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub ex_Click()
SendChat ("<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><b><i><font color= #000000>•SuNz OF M<html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><b><i><font color= #999933>aN • BY: PoO<html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><b><i><font color= #993300>H • UNLoaDeD•<html></html><html><html></html><html><html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html>")
X = UserSN()
MsgBox "Hey " & X & " , Itz PoOH here talkin to ya. I hope you enjoyed the Prog. If you didnt, well fuck all ya. But anywayz look out for my next prog!---Lata", vbExclamation, "C-Ya!"
End
End Sub

