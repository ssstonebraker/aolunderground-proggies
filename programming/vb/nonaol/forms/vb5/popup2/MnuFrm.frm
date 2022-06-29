VERSION 5.00
Begin VB.Form MnuFrm 
   Caption         =   "Form1"
   ClientHeight    =   195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   195
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu PopUpNotOnThisForm 
      Caption         =   "Popup Example"
      Begin VB.Menu ThisOne 
         Caption         =   "This Menu"
      End
      Begin VB.Menu isnot 
         Caption         =   "Is Not"
      End
      Begin VB.Menu onthe 
         Caption         =   "On The"
      End
      Begin VB.Menu mainform 
         Caption         =   "Main Form"
      End
   End
End
Attribute VB_Name = "MnuFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
