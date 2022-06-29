VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu click 
      Caption         =   "click"
      Begin VB.Menu Restartaol5 
         Caption         =   "RestartAol5"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Restartaol5_Click()
ServerForm.List1.Clear 'This is weird cause its what i use for intergers, i'm sure u know another way of doin it
Do:

answer$ = InputBox("How many till restart? nothing stupid!", "4:20") 'Asks the user a question
Call WriteToINI("RestartAol", "RestartAol", answer$, App.Path & "\insain") 'Records the users answer
Form2.Restartaol5.Caption = "RestarAol (" & answer$ & ")" 'Changes the Caption in the pull down menu's
Loop Until IsNumeric(answer$) 'Loops until the user gives a number as an answer
m = 0
Do: 'this will make it easy to keep track of when to restart aol.  Atleast it does for me :-)
    ServerForm.List1.AddItem "sup"
    m = m + 1
Loop Until m = answer$
End Sub
