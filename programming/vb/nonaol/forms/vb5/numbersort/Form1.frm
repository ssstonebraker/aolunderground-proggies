VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Number Sort Example - By: EcCo"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Sort"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This example shows you how to sort your numbers in numerical order."
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Number Sort Example - By: EcCo
'E-Mail:  X EcCo X@hotmail.com

'This examples show you how you can sort your
'numbers in numerical order.  The program
'generates a hundred random number then allows
'you to sort them.

Option Explicit

Private Sub Command1_Click()
Dim Loo, Loo2, Loo3, TmpLoo As Integer
Dim Tmp, Tmp2 As Integer
Dim Count, LCount As Integer

ReDim ListArray(List1.ListCount - 1) As Integer
ReDim TmpListArray(List1.ListCount - 1) As Integer

Tmp = 0: Tmp2 = 0
Count = 0: LCount = List1.ListCount - 1

For TmpLoo = 0 To LCount
    ListArray(TmpLoo) = List1.List(TmpLoo)
Next TmpLoo

List1.Clear

For Loo = 0 To LCount
    If ListArray(Loo) > Tmp Then Tmp = ListArray(Loo)
Next Loo

For Loo2 = 0 To Tmp
    For Loo3 = 0 To LCount
        If ListArray(Loo3) = Loo2 Then
            Tmp2 = ListArray(Loo3)
            ListArray(Loo3) = 0
            TmpListArray(Count) = Tmp2
            Count = Count + 1
        End If
    Next Loo3
Next Loo2

For TmpLoo = 0 To LCount
    List1.AddItem TmpListArray(TmpLoo)
Next TmpLoo

End Sub

Private Sub Form_Load()
Dim Loo As Integer

Randomize Timer

For Loo = 0 To 100
List1.AddItem Int(1000 * Rnd)
Next Loo
End Sub
