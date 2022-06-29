VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add a Player"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2160
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   2160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Name"
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Integer
Dim G As Boolean

If Option1.Value = True Then
For a = 1 To Form1.List1.ListCount
If thename = Form1.List1.List(a) Then
G = True
End If
Next a
For a = 1 To Form1.List3.ListCount
If thename = Form1.List3.List(a) Then
G = True
End If
Next a
If G = False Then
Form1.List1.AddItem (Text1)
Form1.List2.AddItem (0)
Chat1.ChatSend ("<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(" & Text1 & " added to " & Option1.Caption)
Unload Me
Form1.Visible = True
Else
MsgBox "That name is already there!", vbInformation, "Error"
Unload Me
Form1.Visible = True
End If
ElseIf Option2.Value = True Then
For a = 1 To Form1.List1.ListCount
If thename = Form1.List1.List(a) Then
G = True
End If
Next a
For a = 1 To Form1.List3.ListCount
If thename = Form1.List3.List(a) Then
G = True
End If
Next a
If G = False Then
Form1.List3.AddItem (Text1)
Form1.List4.AddItem (0)
Chat1.ChatSend "<Font Color=" & Chr(34) & "#000000" & Chr(34) & "><Font Face=" & Chr(34) & "Arial" & Chr(34) & ">º°`(" & Text1.Text & " added to " & Option2.Caption
Unload Me
Form1.Visible = True
Else
MsgBox "That name is already there!", vbInformation, "Error"
Unload Me
Form1.Visible = True
End If
End If
End Sub

Private Sub Form_Load()
Option1.Caption = Form1.Label6.Caption
Option2.Caption = Form1.Label7.Caption
Option1.Value = True
End Sub
