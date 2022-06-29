VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "average"
   ClientHeight    =   600
   ClientLeft      =   3015
   ClientTop       =   4140
   ClientWidth     =   2130
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   1005
   Left            =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   2130
   Top             =   3795
   Width           =   2250
   Begin VB.CommandButton Command3 
      Caption         =   "exit"
      Height          =   210
      Left            =   1635
      TabIndex        =   4
      Top             =   210
      Width           =   420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "clear"
      Height          =   210
      Left            =   1230
      TabIndex        =   3
      Top             =   210
      Width           =   420
   End
   Begin VB.CommandButton Command2 
      Caption         =   "calculate"
      Height          =   210
      Left            =   1230
      TabIndex        =   1
      Top             =   0
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "sum: 0"
      Height          =   285
      Left            =   30
      TabIndex        =   5
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "numbers: 0"
      Height          =   210
      Left            =   45
      TabIndex        =   2
      Top             =   270
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Caption = "numbers: 0"
Label2.Caption = "sum: 0"
Text1.Text = ""
End Sub

Private Sub Command2_Click()
'Iz representing eXcel 2001

kounter = 0 ' uses 'kounter' as an intiger to count numbers(any word would work)
Sum = 0 ' stores the sum of the numbers (addition) in a variable named sum
Do Until x < 0 'goes until you cancel
On Error GoTo done 'this is an error trap, i love'em so it helps your subs from erroring
'it simple tells vb if there an error go to another part of the code
x = InputBox("enter the numbers you would like to average, press cancel when your done.")
'this is the code for an inputbox...pretty simple, eh?
If x = 0 Then Exit Do '0 is what vb returns when you cancle a inputbox...so basically
'if you press cancle it'll stop askin for numbers :)
Sum = Sum + Val(x) ' remember sum? this is how it adds up the sum :)
kounter = kounter + 1 ' counts the numbers you put in
Loop ' loop, nessicary to send it back to the Do
done: ' done here is your error trap...if the program had errored /\ up there it would bring it down to here
On Error GoTo done2 ' another error trap
Text1.Text = Sum / (kounter) ' divides the sum of all the numbers by how many numbers there were
'cuz thats how you get the average
Label1.Caption = "numbers: " & kounter 'displays the amount of numbers in a label
Label2.Caption = "sum: " & Sum 'displays the sum of all the numbers in a label
done2: ' the end of your second error trap
End Sub


Private Sub Command3_Click()
End
End Sub


