VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Example by Twirp"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2805
   Icon            =   "Example.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Test it"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Example.frx":0442
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Testing Label"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Code: See code of form for details"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By Twirp
'Email-TonyTwirp@hotmail.com
'AIM-Tmat989

'Ok here goes, this is saying With this form that we're on,
'Make a property for it called timeout which can be used
'with the form. Example is below, 'Form1.timeout=2'
'The timout doesn't help much but you can use other properties
'with it. You can incorporate this into control programming
'I would think, but i have never tried to make a control so
'I am not sure.
Property Let Timeout(Duration)
 'Normal timeout code
 Dim Now As Long
    Now = Timer
    Do Until Timer - Now >= Duration
        DoEvents
    Loop
'End property statement, i think. heh
End Property

Private Sub Command1_Click()
'Thats me
Label2.Caption = "Twirp is great!"
'Here i use the new property timeout before changing font to bold
Form1.Timeout = 2
Label2.FontBold = True
'Same as above but changes caption after timeout
Form1.Timeout = 2
Label2.Caption = "He is allright"
End Sub

'Conclusion: This is simple simple stuff.
'If you cant figure it out email TonyTwirp@hotmail.com

