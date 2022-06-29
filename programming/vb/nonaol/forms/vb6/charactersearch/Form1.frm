VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Search"
   ClientHeight    =   2265
   ClientLeft      =   6855
   ClientTop       =   5970
   ClientWidth     =   2670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   2670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "M"
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Geox2000@hotmail.com"
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Character To Search For:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Type Here:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
'Geox2000@hotmail.com

Dim Said, Result As String
    'Defines The Variables
Said = "" + Text1.Text + ""
    'Said is the same thing as text1.text (you can use said if you don't want to use text1.text)

For Counter = 1 To Len(Text1.Text)
    'This will Loop as many times as the lenth of characters in text1.text
Result = Mid(Said, Counter, 1)
    'This makes Result equal to Mid(Said, Counter,1)
    'Without it having the Mid the program will read Character + next character and so on
    'For example if the string you wanted to search was "Geox" then the program will read it like "G" "Ge" "Geo" "Geox"
    'But We want it to read like "G" "e" "o" "x" that way it will search only one(1) character at a time.
    'Mid(Said,Counter,1) just means it's going to extract part of a string from Said(remember, Said is equal to text1.text)
    'Then the Counter means at what position of the string is it going to start (Counter will Change its value as many time as the lenth of text1.text)
    'The 1 at the end means how many characters to extract begining from Counter so if Counter's value was 3 it simply take 1 character in this case it will take character 4)
    'It's not hard but I can't find the words to say it. so if you need more help e-mail me at Geox2000@hotmail.com ;)
If LCase(Result) = LCase(Text2.Text) Then
    'Here is simply telling the program that If Result is equal to text2.text then...Character found! (the LCase is there so VB can read "G" the same as "g" and "g" the same as "G" ;)
Total = Total + Result
    'Here is simply adding any matching characters it found in the string (this is needed to inform the user how many matching characters in the string were found
End If
    'You know, its the end of the If statement
Next Counter
    'End of the For...Next Statement (i think) hehe
MsgBox "" & Len(Total) & " Matching Character(s) Found!", vbInformation, "Search Results"
    'When its done looping through every character in the textbox it will
    'Display how many matching characters were found!
    'It was easy, right? If not e-mail me at Geox2000@hotmail.com
    'I will gladly help you....
End Sub


Private Sub Command2_Click()
'Geox2000@Hotmail.com

Form2.Show
    'This will make form2 visible
Beep
    'This will make a beep sound!
End Sub

Private Sub Form_Load()
'Geox2000@hotmail.com

Call MsgBox("This Example Was Made By: -Geox- Contact Me At Geox2000@hotmail.com", vbInformation, "About '99")
    'Displays a msgbox. First the string thats going to display, then the style, then the tittle. Theres more but I only use these three (usually).

Call MsgBox("This Example Doesent Require Any Basic Files, Or Extra Controls.", vbInformation, "Requirements?")
    'Displays a msgbox. First the string thats going to display, then the style, then the tittle. Theres more but I only use these three (usually).

Text1.Text = "geox2000@hotmail.com"
    'This will make text1.text display my e-mail adrress ;)
End Sub


