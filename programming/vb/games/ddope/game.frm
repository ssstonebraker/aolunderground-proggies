VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "dealin' dope version 2 remake"
   ClientHeight    =   4264
   ClientLeft      =   130
   ClientTop       =   559
   ClientWidth     =   5434
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4264
   ScaleWidth      =   5434
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   2691
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "  Matt"
      ToolTipText     =   "Double Click to change name"
      Top             =   3978
      Width           =   2743
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   " Beginner Dealer"
      ToolTipText     =   "Beginner Mode"
      Top             =   3978
      Width           =   2704
   End
   Begin VB.Frame Frame9 
      Height          =   715
      Left            =   60
      TabIndex        =   38
      Top             =   3159
      Width           =   5300
      Begin VB.PictureBox Picture1 
         Height          =   247
         Left            =   1638
         ScaleHeight     =   195
         ScaleWidth      =   2067
         TabIndex        =   39
         Top             =   280
         Width           =   2119
      End
      Begin VB.Label status 
         Caption         =   "Begin Trade"
         Height          =   247
         Left            =   117
         TabIndex        =   44
         Top             =   351
         Width           =   1300
      End
      Begin VB.Label days 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.54
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   247
         Left            =   4563
         TabIndex        =   43
         Top             =   351
         Width           =   598
      End
      Begin VB.Label Label21 
         Caption         =   "Day"
         Height          =   247
         Left            =   4095
         TabIndex        =   42
         Top             =   351
         Width           =   364
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Weapons"
      Height          =   832
      Left            =   3800
      TabIndex        =   23
      Top             =   2223
      Width           =   1580
      Begin VB.ListBox List2 
         Height          =   559
         Left            =   65
         TabIndex        =   35
         Top             =   234
         Width           =   1417
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Score"
      Height          =   598
      Left            =   3800
      TabIndex        =   22
      Top             =   1521
      Width           =   1580
      Begin VB.Label score 
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.97
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   247
         Left            =   351
         TabIndex        =   34
         Top             =   234
         Width           =   1183
      End
      Begin VB.Label Label17 
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.97
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   247
         Left            =   117
         TabIndex        =   33
         Top             =   234
         Width           =   247
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Status"
      Height          =   1534
      Left            =   1638
      TabIndex        =   21
      Top             =   1521
      Width           =   2119
      Begin VB.CommandButton Command10 
         Caption         =   "Shoot"
         Height          =   280
         Left            =   1090
         TabIndex        =   37
         Top             =   1170
         Width           =   949
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Punch"
         Height          =   280
         Left            =   70
         TabIndex        =   36
         Top             =   1170
         Width           =   949
      End
      Begin VB.Label about 
         Caption         =   "Choose a place to goto on the left then choose to buy or sell any of your drugs or attempt to harm them."
         Height          =   832
         Left            =   117
         TabIndex        =   45
         Top             =   234
         Width           =   1885
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Location"
      Height          =   1534
      Left            =   60
      TabIndex        =   19
      Top             =   1521
      Width           =   1534
      Begin VB.ListBox List1 
         Height          =   1235
         Left            =   65
         TabIndex        =   20
         Top             =   234
         Width           =   1417
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Herion"
      Height          =   1300
      Left            =   4095
      TabIndex        =   3
      Top             =   117
      Width           =   1300
      Begin VB.CommandButton herionsell 
         Caption         =   "Sell"
         Height          =   247
         Left            =   702
         TabIndex        =   18
         Top             =   936
         Width           =   481
      End
      Begin VB.CommandButton herionbuy 
         Caption         =   "Buy"
         Height          =   247
         Left            =   117
         TabIndex        =   17
         Top             =   936
         Width           =   481
      End
      Begin VB.Label Label16 
         Caption         =   "Price:"
         Height          =   247
         Left            =   117
         TabIndex        =   32
         Top             =   270
         Width           =   481
      End
      Begin VB.Label Label14 
         Caption         =   "Supply:"
         Height          =   247
         Left            =   117
         TabIndex        =   30
         Top             =   610
         Width           =   598
      End
      Begin VB.Label herions 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.97
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   247
         Left            =   702
         TabIndex        =   26
         Top             =   585
         Width           =   481
      End
      Begin VB.Label herionp 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.97
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   247
         Left            =   702
         TabIndex        =   8
         Top             =   234
         Width           =   481
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cocain"
      Height          =   1300
      Left            =   2760
      TabIndex        =   2
      Top             =   117
      Width           =   1300
      Begin VB.CommandButton cocainbuy 
         Caption         =   "Buy"
         Height          =   247
         Left            =   117
         TabIndex        =   16
         Top             =   936
         Width           =   481
      End
      Begin VB.CommandButton cocainsell 
         Caption         =   "Sell"
         Height          =   247
         Left            =   702
         TabIndex        =   15
         Top             =   936
         Width           =   481
      End
      Begin VB.Label Label15 
         Caption         =   "Price:"
         Height          =   247
         Left            =   117
         TabIndex        =   31
         Top             =   270
         Width           =   481
      End
      Begin VB.Label Label13 
         Caption         =   "Supply:"
         Height          =   247
         Left            =   117
         TabIndex        =   29
         Top             =   610
         Width           =   598
      End
      Begin VB.Label cocains 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.97
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   247
         Left            =   702
         TabIndex        =   25
         Top             =   585
         Width           =   481
      End
      Begin VB.Label cocainp 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.97
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   247
         Left            =   702
         TabIndex        =   7
         Top             =   234
         Width           =   481
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Acid"
      Height          =   1300
      Left            =   1404
      TabIndex        =   1
      Top             =   117
      Width           =   1300
      Begin VB.CommandButton acidsell 
         Caption         =   "Sell"
         Height          =   247
         Left            =   702
         TabIndex        =   14
         Top             =   936
         Width           =   481
      End
      Begin VB.CommandButton acidbuy 
         Caption         =   "Buy"
         Height          =   247
         Left            =   117
         TabIndex        =   13
         Top             =   936
         Width           =   481
      End
      Begin VB.Label Label12 
         Caption         =   "Supply:"
         Height          =   247
         Left            =   117
         TabIndex        =   28
         Top             =   610
         Width           =   598
      End
      Begin VB.Label Label11 
         Caption         =   "Price:"
         Height          =   247
         Left            =   117
         TabIndex        =   27
         Top             =   270
         Width           =   481
      End
      Begin VB.Label acids 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.97
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   247
         Left            =   702
         TabIndex        =   24
         Top             =   585
         Width           =   481
      End
      Begin VB.Label acidp 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.97
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   247
         Left            =   702
         TabIndex        =   6
         Top             =   234
         Width           =   481
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Weed"
      Height          =   1300
      Left            =   60
      TabIndex        =   0
      Top             =   117
      Width           =   1300
      Begin VB.CommandButton weedsell 
         Caption         =   "Sell"
         Height          =   247
         Left            =   702
         TabIndex        =   12
         Top             =   936
         Width           =   481
      End
      Begin VB.CommandButton weedbuy 
         Caption         =   "Buy"
         Height          =   247
         Left            =   117
         TabIndex        =   11
         Top             =   936
         Width           =   481
      End
      Begin VB.Label Label7 
         Caption         =   "Supply:"
         Height          =   247
         Left            =   117
         TabIndex        =   10
         Top             =   610
         Width           =   598
      End
      Begin VB.Label weeds 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.97
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   247
         Left            =   702
         TabIndex        =   9
         Top             =   585
         Width           =   481
      End
      Begin VB.Label weedp 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.97
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   247
         Left            =   702
         TabIndex        =   5
         Top             =   234
         Width           =   481
      End
      Begin VB.Label Label1 
         Caption         =   "Price:"
         Height          =   247
         Left            =   117
         TabIndex        =   4
         Top             =   270
         Width           =   598
      End
   End
   Begin VB.Label Label20 
      Caption         =   "30"
      Height          =   247
      Left            =   2340
      TabIndex        =   41
      Top             =   4797
      Width           =   1300
   End
   Begin VB.Label Label19 
      Caption         =   "15"
      Height          =   247
      Left            =   702
      TabIndex        =   40
      Top             =   4797
      Width           =   1300
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu gfdgd 
         Caption         =   "-"
      End
      Begin VB.Menu sgame 
         Caption         =   "Save Game"
      End
      Begin VB.Menu lgame 
         Caption         =   "Load Game"
      End
      Begin VB.Menu fff 
         Caption         =   "-"
      End
      Begin VB.Menu ngame 
         Caption         =   "New Game"
      End
      Begin VB.Menu fdfd 
         Caption         =   "-"
      End
      Begin VB.Menu exitt 
         Caption         =   "Exit"
      End
      Begin VB.Menu min 
         Caption         =   "Minimize"
      End
      Begin VB.Menu sdf 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub acidbuy_Click()
If Val(acidp.Caption) > Val(score.Caption) Then
MsgBox "not enough money to buy"
Else
score.Caption = score.Caption - acidp.Caption
acids.Caption = acids.Caption + 1
End If
End Sub

Private Sub acidsell_Click()
If acids.Caption = "0" Then
FormNotOnTop Me
MsgBox "Out of acid you bum!"
FormOnTop Me
Else
acids.Caption = acids.Caption - 1
score.Caption = Val(score.Caption) + Val(acidp.Caption)
End If
End Sub

Private Sub cocainbuy_Click()
If Val(cocainp.Caption) > Val(score.Caption) Then
MsgBox "not enough money to buy"
Else
score.Caption = score.Caption - cocainp.Caption
cocains.Caption = cocains.Caption + 1
End If
End Sub

Private Sub cocainsell_Click()
If cocains.Caption = "0" Then
FormNotOnTop Me
MsgBox "You Crack Whore, you have no crack!"
FormOnTop Me
Else
cocains.Caption = cocains.Caption - 1
score.Caption = Val(score.Caption) + Val(cocainp.Caption)
End If
End Sub

Private Sub Command10_Click()
status.Caption = "You busted a cap in his ass!"
End Sub

Private Sub Command9_Click()
Dim X As Integer
X = Int(Rnd * 9)
Select Case X
Case 0
about.Caption = "You beat the shit out of em' and jacked a ak-47"
List2.AddItem "ak-47"
Command9.Enabled = False
Case 1
about.Caption = "You just got your ass kicked and lost all your guns!"
List2.Clear
Command9.Enabled = False
Case 2
about.Caption = "You picked up a 9 millimeter"
List2.AddItem "9 millimeter"
Command9.Enabled = False
Case 3
about.Caption = "You pussy, he kicked your ass."
Command9.Enabled = False
Case 4
about.Caption = "You lost all your money, all you have now is 50 that was in your sock."
score.Caption = "50"
Command9.Enabled = False
Case 5
about.Caption = "You kicked his ass, and jacked a tech-9"
List2.AddItem "tech-9"
Command9.Enabled = False
Case 6
about.Caption = "You got butt raped and lost all your guns."
List2.Clear
Command9.Enabled = False

Case 7
about.Caption = "You punched him and stole his weed!"
weeds.Caption = weeds.Caption + 1
Command9.Enabled = False
Case 8
about.Caption = "Dumbass, that was a cop, you lost all your drugs."
weeds.Caption = "0"
acids.Caption = "0"
cocains.Caption = "0"
herions.Caption = "0"
Command9.Enabled = False
End Select
End Sub

Private Sub exitt_Click()
Unload Me
End

End Sub

Private Sub Form_Load()
Form1.Show
'reset score
score.Caption = "50"
'reset weed price/supply
weedp.Caption = "0"
weeds.Caption = "0"
'reset acid price/supply
acidp.Caption = "0"
acids.Caption = "0"
'reset cocain price/supply
cocainp.Caption = "0"
cocains.Caption = "0"
'reset herion price/supply
herionp.Caption = "0"
herions.Caption = "0"
'new user reset
Dim handle As String
handle = InputBox("Enter your name.", "handle")
Text2.Text = handle
Text1.Text = "Beginner Dealer"
'add locations
List1.AddItem ("Africa")
List1.AddItem ("Amsterdam")
List1.AddItem ("Birmingham")
List1.AddItem ("Chicago")
List1.AddItem ("China")
List1.AddItem ("Colombia")
List1.AddItem ("Dallas")
List1.AddItem ("Denver")
List1.AddItem ("Detroit City")
List1.AddItem ("Downtown L.A.")
List1.AddItem ("Hawaii")
List1.AddItem ("Jamacia")
List1.AddItem ("Japan")
List1.AddItem ("Manhatten")
List1.AddItem ("Miami")
List1.AddItem ("New Orleans")
List1.AddItem ("Romulas")
List1.AddItem ("The Bronx")
FormOnTop Me

End Sub

Private Sub kmot_Click()
FormOnTop Form1

End Sub

Private Sub herionbuy_Click()
If Val(herionp.Caption) > Val(score.Caption) Then
MsgBox "not enough money to buy"
Else
score.Caption = score.Caption - herionp.Caption
herions.Caption = herions.Caption + 1
End If
End Sub

Private Sub herionsell_Click()
If herions.Caption = "0" Then
FormNotOnTop Me
MsgBox "Stop smoking all that crack nigga, you dont have any more!"
FormOnTop Me
Else
herions.Caption = herions.Caption - 1
score.Caption = Val(score.Caption) + Val(herionp.Caption)
End If
End Sub

Private Sub lgame_Click()
FormNotOnTop Me
MsgBox "Not finished yet.."
FormOnTop Me
End Sub

Private Sub List1_Click()

If List1 = List1.List(0) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase
Dim X As Integer
X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Rosanne from Africa: What the fuck you looking at??"
Case 1
about.Caption = "Kathy from Africa: You shouldnt have came."
Case 2
about.Caption = "A negro from Africa: Buy my weed while you can man!"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "48"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "124"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "162"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True
Else

If List1 = List1.List(1) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Kid Rock from Amsterdam: This city is known for it's high quality drugs."
Case 1
about.Caption = "Justen from Amsterdam: Good time to sell, not to buy."
Case 2
about.Caption = "Paco from Amsterdam: You shouldnt have came."
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True

Else




If List1 = List1.List(2) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Rosanne from Birmingham: We got the best dank around."
Case 1
about.Caption = "Carla from Birmingham: Buy all the shit you can!"
Case 2
about.Caption = "A negro from Birmingham: Can you spare some weed man?"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True
Else
If List1 = List1.List(3) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Home Boy from Chicago: y0 nigga, got a light?"
Case 1
about.Caption = "Kathy from Chicago: You shouldnt have came."
Case 2
about.Caption = "Billy from Chicago: Theres not much out here."
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True
Else
If List1 = List1.List(4) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Angel from China: Cmon man, it's not that much."
Case 1
about.Caption = "Jane from China: Can you give me some for free?"
Case 2
about.Caption = "Bob from China: Buy my weed while you can man!"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True

Else
If List1 = List1.List(5) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Rosanne from Colombia: Colombia is the best place to get your goods."
Case 1
about.Caption = "Jim from Colombia: Would you like to smoke down?"
Case 2
about.Caption = "Rica from Colombia: Good time to see, not to buy"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True

Else
If List1 = List1.List(6) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Kathy from Dallas: I need some weed can you hook me up?"
Case 1
about.Caption = "Danie from Dallas: We just got busted so shit is low."
Case 2
about.Caption = "Mary from Dallas: What do you say we go back to my place."
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True

Else
If List1 = List1.List(7) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Oj from Denver: Can you hook me up?  This is a good place to sell."
Case 1
about.Caption = "Jesus from Denver: You shouldnt have came."
Case 2
about.Caption = "A negro from Denver: I got an ak-47 so back-off!!!!"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True
Else
If List1 = List1.List(8) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Gary from Detroit City: What the fuck you looking at??"
Case 1
about.Caption = "Kathy from Detroit City: You shouldnt have came."
Case 2
about.Caption = "Lee from Detroit City: Can you spare some dank man?"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True
Else
If List1 = List1.List(9) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Bill from Downtown L.A.: Our shit is the best!"
Case 1
about.Caption = "Joe from Downtown L.A.: Who the fuck are you?"
Case 2
about.Caption = "Bob from Downtown L.A.: Cmon in man!"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True
Else
If List1 = List1.List(10) Then
Command9.Enabled = True
'don't mess up the game!!!
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "G from Hawaii: Can you give me some for free?"
Case 1
about.Caption = "Brian from Hawaii: Who the fuck are you?"
Case 2
about.Caption = "Buddah from Hawaii: I need some coke! Gimme it!!"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True
Else
If List1 = List1.List(11) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Mary from Jamacia: Buy all the shit you can!!"
Case 1
about.Caption = "G from Jamacia: Would you like to smoke down?"
Case 2
about.Caption = "Blair from Jamacia: Our prices bottomed out!"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True

Else
If List1 = List1.List(12) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "A Spec from Japan: I will suck your dick for weed."
Case 1
about.Caption = "Your mom from Japan: Damn, you havin fun?"
Case 2
about.Caption = "A Hick from Japan: Are you a Cop?"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True
Else
If List1 = List1.List(13) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Your mom from Manhatten: I'm fucked up right now!"
Case 1
about.Caption = "Macro from Manhatten: There aint much out here."
Case 2
about.Caption = "Justen from Manhatten: We got the best dank around!"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True
Else
If List1 = List1.List(14) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "O.J from Miami: I got an ak-47 so back-off!"
Case 1
about.Caption = "Jo from Miami: Good time to buy, not to sell!"
Case 2
about.Caption = "The Police from Miami: We got the best dank around!"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True
Else
If List1 = List1.List(15) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Celsie from New Orleans: Damn you have alot!!"
Case 1
about.Caption = "Kathy from New Orleans: You shouldnt have came."
Case 2
about.Caption = "Mike from New Orleans: You havin' fun?"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True

Else
If List1 = List1.List(16) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "Matt from Romulas: Do you have any smokes?"
Case 1
about.Caption = "Rhino from Romulas: The Prices are crazy!"
Case 2
about.Caption = "Ravage from Romulas: Get the fuck out of here!"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True

Else
If List1 = List1.List(17) Then
Command9.Enabled = True
'disable all commands so they can't mess game up
weedbuy.Enabled = False
weedsell.Enabled = False
acidbuy.Enabled = False
acidsell.Enabled = False
cocainbuy.Enabled = False
cocainsell.Enabled = False
herionbuy.Enabled = False
herionsell.Enabled = False
List1.Enabled = False
'random phrase

X = Int(Rnd * 3)
Select Case X
Case 0
about.Caption = "TJ from The Bronx: I have a tech-9 do back off!!!"
Case 1
about.Caption = "Dmx from The Bronx: Yo homes Cmon in."
Case 2
about.Caption = "A Bitch from The Bronx: Our shit is the best man!"
End Select
status.Caption = "Packing....."
Call PercentBar(Picture1, "5", "15")
Pause 0.6
status.Caption = "Going to airport...."
Call PercentBar(Picture1, "8", "15")
Pause 0.8
status.Caption = "Flying....."
Call PercentBar(Picture1, "10", "15")
Pause 0.7
status.Caption = "Checking in....."
Call PercentBar(Picture1, "15", "15")
Pause 0.9
Call PercentBar(Picture1, "0", "15")
days.Caption = days.Caption + 1
status.Caption = "Begin Trade...."
X = Int(Rnd * 15)
Select Case X
Case 0
weedp.Caption = "11"
Case 1
weedp.Caption = "12"
Case 2
weedp.Caption = "15"
Case 3
weedp.Caption = "18"
Case 4
weedp.Caption = "21"
Case 5
weedp.Caption = "24"
Case 6
weedp.Caption = "27"
Case 7
weedp.Caption = "30"
Case 8
weedp.Caption = "33"
Case 9
weedp.Caption = "36"
Case 10
weedp.Caption = "39"
Case 11
weedp.Caption = "42"
Case 12
weedp.Caption = "45"
Case 13
weedp.Caption = "48"
Case 14
weedp.Caption = "50"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
acidp.Caption = "90"
Case 1
acidp.Caption = "125"
Case 2
acidp.Caption = "112"
Case 3
acidp.Caption = "180"
Case 4
acidp.Caption = "142"
Case 5
acidp.Caption = "102"
Case 6
acidp.Caption = "101"
Case 7
acidp.Caption = "111"
Case 8
acidp.Caption = "133"
Case 9
acidp.Caption = "136"
Case 10
acidp.Caption = "129"
Case 11
acidp.Caption = "142"
Case 12
acidp.Caption = "175"
Case 13
acidp.Caption = "118"
Case 14
acidp.Caption = "151"
End Select


X = Int(Rnd * 15)
Select Case X
Case 0
cocainp.Caption = "68"
Case 1
cocainp.Caption = "122"
Case 2
cocainp.Caption = "79"
Case 3
cocainp.Caption = "180"
Case 4
cocainp.Caption = "145"
Case 5
cocainp.Caption = "124"
Case 6
cocainp.Caption = "127"
Case 7
cocainp.Caption = "130"
Case 8
cocainp.Caption = "163"
Case 9
cocainp.Caption = "146"
Case 10
cocainp.Caption = "97"
Case 11
cocainp.Caption = "101"
Case 12
cocainp.Caption = "114"
Case 13
cocainp.Caption = "118"
Case 14
cocainp.Caption = "136"
End Select

X = Int(Rnd * 15)
Select Case X
Case 0
herionp.Caption = "168"
Case 1
herionp.Caption = "122"
Case 2
herionp.Caption = "179"
Case 3
herionp.Caption = "180"
Case 4
herionp.Caption = "145"
Case 5
herionp.Caption = "138"
Case 6
herionp.Caption = "127"
Case 7
herionp.Caption = "130"
Case 8
herionp.Caption = "163"
Case 9
herionp.Caption = "139"
Case 10
herionp.Caption = "97"
Case 11
herionp.Caption = "101"
Case 12
herionp.Caption = "114"
Case 13
herionp.Caption = "200"
Case 14
herionp.Caption = "157"
End Select


weedbuy.Enabled = True
weedsell.Enabled = True
acidbuy.Enabled = True
acidsell.Enabled = True
cocainbuy.Enabled = True
cocainsell.Enabled = True
herionbuy.Enabled = True
herionsell.Enabled = True
List1.Enabled = True

Else

End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If

End Sub

Private Sub min_Click()
Form1.WindowState = 1

End Sub

Private Sub ngame_Click()
FormNotOnTop Me
'reset score
score.Caption = "50"
'reset weed price/supply
weedp.Caption = "0"
weeds.Caption = "0"
'reset acid price/supply
acidp.Caption = "0"
acids.Caption = "0"
'reset cocain price/supply
cocainp.Caption = "0"
cocains.Caption = "0"
'reset herion price/supply
herionp.Caption = "0"
herions.Caption = "0"
'new user reset
Dim handle As String
handle = InputBox("Enter your name.", "handle")
Text2.Text = handle
Text1.Text = "Beginner Dealer"
FormOnTop Me
End Sub

Private Sub sgame_Click()
FormNotOnTop Me
MsgBox "Not finished yet.."
FormOnTop Me
End Sub

Private Sub weedbuy_Click()
If Val(weedp.Caption) > Val(score.Caption) Then
MsgBox "not enough money to buy"
Else
score.Caption = score.Caption - weedp.Caption
weeds.Caption = weeds.Caption + 1
End If

End Sub

Private Sub weedsell_Click()
If weeds.Caption = "0" Then
FormNotOnTop Me
MsgBox "Stop smoking all that crack nigga, you dont have any more!"
FormOnTop Me
Else
weeds.Caption = weeds.Caption - 1
score.Caption = Val(score.Caption) + Val(weedp.Caption)
End If
End Sub
