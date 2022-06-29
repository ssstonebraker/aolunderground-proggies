VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   0  'None
   Caption         =   "Evil Aim Sn Make 2 By Source"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Main 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   0
      Picture         =   "frmmain.frx":0000
      ScaleHeight     =   4635
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   0
      Width           =   4725
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   1470
         Left            =   2400
         TabIndex        =   4
         Top             =   2520
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1470
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtpw 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtsn 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   1
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label list2count 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "list count"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3480
         TabIndex        =   16
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label list1count 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "list count"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblstatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "status bar..."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   4080
         Width           =   3975
      End
      Begin VB.Label lblmake 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3240
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblmore 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3600
         MousePointer    =   10  'Up Arrow
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbllistoptions 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2040
         MousePointer    =   10  'Up Arrow
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblsettings 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   840
         MousePointer    =   10  'Up Arrow
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblfile 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         MousePointer    =   10  'Up Arrow
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblx 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   4080
         MousePointer    =   2  'Cross
         TabIndex        =   7
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblminimize 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   3720
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lbldragform 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   120
         MousePointer    =   5  'Size
         TabIndex        =   5
         Top             =   120
         Width           =   3375
      End
   End
   Begin InetCtlsObjects.Inet AIM 
      Left            =   4920
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim listItemsVisible As Long 'dim
Private Type RECT 'sets rect for later use by
    Left As Long  'defining left, top, right, bottom
    Top As Long
    Right As Long
    Bottom As Long
'ends it
End Type

'PRIVATE declares, used in inet control
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'private const used in Form_Load
Private Const LB_GETITEMHEIGHT = &H1A1

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Call adv

'loadformstate sub from module, selects: frmmain
LoadFormState frmmain

'loadformstate sub from module, selects: menu
LoadFormState menu

'makes frmmain on top of other windows
OnTOP Me

'addhscroll from module, if the list is big enough
'it will had horizontal scroll bar to list1
Call AddHScroll(List1)

'addhscroll from module, if the list is big enough
'it will had horizontal scroll bar to list2
Call AddHScroll(List2)

'loadsformstate from module, for frmmain
LoadFormState frmmain

'loadsformstate from module, for menu
LoadFormState menu
 
 'dim returnvalue
    Dim returnvalue As Long
 'dim myrect
    Dim myrect As RECT
    'get height of listbox in pixels
    'myrect type receives the window dimensions
    returnvalue = GetClientRect(List1.hwnd, myrect)
    itemHeight = SendMessage(List1.hwnd, LB_GETITEMHEIGHT, 0&, 0&)

   'calculate number of visible items (not listcount)
    listItemsVisible = myrect.Bottom \ itemHeight

End Sub
Public Function GenerateEmail()
'100% not by source...i reviewed this code...
'basically what it does is randomize valid
'length of characters for a e-mail addy, then it
'processes them and randomizes the output.
'this is used in the making code, because since
'you can be track0rd by e-mail it randomizes some
'trash and you dont have to enter an e-mail addy.
Dim MyValue
Randomize
GenerateEmail = ""
MyValue = Int((26 * Rnd) + 1)   ' Generate random
If MyValue = 1 Then GenerateEmail = GenerateEmail + "a"
If MyValue = 2 Then GenerateEmail = GenerateEmail + "b"
If MyValue = 3 Then GenerateEmail = GenerateEmail + "c"
If MyValue = 4 Then GenerateEmail = GenerateEmail + "d"
If MyValue = 5 Then GenerateEmail = GenerateEmail + "e"
If MyValue = 6 Then GenerateEmail = GenerateEmail + "f"
If MyValue = 7 Then GenerateEmail = GenerateEmail + "g"
If MyValue = 8 Then GenerateEmail = GenerateEmail + "h"
If MyValue = 9 Then GenerateEmail = GenerateEmail + "i"
If MyValue = 10 Then GenerateEmail = GenerateEmail + "j"
If MyValue = 11 Then GenerateEmail = GenerateEmail + "k"
If MyValue = 12 Then GenerateEmail = GenerateEmail + "l"
If MyValue = 13 Then GenerateEmail = GenerateEmail + "m"
If MyValue = 14 Then GenerateEmail = GenerateEmail + "n"
If MyValue = 15 Then GenerateEmail = GenerateEmail + "o"
If MyValue = 16 Then GenerateEmail = GenerateEmail + "p"
If MyValue = 17 Then GenerateEmail = GenerateEmail + "q"
If MyValue = 18 Then GenerateEmail = GenerateEmail + "r"
If MyValue = 19 Then GenerateEmail = GenerateEmail + "s"
If MyValue = 20 Then GenerateEmail = GenerateEmail + "t"
If MyValue = 21 Then GenerateEmail = GenerateEmail + "u"
If MyValue = 22 Then GenerateEmail = GenerateEmail + "v"
If MyValue = 23 Then GenerateEmail = GenerateEmail + "w"
If MyValue = 24 Then GenerateEmail = GenerateEmail + "x"
If MyValue = 25 Then GenerateEmail = GenerateEmail + "y"
If MyValue = 26 Then GenerateEmail = GenerateEmail + "z"
MyValue = Int((26 * Rnd) + 1)   ' Generate random
If MyValue = 1 Then GenerateEmail = GenerateEmail + "a"
If MyValue = 2 Then GenerateEmail = GenerateEmail + "b"
If MyValue = 3 Then GenerateEmail = GenerateEmail + "c"
If MyValue = 4 Then GenerateEmail = GenerateEmail + "d"
If MyValue = 5 Then GenerateEmail = GenerateEmail + "e"
If MyValue = 6 Then GenerateEmail = GenerateEmail + "f"
If MyValue = 7 Then GenerateEmail = GenerateEmail + "g"
If MyValue = 8 Then GenerateEmail = GenerateEmail + "h"
If MyValue = 9 Then GenerateEmail = GenerateEmail + "i"
If MyValue = 10 Then GenerateEmail = GenerateEmail + "j"
If MyValue = 11 Then GenerateEmail = GenerateEmail + "k"
If MyValue = 12 Then GenerateEmail = GenerateEmail + "l"
If MyValue = 13 Then GenerateEmail = GenerateEmail + "m"
If MyValue = 14 Then GenerateEmail = GenerateEmail + "n"
If MyValue = 15 Then GenerateEmail = GenerateEmail + "o"
If MyValue = 16 Then GenerateEmail = GenerateEmail + "p"
If MyValue = 17 Then GenerateEmail = GenerateEmail + "q"
If MyValue = 18 Then GenerateEmail = GenerateEmail + "r"
If MyValue = 19 Then GenerateEmail = GenerateEmail + "s"
If MyValue = 20 Then GenerateEmail = GenerateEmail + "t"
If MyValue = 21 Then GenerateEmail = GenerateEmail + "u"
If MyValue = 22 Then GenerateEmail = GenerateEmail + "v"
If MyValue = 23 Then GenerateEmail = GenerateEmail + "w"
If MyValue = 24 Then GenerateEmail = GenerateEmail + "x"
If MyValue = 25 Then GenerateEmail = GenerateEmail + "y"
If MyValue = 26 Then GenerateEmail = GenerateEmail + "z"
MyValue = Int((26 * Rnd) + 1)   ' Generate random
If MyValue = 1 Then GenerateEmail = GenerateEmail + "a"
If MyValue = 2 Then GenerateEmail = GenerateEmail + "b"
If MyValue = 3 Then GenerateEmail = GenerateEmail + "c"
If MyValue = 4 Then GenerateEmail = GenerateEmail + "d"
If MyValue = 5 Then GenerateEmail = GenerateEmail + "e"
If MyValue = 6 Then GenerateEmail = GenerateEmail + "f"
If MyValue = 7 Then GenerateEmail = GenerateEmail + "g"
If MyValue = 8 Then GenerateEmail = GenerateEmail + "h"
If MyValue = 9 Then GenerateEmail = GenerateEmail + "i"
If MyValue = 10 Then GenerateEmail = GenerateEmail + "j"
If MyValue = 11 Then GenerateEmail = GenerateEmail + "k"
If MyValue = 12 Then GenerateEmail = GenerateEmail + "l"
If MyValue = 13 Then GenerateEmail = GenerateEmail + "m"
If MyValue = 14 Then GenerateEmail = GenerateEmail + "n"
If MyValue = 15 Then GenerateEmail = GenerateEmail + "o"
If MyValue = 16 Then GenerateEmail = GenerateEmail + "p"
If MyValue = 17 Then GenerateEmail = GenerateEmail + "q"
If MyValue = 18 Then GenerateEmail = GenerateEmail + "r"
If MyValue = 19 Then GenerateEmail = GenerateEmail + "s"
If MyValue = 20 Then GenerateEmail = GenerateEmail + "t"
If MyValue = 21 Then GenerateEmail = GenerateEmail + "u"
If MyValue = 22 Then GenerateEmail = GenerateEmail + "v"
If MyValue = 23 Then GenerateEmail = GenerateEmail + "w"
If MyValue = 24 Then GenerateEmail = GenerateEmail + "x"
If MyValue = 25 Then GenerateEmail = GenerateEmail + "y"
If MyValue = 26 Then GenerateEmail = GenerateEmail + "z"
MyValue = Int((1000 * Rnd) + 1)   ' Generate random
GenerateEmail = GenerateEmail & MyValue
End Function


Private Sub lbldragform_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'when lbldragform is pressed down, calls sub to drag
'form of which is Me, and that is frmmain.
Drag Me
End Sub

Private Sub lblfile_Click()
'calls mnufile to popup from the menu form
PopupMenu menu.mnufile
End Sub

Private Sub lbllistoptions_Click()
'calls mnulistoptions to popup from menu form
PopupMenu menu.mnulistoptions
End Sub

Private Sub lblmake_Click()
'disables the label just clicked, so you dont get a
'runtime error for doing two events at once.
lblmake.Enabled = False

'so you cant enter text in the txtsn textbox or
'txtpw textbox, so it doesnt interfear with the making
'of the aim.
txtsn.Enabled = False
txtpw.Enabled = False

'sets lblstatus's caption to below text
lblstatus.Caption = "Creating the AIM SN [" & Replace(txtsn.Text, " ", "") & "]"

'main is where the action happens, its the base
'of the internet tranfer control.
'we load up the url to make an aim screen name
'as if we were doing it in a webbrowser, but
'replace the screen name and password with
'txtsn (screen name, and txtpw(password)
Main.Text = AIM.OpenURL("http://aim.aol.com/aimnew/create_new.adp?name=" & txtsn.Text & "&password=" & txtpw.Text & "&confirm=" & txtpw.Text & "&email=" & GenerateEmail & "@aol.com&month=01&day=12&year=1951&promo=106712&pageset=Aim&privacy=1&client=no")

'if the string in main.text is Your new Screen Name is..then does the following
'If the string is that, then you got yourself the screen name you wanted
If InStr(Main.Text, "Your new screen name is") Then

'sets status for your new screen name
lblstatus.Caption = "[" & Replace(txtsn.Text, " ", "") & "-" & txtpw.Text & "] successfully made."

'since you got the screen name, it adds it to list1
List1.AddItem txtsn

'list1count is a lablel..sets the caption to list1's listcount
list1count.Caption = List1.ListCount

'adds cordinating password your choose (txtpw.text)
List2.AddItem txtpw

'list2count is a label, sets it to list2's listcount
list2count.Caption = List2.ListCount

'if enough characters, horizontal scroll bar will appear on list1
Call AddHScroll(List1)

'if enough characters, horizontal scroll bar will appear on list2
Call AddHScroll(List2)

'Call Save2Lists(List1, List2, menu.CmDialog1.FileName)

'clears the clipboard of previous content
Call Clipboard.Clear

'copys following text to your clipboard
Call Clipboard.SetText("[*AIM SN [" & Replace(txtsn.Text, " ", "") & ":" & txtpw.Text & "]*]")

'clears main.text (where the action happens, lol)
Main.Text = ""

'chatsends the sn you made:
Call sendaims("" + txtsn + "")

If menu.Check3.value = 1 Then 'checks to see if auto save is on?
Call Save2Lists(List1, List2, menu.CmDialog1.FileName)
End If
'now all of that was done if main.text was, Your New Screen Name is....
'now if the screen name is invalid...we continue
'on with the ElseIf statement....

'if main.text was "Sorry" then...
ElseIf InStr(Main.Text, "Sorry") Then

'sets lblstatus to..error with the sn.....etc
lblstatus.Caption = "Error with the SN [" & Replace(txtsn.Text, " ", "") & "] -evil aim sn maker 2"

'clears main.text of any text
Main.Text = ""
'else statement..now if its not Your New Screen name...
'and if its not Sorry... the else statement makes it
'continue through the coding.
Else

'if the string in main.text is Try a new Screen Name..
'this means your screen name in txtsn was taken.
If InStr(Main.Text, "Try a new Screen Name") Then
'sets lblstatus's caption to the sn(you wanted) is taken
lblstatus.Caption = "The SN [" & Replace(txtsn.Text, " ", "") & "] is taken' -evil aim sn maker 2"

'clears main.text
Main.Text = ""

'ends the if statements
End If
End If

'clears then sn textbox
txtsn.Text = ""
'alot of times people like to use the same password
'so i kept it as if they were, so txtpw stays there
'txtpw.text=""

'enabled txtsn and txtpw so you can make another sn
txtsn.Enabled = True
txtpw.Enabled = True

'clears main.text again.
Main.Text = ""

'sets focus back on txtsn.text
txtsn.SetFocus

'this enables the label that is over Make, so you
'can click it again and go through this process.
lblmake.Enabled = True
End Sub

Private Sub lblminimize_Click()
'me(frmmain) .windowstate=1, which means minimized
Me.WindowState = 1
End Sub

Private Sub lblmore_Click()
'calls mnumore to popup from menu form
PopupMenu menu.mnumore
End Sub

Private Sub lblsettings_Click()
'calls mnusettings to popup from menu form
PopupMenu menu.mnusettings
End Sub

Private Sub lblx_Click()
'unloads all the forms
Unload frmintro
Unload frmmain
Unload menu
End Sub

Private Sub List1_Click()
On Error Resume Next
'if list3's index is a negitive one then exit sub
If List1.ListIndex = -1 Then Exit Sub
'line up what was clicked in each list
List2.Selected(List1.ListIndex) = True
End Sub
Private Sub List1_DblClick()
'removes selected from list1
Call ListRemoveSelected(List1)
'resets the listcount caption
list1count.Caption = List1.ListCount

'removes the selected from list2...
Call ListRemoveSelected(List2)
'resets the listcount's caption
list2count.Caption = List2.ListCount
If menu.Check3.value = 1 Then
Call Save2Lists(frmmain.List1, frmmain.List2, menu.CmDialog1.FileName)
End If
End Sub

Private Sub List2_Click()
On Error Resume Next

'if list4's index is a negitive one the exit sub
If List2.ListIndex = -1 Then Exit Sub

'line up what was clicked in each list
List1.Selected(List2.ListIndex) = True
End Sub

Private Sub txtsn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lblmake_Click
End If
End Sub
