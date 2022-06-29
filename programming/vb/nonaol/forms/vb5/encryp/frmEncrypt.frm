VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmEncrypt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encrypter 3.o Final   Made by SaBrE"
   ClientHeight    =   4272
   ClientLeft      =   120
   ClientTop       =   576
   ClientWidth     =   7008
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEncrypt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4272
   ScaleWidth      =   7008
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   5160
      Top             =   1920
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   327681
      DialogTitle     =   "Save"
      FileName        =   "Encrypt1.txt"
      Filter          =   "Text Documents (*.txt)|*.txt"
      FilterIndex     =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   1560
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   327681
      DefaultExt      =   ".txt"
      DialogTitle     =   "Open"
      FileName        =   ".txt"
      Filter          =   "Text Documents (*.txt)|*.txt"
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4320
      TabIndex        =   18
      ToolTipText     =   "Click Here to Print the text"
      Top             =   3840
      Width           =   1176
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Check It"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2880
      TabIndex        =   17
      ToolTipText     =   "Click Here to Spell Check the text"
      Top             =   3840
      Width           =   1176
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3132
      Left            =   5640
      TabIndex        =   10
      Top             =   1080
      Width           =   1332
      Begin VB.CommandButton Command10 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Bart"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2640
         Width           =   372
      End
      Begin VB.CommandButton Command9 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Bart"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   600
         TabIndex        =   12
         Top             =   1680
         Width           =   372
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   612
      End
      Begin VB.Label Label3 
         Caption         =   "Back Color:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Forecolor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "Font Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   852
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      TabIndex        =   9
      ToolTipText     =   "Click Here to Cut the text"
      Top             =   3360
      Width           =   1176
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   8
      ToolTipText     =   "Click Here to Open file"
      Top             =   3840
      Width           =   1176
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Click Here to Save file"
      Top             =   3840
      Width           =   1176
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Copy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "Click Here to Copy the text"
      Top             =   3360
      Width           =   1176
   End
   Begin VB.Frame Frame1 
      Caption         =   "Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   5640
      TabIndex        =   3
      Top             =   0
      Width           =   1332
      Begin VB.OptionButton Option2 
         Caption         =   "Small"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   732
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Large"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1260
      Left            =   0
      MaxLength       =   5000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   "Type the String to be Encrypted Here"
      Top             =   0
      Width           =   3960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Click Here to Clear the Text and Reset"
      Top             =   3360
      Width           =   1176
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "&Encrypt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Click here to Encrypt/Decrypt"
      Top             =   3360
      Width           =   1176
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
      End
      Begin VB.Menu Open 
         Caption         =   "&Open..."
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu Divide1 
         Caption         =   "-"
      End
      Begin VB.Menu Print 
         Caption         =   "&Print..."
      End
      Begin VB.Menu Divide2 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
      Begin VB.Menu Quit 
         Caption         =   "E&xit..."
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Cut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu Clear 
         Caption         =   "Clear..."
      End
      Begin VB.Menu Divide3 
         Caption         =   "-"
      End
      Begin VB.Menu TimeDate 
         Caption         =   "Time/&Date"
      End
      Begin VB.Menu SetFont 
         Caption         =   "Set &Font..."
      End
   End
End
Attribute VB_Name = "frmEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MydsEncrypt As dsEncrypt

Private Sub About_Click()
MsgBox "This is my final encrypter.  I included all that u could think of for a full encypter that people would use.  Cya", , ""
End Sub

Private Sub Clear_Click()
If Command1(0).Caption = "U&nEncrypt" Then Command1(0).Caption = "&Encrypt"
Text1.Text = ""
Text1.Font = "Arial"
Combo1.Text = "9"
Text1.ForeColor = vbBlack
Text1.BackColor = vbWhite
Text1.Height = 1260
Text1.Width = 3960
Option2.Value = True
Text1.Enabled = True
End Sub

Private Sub Combo1_Click()
Text1.FontSize = Combo1.Text
End Sub

Private Sub Command1_Click(Index As Integer)
If Command1(0).Caption = "&Encrypt" Then
  Command1(0).Caption = "U&nEncrypt"
   Text1.Enabled = False
   Paste.Enabled = False
  Else
  Command1(0).Caption = "&Encrypt"
  Text1.Enabled = True
  Paste.Enabled = True
End If
If Text1.Text = "" Then
  MsgBox "Please enter sumthing to encrypt/decrypt", , ""
  Text1.Enabled = True
  Command1(0).Caption = "&Encrypt"
Else
  Text1.Text = MydsEncrypt.Encrypt(Text1.Text)
  Text1.Height = 3200
  Text1.Width = 5500
  Option1.Value = True
End If
End Sub

Private Sub Command10_Click()
CommonDialog1.Flags = &H1&
CommonDialog1.ShowColor
Text1.BackColor = CommonDialog1.Color
End Sub

Private Sub Command2_Click()
If Command1(0).Caption = "U&nEncrypt" Then Command1(0).Caption = "&Encrypt"
Text1.Text = ""
Text1.Font = "Arial"
Combo1.Text = "9"
Text1.ForeColor = vbBlack
Text1.BackColor = vbWhite
Text1.Height = 1260
Text1.Width = 3960
Option2.Value = True
Text1.Enabled = True
End Sub

Private Sub Command3_Click()
If Text1.Enabled = False Then
Clipboard.SetText (Text1.Text)
Else
Clipboard.SetText (Text1.SelText)
End If
End Sub

Private Sub Command4_Click()
If Text1.Text = "" Then
MsgBox "Nothing to save!", , ""
Exit Sub
Else
CommonDialog2.ShowSave 'display Save dialog box
Open CommonDialog2.filename For Output As #1
Print #1, Text1.Text 'make file and add text
Close #1 'close file
End If
End Sub

Private Sub Command5_Click()
Dim Wrap As String, AllText As String, LineOfText As String
Wrap$ = Chr$(13) + Chr$(10)  'create wrap character
    CommonDialog1.ShowOpen       'display Open dialog box
    On Error GoTo Error:
    If CommonDialog1.filename <> "" Then
        Open CommonDialog1.filename For Input As #1
        On Error GoTo TooBig:    'set error handler
        Do Until EOF(1)          'read lines from file
            Line Input #1, LineOfText$
            AllText$ = AllText$ & LineOfText$ & Wrap$
        Loop
        Text1.Text = AllText$  'display file
        Text1.Enabled = True
        Text1.Height = 3200
        Text1.Width = 5500
        Option1.Value = True
        Command1(0).Caption = "&Encrypt"
CleanUp:
        Close #1                 'close file
    End If
    Exit Sub
TooBig:             'error handler displays message
    MsgBox "The file is just too large, captain!", , ""
    Resume CleanUp: 'then jumps to CleanUp routine
Error:
    Exit Sub
End Sub

Private Sub Command6_Click()
If Text1.Enabled = False Then
Clipboard.SetText (Text1.Text)
Text1.Text = ""
Else
Clipboard.SetText (Text1.SelText)
Text1.SelText = ""
End If
End Sub

Private Sub Command7_Click()
Dim wdDoNotSaveChanges As Variant
Dim X As Object      'create Word object variable
If Text1.Text = "" Then
MsgBox "No text to spellcheck"
Exit Sub
End If
Set X = CreateObject("Word.Application") 'create it
X.Visible = False    'hide Word
X.Documents.Add      'open a new document
X.Selection.Text = Text1.Text  'copy text box to document
X.ActiveDocument.CheckSpelling 'run spell/grammar check
Text1.Text = X.Selection.Text  'copy results back
X.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges 'dont say changes
X.Quit               'quit Word
Set X = Nothing      'release object variable
MsgBox "Spellcheck complete!", , ""
End Sub

Private Sub Command8_Click()
If Text1.Text = "" Then
MsgBox "Nothing to print!", , ""
Exit Sub
Else
Printer.FontSize = Combo1.Text 'set fontsize
Printer.Print Text1 'print the text
End If
End Sub

Private Sub Command9_Click()
CommonDialog1.Flags = &H1&
CommonDialog1.ShowColor
Text1.ForeColor = CommonDialog1.Color
End Sub

Private Sub Copy_Click()
If Text1.Enabled = False Then
Clipboard.SetText (Text1.Text)
Else
Clipboard.SetText (Text1.SelText)
End If
End Sub

Private Sub Cut_Click()
If Text1.Enabled = False Then
Clipboard.SetText (Text1.Text)
Text1.Text = ""
Else
Clipboard.SetText (Text1.SelText)
Text1.SelText = ""
End If
End Sub

Private Sub Form_Load()
    Set MydsEncrypt = New dsEncrypt
    MydsEncrypt.KeyString = ("KATHER")
    Command1(0).Caption = "&Encrypt"
    Option2.Value = True
    Dim i
    For i = 1 To 54
    Combo1.AddItem i
    Next i
    Combo1.Text = "9"
End Sub

Private Sub New_Click()
Text1.Text = ""
Text1.Font = "Arial"
Combo1.Text = "9"
Text1.ForeColor = vbBlack
Text1.BackColor = vbWhite
Text1.Height = 1260
Text1.Width = 3960
Text1.Enabled = True
End Sub

Private Sub Open_Click()
Dim Wrap As String, AllText As String, LineOfText As String
Wrap$ = Chr$(13) + Chr$(10)  'create wrap character
    CommonDialog1.ShowOpen       'display Open dialog box
    On Error GoTo Error:
    If CommonDialog1.filename <> "" Then
        Open CommonDialog1.filename For Input As #1
        On Error GoTo TooBig:    'set error handler
        Do Until EOF(1)          'read lines from file
            Line Input #1, LineOfText$
            AllText$ = AllText$ & LineOfText$ & Wrap$
        Loop
        Text1.Text = AllText$  'display file
        Text1.Enabled = True
        Text1.Height = 3200
        Text1.Width = 5500
        Option1.Value = True
        Command1(0).Caption = "&Encrypt"
CleanUp:
        Close #1                 'close file
    End If
    Exit Sub
TooBig:             'error handler displays message
    MsgBox "The file is just too large, captain!", , ""
    Resume CleanUp: 'then jumps to CleanUp routine
Error:
    Exit Sub
End Sub

Private Sub Option1_Click()
Text1.Height = 3200
Text1.Width = 5500
End Sub

Private Sub Option2_Click()
Text1.Height = 1260
Text1.Width = 3960
End Sub

Private Sub Paste_Click()
Text1.Text = Text1.Text + Clipboard.GetText
End Sub

Private Sub Print_Click()
Printer.FontSize = Combo1.Text
Printer.Print Text1
End Sub

Private Sub Quit_Click()
End
Unload Me
Unload frmFont
End Sub

Private Sub Save_Click()
If Text1.Text = "" Then
MsgBox "Nothing to save!", , ""
Exit Sub
Else
CommonDialog2.ShowSave 'display Save dialog box
Open CommonDialog2.filename For Output As #1
Print #1, Text1.Text 'make file and add text
Close #1 'close file
End If
End Sub

Private Sub SetFont_Click()
frmEncrypt.Visible = False
frmFont.Show 1
End Sub

Private Sub TimeDate_Click()
Dim Tme As String, Dte As String
Tme$ = Time
Dte$ = Date
Text1.Text = Text1.Text + " " + Tme$ + " / " + Dte$ + " "
End Sub

