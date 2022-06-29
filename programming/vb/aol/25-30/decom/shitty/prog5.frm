VERSION 4.00
Begin VB.Form Form5 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "-=RÍ–i„£ÍR=-"
   ClientHeight    =   1755
   ClientLeft      =   3540
   ClientTop       =   2370
   ClientWidth     =   2445
   ForeColor       =   &H00800000&
   Height          =   2160
   Icon            =   "PROG5.frx":0000
   Left            =   3480
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   2445
   Top             =   2025
   Width           =   2565
   Begin VB.OptionButton Option2 
      BackColor       =   &H00800000&
      Caption         =   "Option2"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1530
      TabIndex        =   8
      Top             =   810
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1560
      TabIndex        =   21
      Text            =   "Text6"
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   480
      TabIndex        =   20
      Text            =   "Text5"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   630
      Left            =   0
      TabIndex        =   17
      Top             =   -90
      Width           =   2445
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1500
         TabIndex        =   19
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Tries:"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   180
         Width           =   1200
      End
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   480
      TabIndex        =   15
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Let's Go"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1530
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   480
      TabIndex        =   11
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   10
      Top             =   2760
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      Caption         =   "Option1"
      Height          =   225
      Left            =   180
      TabIndex        =   6
      Top             =   810
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Let's Go"
      Height          =   195
      Left            =   1470
      TabIndex        =   3
      Top             =   1530
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   1050
      Width           =   930
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      PasswordChar    =   "@"
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      Caption         =   "Yepper"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   450
      TabIndex        =   7
      Top             =   780
      Width           =   615
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00800000&
      Caption         =   "Fast Method"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   1290
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800000&
      Caption         =   "Slow Method"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1440
      TabIndex        =   12
      Top             =   1290
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      Caption         =   "Nope"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1770
      TabIndex        =   9
      Top             =   780
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Auto Download?"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   570
      Width           =   1245
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Enter your password"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form5"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
If label8.Caption = "Stop" Then
Command1.Caption = "Cancel"
label8.Caption = "ok"
Exit Sub
End If

Form5.Hide
Form4.Show
End Sub

Private Sub Command2_Click()
    'Exit out if no password is entered
    If Text1.text = "" Then
        MsgBox "Hmmm somethings wrong here oh yeah...it helps if you put in a password..."
        Exit Sub
    End If
    'Start up america online if it hasn't already...change focus
  

    If America% = 0 Then
        AmericaStart = Shell("C:\aol30\aol", 1)
        Else
        AppActivate "America"
    End If
    
    Command1.Caption = "STOP!!!"
    label8.Caption = "Stop"
    'Declare Windows
    America% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(America%, "MDIClient")
    welcome% = FindChildByTitle(MDI%, "Welcome")
    Online% = FindChildByTitle(MDI%, "Welcome,")
    Call sendtext(America%, "Deicide's OnLiNe HeLL!")
    ' This is the start of the finding edit process
    Z = aolhwnd = welcome%
    hChild = GetWindow(welcome%, GW_CHILD)
    Edit = GetClass(hChild)
    Text2.text = Edit
    If Text2.text = "_AOL_Edit" Then
        Z = aolhwnd = hChild
        Call sendtext(hChild, "")
        Call sendtext(hChild, Text1.text)
    Else
        Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Edit = GetClass(hChild)
        Text2.text = Edit
        Loop Until Text2.text = "_AOL_Edit"
        Z = aolhwnd = hChild
        Call sendtext(hChild, "")
        Call sendtext(hChild, Text1.text)
    End If
        ' This is the start of the finding button process
    Call SendCharNum(hChild, 13)
    'end of clicking
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
    SendMessage FindWindow("#32770", "America Online"), WM_CLOSE, 0, 0
    Call sendtext(hChild, Text1.text)
    Sendclick (hchild2)
    End If
    Start = Timer
    X = 1
    Label12.Caption = X
    
    Do
Do: DoEvents
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
    SendMessage FindWindow("#32770", "America Online"), WM_CLOSE, 0, 0
    Call sendtext(hChild, Text1.text)
    Sendclick (hchild2)
    End If
    America% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(America%, "MDIClient")
    Online% = FindChildByTitle(MDI%, "Welcome, ")
    If Online% <> 0 Then
    Dim TotalTime As Integer
    Finish = Timer
    TotalTime = Finish - Start
    If Option1.Value = True Then
    Call RunMenuByString(America%, "Download Manager")
    End If
    America% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(America%, "MDIClient")
    Do: DoEvents
    Download% = FindChildByTitle(MDI%, "Download Manager")
    Loop Until Download% <> 0
    Z = aolhwnd = Download%
    hchild11 = GetWindow(Download%, GW_CHILD)
    down = GetClass(hchild11)
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Sendclick (hchild11)
        MsgBox "Redialer redialed " & X & " times this took " & TotalTime & " seconds"
    Command1.Caption = "Done"
    label8.Caption = ""
    Exit Sub
    Else
        MsgBox "Redialer redialed " & X & " times this took " & TotalTime & " seconds"
    Command1.Caption = "Done"
    label8.Caption = ""
    Exit Sub
    End If
    Loop Until FindWindow("#32770", "Connect Error")
    SendMessage FindWindow("#32770", "Connect Error"), WM_CLOSE, 0, 0
    Call sendtext(hChild, Text1.text)
    Call SendCharNum(hChild, 13)
    X = 1 + X
    Label12.Caption = X
    Y = X
    Loop
End Sub


Public Sub Command3_Click()
    
'Exit out if no password is entered
    If Text1.text = "" Then
        MsgBox "Hmmm somethings wrong here oh yeah...it helps if you put in a password..."
        Exit Sub
    End If
        
        America% = FindWindow("AOL Frame25", vbNullString)
    'Start up america online if it hasn't already...change focus
        If America% = 0 Then
            mypath = CurDir
            Open mypath & "\deicide.ini" For Random As #1
            Get #1, 1, Path$
                                    
            'If user hasnt found Dir yet...this tells them to
            If Path$ = "" Then
                MsgBox "Umm you need to go to AOL dir in the prog before you can use this :("
                Exit Sub
            End If
            AmericaStart = Shell(Path$, 1)
            Close #1
            'Loops until AOL is open then finds windows
            Do: DoEvents
                America% = FindWindow("AOL Frame25", vbNullString)
                Call sendtext(America%, "Deicide's OnLiNe HeLL!")
                MDI% = FindChildByClass(America%, "MDIClient")
                welcome% = FindChildByTitle(MDI%, "Welcome")
            Loop Until welcome% <> 0
        'If AOL is started finds windows
        Else
                Do: DoEvents
                America% = FindWindow("AOL Frame25", vbNullString)
                Call sendtext(America%, "Deicide's OnLiNe HeLL!")
                MDI% = FindChildByClass(America%, "MDIClient")
                welcome% = FindChildByTitle(MDI%, "Welcome")
                Loop Until welcome% <> 0
        End If
    
    Frame1.Visible = True
    Command1.Caption = "STOP!!!"
    label8.Caption = "Stop"
    'Declare Windows
    Online% = FindChildByTitle(MDI%, "Welcome,")
    
    ' This is the start of the finding edit process
    Z = aolhwnd = welcome%
    hChild = GetWindow(welcome%, GW_CHILD)
    Edit = GetClass(hChild)
    Text2.text = Edit
    If Text2.text = "_AOL_Edit" Then
        Z = aolhwnd = hChild
        Call sendtext(hChild, "")
        Call sendtext(hChild, Text1.text)
    Else
        Do: DoEvents
        hChild = GetWindow(hChild, GW_HWNDNEXT)
        Edit = GetClass(hChild)
        Text2.text = Edit
        Loop Until Text2.text = "_AOL_Edit"
        Z = aolhwnd = hChild
        Call sendtext(hChild, "")
        Call sendtext(hChild, Text1.text)
    End If
        ' This is the start of the finding button process
   Call SendCharNum(hChild, 13)
   
    'start our wonderful timer :)
    Start = Timer
 
      'start of finding Static window
    Do
    Do: DoEvents
    freak% = FindWindow("_AOL_Modal", vbNullString)
    Pass% = FindWindow("#32770", "America Online")
    shitz% = FindWindow("#32770", "Connect Error")
    If Command1.Caption = "ok" Then
    Command1.Caption = "Cancel"
    Exit Sub
    End If
    If Pass% <> 0 Then
    SendMessage FindWindow("#32770", "America Online"), WM_CLOSE, 0, 0
    Call sendtext(hChild, Text1.text)
    Sendclick (hchild2)
    End If
    'See if it screwed up :)
    If shitz% <> 0 Then
    SendMessage FindWindow("#32770", "Connect Error"), WM_CLOSE, 0, 0
    X = X + 1
    Label12.Caption = X
    Call sendtext(hChild, Text1.text)
    Sendclick (hchild2)
    End If
    Loop Until freak% <> 0
  
    Z = aolhwnd = freak%
    hChild6 = GetWindow(freak%, GW_CHILD)
    statics = GetClass(hChild6)
    text4.text = statics
    If text4.text = "_AOL_Static" Then
    Else
        Do: DoEvents
        hChild6 = GetWindow(hChild6, GW_HWNDNEXT)
        statics = GetClass(hChild6)
        text4.text = statics
        Loop Until text4.text = "_AOL_Static"
        Z = aolhwnd = hChild6
        
        'Start of finding Cancel button
      
        Z = aolhwnd = freak%
    
    hchild3 = GetWindow(freak%, GW_CHILD)
    Button = GetClass(hchild3)
    text4.text = Button
    If text4.text = "_AOL_Icon" Then
    Else
        Do: DoEvents
        hchild3 = GetWindow(hchild3, GW_HWNDNEXT)
        Button = GetClass(hchild3)
        text4.text = Button
        Loop Until text4.text = "_AOL_Icon"
        Z = aolhwnd = hchild3
      
        End If
        
        'sending stats to canceling when busy
        Do: DoEvents
        abc = AOLGetText(hChild6)
        America% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(America%, "MDIClient")
    Online% = FindChildByTitle(MDI%, "Welcome,")
    If Online% <> 0 Then
    Dim TotalTime As Integer
    Finish = Timer
    TotalTime = Finish - Start
    If Option1.Value = True Then
    Call RunMenuByString(America%, "Download Manager")
    Command1.Caption = "Done"
    label8.Caption = ""
    America% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(America%, "MDIClient")
    Do: DoEvents
    Download% = FindChildByTitle(MDI%, "Download Manager")
    Loop Until Download% <> 0
    Z = aolhwnd = Download%
    hchild11 = GetWindow(Download%, GW_CHILD)
    down = GetClass(hchild11)
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Sendclick (hchild11)
        Sendclick (hchild11)
        If Y = 1 Then
        MsgBox "Redialer Just had to dial once this took " & TotalTime & " seconds"
        Exit Sub
        Else
        MsgBox "Redialer redialed " & X & " times this took " & TotalTime & " seconds"
        Exit Sub
        End If
        Else
        If Y = 1 Then
        Command1.Caption = "Done"
        label8.Caption = ""
        MsgBox "Redialer Just had to dial once this took " & TotalTime & " seconds"
        Exit Sub
        Else
        Exit Sub
        MsgBox "Redialer redialed " & X & " times this took " & TotalTime & " seconds"
        Command1.Caption = "Done"
        label8.Caption = ""
        End If
    End If
    Exit Sub
    End If
        Loop Until abc = "The first try was busy, trying second number" Or abc = "The first try had no dial tone, trying second number" Or abc = "The first try did not respond, trying second number"
        Sendclick (hchild3)
        Sendclick (hchild3)
        X = 1 + X
        Label12.Caption = X
        End If
        
        'Starting process over
        If label8.Caption = "ok" Then
        label8.Caption = "bob"
        Exit Sub
        End If
        Call sendtext(hChild, Text1.text)
        Call SendCharNum(hChild, 13)
    Do: DoEvents
    freazz% = FindWindow("_AOL_Modal", vbNullString)
    Pass% = FindWindow("#32770", "America Online")
    If Pass% <> 0 Then
    SendMessage FindWindow("#32770", "America Online"), WM_CLOSE, 0, 0
    Call sendtext(hChild, Text1.text)
    Call SendCharNum(hChild, 13)
    End If
    America% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(America%, "MDIClient")
    Online% = FindChildByTitle(MDI%, "Welcome,")
    If Online% <> 0 Then
    Dim totaltime2 As Integer
    Finish = Timer
    TotalTime = Finish - Start
    Dim Min As Integer
    Dim Sec As Integer
    Min = TotalTime / 60
    Sec = TotalTime Mod 60
    
    'run options
    If Option1.Value = True Then
    Call RunMenuByString(America%, "Download Manager")
    Command1.Caption = "Done"
    label8.Caption = ""
    America% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(America%, "MDIClient")
    Do: DoEvents
    Download% = FindChildByTitle(MDI%, "Download Manager")
    Loop Until Download% <> 0
    'find right button :) takes a bit huh?
    Z = aolhwnd = Download%
    hchild11 = GetWindow(Download%, GW_CHILD)
    down = GetClass(hchild11)
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        down = GetClass(hchild11)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Do: DoEvents
        hchild11 = GetWindow(hchild11, GW_HWNDNEXT)
        Text2.text = down
        Loop Until Text2.text = "_AOL_Icon"
        Z = aolhwnd = hchild11
        Call SendCharNum(hChild, 13)
        Command1.Caption = "Done"
        label8.Caption = ""
        MsgBox "Redialer redialed " & Min & " " & Sec & " times this took " & totaltime2 & " seconds"
        Exit Sub
    Else
        Command1.Caption = "Done"
        label8.Caption = ""
        MsgBox "Redialer redialed " & Min & " " & Sec & " times this took " & TotalTime & " seconds"
        Exit Sub
    End If
    Command1.Caption = "Done"
    label8.Caption = ""
    Exit Sub
    End If
    Loop Until freazz% <> 0
    Loop
End Sub


Private Sub Command4_Click()
a% = FindWindow("AOL Frame25", vbNullString)  'Find AOL
    

If Command4.Caption = "H" Then
  
    X = ShowWindow(a%, SW_HIDE)
    Command4.Caption = "S"
    Exit Sub
End If
If Command4.Caption = "S" Then
    
    X = ShowWindow(a%, SW_SHOW)
    Command4.Caption = "H"
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Call StayOnTop(Form5)
Frame1.Visible = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


Private Sub Option2_Click()
MsgBox "Ok I won't start up the Download Manager"
End Sub


Private Sub Timer1_Timer()

End Sub


