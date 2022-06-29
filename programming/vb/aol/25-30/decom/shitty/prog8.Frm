VERSION 4.00
Begin VB.Form Form8 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bust"
   ClientHeight    =   615
   ClientLeft      =   3570
   ClientTop       =   3300
   ClientWidth     =   1890
   Height          =   1020
   Left            =   3510
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   1890
   ShowInTaskbar   =   0   'False
   Top             =   2955
   Width           =   2010
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   1230
      Width           =   1425
   End
   Begin VB.CommandButton Command3 
      Caption         =   "--"
      Height          =   195
      Left            =   810
      TabIndex        =   8
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Private Rooms"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   900
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Member Rooms"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1425
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1470
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   195
      Left            =   1110
      TabIndex        =   2
      Top             =   360
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bust in"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   285
      Left            =   330
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
End
Attribute VB_Name = "Form8"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
'Declare Windows
America% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(America%, "MDIClient")
Roomsm% = FindChildByClass(MDI%, "AOL Child")
Imagex% = FindChildByClass(Roomsm%, "_AOL_Image")

If America% = 0 Then
MsgBox "Gee it helps if AOL is open"
End If

'Change Caption so it will stop
If Command1.Caption = "Stop" Then
Command1.Caption = "Bust me in"
Label2.Caption = "bye"
Form4.Show
Unload Form8
Pause (5)
Exit Sub
Else
Label2.Caption = "hi"
Command1.Caption = "Stop"
End If
        'Trying to find room
        Roomsm% = FindChildByClass(MDI%, "AOL Child")
        If Roomsm% <> 0 Then
        chatz% = FindChildByClass(Roomsm%, "_AOL_View")
        End If
        If chatz% <> 0 Then
        abb = AOLGetText(Roomsm%)
        End If
        If chatz% = 0 Or Roomsm% = 0 Then
        abb = 0
        End If
        bcb = abb

'Call Go to menu
Call RunMenuByString(America%, "Edit Go To Menu")

'Find Window's in Go to menu to type in
Do: DoEvents
Modal% = FindWindow("_AOL_Modal", "Favorite Places")
Loop Until Modal% <> 0
Z = aolhwnd = Modal%
hChild = GetWindow(Modal%, GW_CHILD)
place = GetClass(hChild)
Text2.text = place
If Text2.text = "_AOL_Edit" Then
Y = aolhwnd = hChild
Call sendtext(hChild, "")
Call sendtext(hChild, Text1.text)
Else
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
place = GetClass(hChild)
Text2.text = place
Loop Until Text2.text = "_AOL_Edit"
Y = aolhwnd = hChild
Call sendtext(hChild, "")
Call sendtext(hChild, Text1.text)
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
place = GetClass(hChild)
Text2.text = place
Loop Until Text2.text = "_AOL_Edit"
Y = aolhwnd = hChild
Call sendtext(hChild, "")

If Option1.Value = True Then
Call sendtext(hChild, "aol://2719:62-2-" & Text1.text)
Else
Call sendtext(hChild, "aol://2719:2-2-" & Text1.text)
End If

End If
Call SendCharNum(hChild, 13)
X = 0
MDI% = FindChildByClass(America%, "MDIClient")
Dim TotalTime As Integer
Start = Timer
Do: DoEvents
X = X + 1
Call RunMenuByString(America%, Text1.text)
    Do: DoEvents
    Roomsm% = FindChildByClass(MDI%, "AOL Child")
    chatz% = FindChildByClass(Roomsm%, "_AOL_View")
    If chatz% <> 0 Then
        abb = AOLGetText(Roomsm%)
    End If
    Loop Until FindWindow("#32770", "America Online") Or bcb <> abb
    SendMessage FindWindow("#32770", "America Online"), WM_CLOSE, 0, 0
Loop Until bcb <> abb
Finish = Timer
TotalTime = Finish - Start
roomsmz% = FindChildByClass(MDI%, "AOL Child")
Z = aolhwnd = roomsmz%
hChild = GetWindow(roomsmz%, GW_CHILD)
room = GetClass(hChild)
Text2.text = room
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
room = GetClass(hChild)
Text2.text = room
Y = aolhwnd = hChild
Loop Until Text2.text = "_AOL_Edit"
If Label2.Caption = "bye" Then
Exit Sub
Else
Call sendtext(hChild, "-=Ðê‡ç‡Ðê™=- broke me in this room in " & X & " tries and in " & TotalTime & " sec's")
Call SendCharNum(hChild, 13)
Call SendCharNum(hChild, 13)
Command1.Caption = "Bust me in"
End If
End Sub


Private Sub Command2_Click()
Unload Form8
Form4.Show
End Sub


Private Sub Command3_Click()
X = Form8.Height
If X < 1500 Then
Do
X = X + 1
Form8.Height = X
Loop Until Form8.Height > 1575
Form8.Height = 1575
Exit Sub
End If
If X > 1500 Then
Do
X = X - 1
Form8.Height = X
Loop Until Form8.Height < 960
Form8.Height = 960
Exit Sub
End If
End Sub

Private Sub Form_Load()
Call StayOnTop(Form8)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


Private Sub Option3_Click()
'aol://2719:61(mr)/21(publicroom)/2(pr)-2-room name
End Sub


