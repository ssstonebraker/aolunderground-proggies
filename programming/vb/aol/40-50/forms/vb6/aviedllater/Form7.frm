VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   0  'None
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   3855
   ControlBox      =   0   'False
   Icon            =   "FORM7.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   3465
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      ItemData        =   "FORM7.frx":030A
      Left            =   120
      List            =   "FORM7.frx":030C
      OLEDropMode     =   1  'Manual
      Sorted          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "double click to remove"
      Top             =   1680
      Width           =   3615
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      ItemData        =   "FORM7.frx":030E
      Left            =   120
      List            =   "FORM7.frx":0310
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      TabIndex        =   2
      ToolTipText     =   "drag to bottom list to add"
      Top             =   360
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   6
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   135
      Left            =   3600
      TabIndex        =   12
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "#selected (label8)"
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "other"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   135
      Left            =   3600
      TabIndex        =   8
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "start"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "select"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "add mails"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "status: idle"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "AviE auto download later by knot"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Menu blah 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu selall 
         Caption         =   "select all"
      End
      Begin VB.Menu desALL 
         Caption         =   "remove all"
      End
      Begin VB.Menu warez 
         Caption         =   "click for warez"
      End
      Begin VB.Menu blahhdhgfhfhgjfrg 
         Caption         =   "-"
      End
      Begin VB.Menu unLd 
         Caption         =   "unload me"
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DlLater
End Sub

Private Sub desALL_Click()
List2.Clear
End Sub

Private Sub Form_Load()
FormOnTop Me, True
Call SendChat("<Font Face=Tahoma></b></i></u>• AviE auto download later: loaded.")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontUnderline = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call SendChat("<Font Face=Tahoma></b></u></i>• AviE auto download later: unloaded.")

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
knot2.FormDrag Me
End Sub

Private Sub Label10_Click()
Me.WindowState = 1
End Sub

Private Sub Label2_Click()
Dim Mailbox As Long, i As Long, Number As String
If GetAolVer = False Then
Call MsgBox("This is only for aol4/5.")
Exit Sub
End If
List1.Clear
Label3.Caption = "status: refreshing/adding mails"
knot2.MailOpenNew
Pause 1.09
Mailbox = FindMailBox

    For i = 0 To MailCountNew - 1
        If Len(CStr(i + 1)) = 1 Then
            Number = "  " & CStr(i + 1)
        ElseIf Len(CStr(i + 1)) = 2 Then
            Number = " " & CStr(i + 1)
        Else
            Number = CStr(i + 1)
        End If

        List1.AddItem Number & ". " & MailGetSubject(i)
    Next

Pause 0.1
Label3.Caption = "status: closing mailbox"
SendMessage Mailbox, knot2.WM_CLOSE, 0&, 0&
Pause 0.1
Label3.Caption = "status: idle"
End Sub

Private Sub Label4_Click()
'Dim i As Long
'For i = 0 To List1.ListCount - 1
'If InStr(LCase(TrimSpacess(List1.List(i))), (LCase(TrimSpacess(Text1.Text)))) Then
'List2.AddItem
'End If
'Next i
'------- new code here -----
For i = 0 To List1.ListCount - 1
            If InStr(LCase(List1.List(i)), LCase(Text1.Text)) > 0 Then
                If ListBoxCheckDup(List2, List1.List(i)) = False Then
                 List2.AddItem List1.List(i)
                End If
            End If
        Next

End Sub
Public Function MailGetSubject(Index As Long) As String
    Dim Mailbox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String
    Dim Count As Long
    Mailbox = FindMailBox
    If Mailbox = 0& Then Exit Function
    TabControl = FindWindowEx(Mailbox, 0&, "_AOL_TabControl", vbNullString)
    TabPage = FindWindowEx(TabControl, 0&, "_AOL_TabPage", vbNullString)
    mTree = FindWindowEx(TabPage, 0&, "_AOL_Tree", vbNullString)
    Count = SendMessage(mTree, knot2.LB_GETCOUNT, 0&, 0&)
    If Count = 0 Or Index > Count - 1 Or Index < 0 Then Exit Function
    sLength = SendMessage(mTree, LB_GETTEXTLEN, Index, 0&)
    MyString = String(sLength + 1, 0)
    SendMessageByString mTree, knot2.LB_GETTEXT, Index, MyString
    Spot = InStr(MyString, Chr(9))
    Spot = InStr(Spot + 1, MyString, Chr(9))
    MyString = Mid(MyString, Spot + 1)
    MailGetSubject = Left(MyString, Len(MyString) - 1)
End Function

Public Function FindMail(Subject As String) As Long
    Dim aol As Long, MDI As Long, child As Long
    Dim Caption As String
    aol = FindWindow("AOL Frame25", vbNullString)
    MDI = FindWindowEx(aol, 0, "MDIClient", vbNullString)
    child = FindWindowEx(MDI, 0, "AOL Child", vbNullString)
    Caption = Left(Replace(GetCaption(child), " ", ""), 40)
    If Left(Replace(Subject, " ", ""), 40) = Caption Then
        FindMail = child
        Exit Function
    Else
        Do
            child = FindWindowEx(MDI, child, "AOL Child", vbNullString)
            Caption = Left(Replace(GetCaption(child), " ", ""), 40)
            If Left(Replace(Subject, " ", ""), 40) = Caption Then
                FindMail = child
                Exit Function
            End If
        Loop Until child = 0
    End If
    FindMail = 0
End Function

Private Sub Label5_Click()
    Dim Mailbox As Long, Total As Long, i As Long
    Dim Subject As String, Count As Long, MailIndex As Long

Label3.Caption = "status: opening mail box"
knot2.MailOpenNew
Pause 0.99
Mailbox = FindMailBox
    Total = List2.ListCount
    For i = 0 To List2.ListCount - 1
        List2.ListIndex = i
        List2.Selected(i) = True
        If i > 0 Then List2.Selected(i - 1) = False
        'MailIndex = CLng(Trim(Left(List2.List(i), 3))) - 1
        MailIndex = CLng(Trim(Left(List2.List(i), 4))) - 1

        Subject = MailGetSubject(MailIndex)
        MailOpenEmailNew MailIndex
        Do
            DoEvents
        Loop Until FindMail(Subject) <> 0
        'Loop Until FindMail
Pause 0.68867867
       

If FindDlLater = False Then knot2.DlLater
Pause 0.19

        Count = Count + 1
    SendMessage FindMail(Subject), knot2.WM_CLOSE, 0&, 0&
   Label3.Caption = "status: " & Count & " of " & Total & " done, " & knot2.percent2(Count, Total) & " done"
Next i
Pause 0.45
Label3.Caption = "status: closing mail box"
SendMessage Mailbox, WM_CLOSE, 0&, 0&
Pause 0.29
Label3.Caption = "status: idle"
List2.Clear
End Sub

Private Sub Label7_Click()
PopupMenu blah
End Sub

Private Sub Label9_Click()
Unload Me
End Sub

Private Sub List1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Dim i As Long, Temp As String
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            If Temp = "" Then
                Temp = i
            Else
                Temp = Temp & "|" & i
            End If
        End If
    Next
    Data.Clear
    Data.SetData Temp, vbCFText

End Sub

Private Sub List2_DblClick()
    If List2.ListCount = 0 Then Exit Sub
    Do
        DoEvents
        If List2.Selected(i) = True Then
            List2.RemoveItem i
        Else
            i = i + 1
        End If
    Loop Until i >= List2.ListCount

End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If KeyCode = 46 Then
            If List2.ListCount = 0 Then Exit Sub
    Do
        DoEvents
        If List2.Selected(i) = True Then
            List2.RemoveItem i
        Else
            i = i + 1
        End If
    Loop Until i >= List2.ListCount

    End If

End Sub

Private Sub List2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Temp As String, IndexArray() As String
    Temp = Data.GetData(vbCFText)
    IndexArray() = Split(Temp, "|")
    For i = LBound(IndexArray) To UBound(IndexArray)
        If ListBoxCheckDup(List2, List1.List(IndexArray(CInt(i)))) = False Then
            List2.AddItem List1.List(IndexArray(CInt(i)))
        End If
Next
End Sub

Private Sub selall_Click()
    If List1.ListCount = 0 Then Exit Sub
    Dim i As Integer
    For i = 0 To List1.ListCount - 1
        If ListBoxCheckDup(List2, List1.List(i)) = False Then
            List2.AddItem List1.List(i)
        End If
Next
End Sub

Private Sub unLd_Click()
Unload Me
End Sub

Private Sub warez_Click()

Call PrivateRoom("aviempire")
End Sub
