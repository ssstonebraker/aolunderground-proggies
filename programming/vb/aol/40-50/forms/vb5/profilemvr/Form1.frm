VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Profile Transfer"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   7650
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7275
   ScaleWidth      =   7650
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6000
      TabIndex        =   34
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      Height          =   255
      Left            =   6600
      TabIndex        =   32
      Top             =   5760
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Underlined"
      Height          =   255
      Left            =   6240
      TabIndex        =   31
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Italics"
      Height          =   255
      Left            =   6240
      TabIndex        =   30
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Bold"
      Height          =   255
      Left            =   6240
      TabIndex        =   29
      Top             =   6240
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5400
      TabIndex        =   28
      Top             =   6240
      Width           =   615
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   4080
      TabIndex        =   24
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2520
      TabIndex        =   23
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Generate Code"
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   6960
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   21
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add Color"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear ^^^ Profile"
      Height          =   255
      Left            =   4800
      TabIndex        =   19
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create New Profile"
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Profile"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   4440
      Width           =   6495
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3840
      Width           =   6495
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3240
      Width           =   6495
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2160
      Width           =   6495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   1800
      Width           =   6495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000008&
      Height          =   525
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Profile"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Gender:"
      Height          =   255
      Left            =   4920
      TabIndex        =   33
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Font Size:"
      Height          =   255
      Left            =   4560
      TabIndex        =   27
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "NOTE: If font and font size fields left blank, the AOL Default will be used."
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   6600
      Width           =   5175
   End
   Begin VB.Label Label9 
      Caption         =   "Font:"
      Height          =   255
      Left            =   1680
      TabIndex        =   25
      Top             =   6240
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7680
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label8 
      Caption         =   "Quote:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Career:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Computer:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Hobbies:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Marital:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Birth:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim aol As Long, mdi As Long, Quick As Long, Advanced As Long, Mine As Long, Search As Long
Dim Neighbor As Long, Ok As Long, Check As Long, EditPro As Long, Member As Long, GenderF As Long
Dim Birth As Long, Location As Long, Hobbies As Long, Computer As Long, Quote As Long, GenderN As Long
Dim Career As Long, Marry As Long, Male As Long, Female As Long, None As Long, GenderM As Long
Dim BGColor As Long, FontColor As Long

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
EditPro& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit Your Online Profile")
 If EditPro& = 0 Then
  aol& = FindWindow("AOL Frame25", vbNullString)
  mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
  Search& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Directory")
   If Search& = 0 Then
    Call KeyWord("Profile")
    Do
     Search& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Directory")
    Loop Until Search& > 0
   End If
  Pause (0.7)
  Mine& = FindWindowEx(Search&, 0&, "_AOL_Icon", vbNullString)
  Call SendMessage(Mine&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Mine&, WM_LBUTTONUP, 0&, 0&)
  Call SendMessage(Mine&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Mine&, WM_LBUTTONUP, 0&, 0&)
  EditPro& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit Your Online Profile")
   Do
    EditPro& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit Your Online Profile")
   Loop Until EditPro& > 0
  Pause (1.3)
  Neighbor& = FindWindow("_AOL_Modal", vbNullString)
  Check& = FindWindowEx(Neighbor&, 0&, "_AOL_Checkbox", vbNullString)
  Ok& = FindWindowEx(Neighbor&, 0&, "_AOL_icon", vbNullString)
   If Neighbor& > 0 Then
    Call SendMessage(Check&, BM_SETCHECK, True, vbNullString)
    Call SendMessage(Ok&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Ok&, WM_LBUTTONUP, 0&, 0&)
    Pause (0.7)
   End If
 End If
Member& = FindWindowEx(EditPro&, 0&, "_AOL_Edit", vbNullString)
Location& = FindWindowEx(EditPro&, Member&, "_AOL_Edit", vbNullString)
Birth& = FindWindowEx(EditPro&, Location&, "_AOL_Edit", vbNullString)
Male& = FindWindowEx(EditPro&, 0&, "_AOL_Checkbox", vbNullString)
Female = FindWindowEx(EditPro&, Male&, "_AOL_Checkbox", vbNullString)
None& = FindWindowEx(EditPro&, Female&, "_AOL_Checkbox", vbNullString)
Marry& = FindWindowEx(EditPro&, Birth&, "_AOL_Edit", vbNullString)
Hobbies& = FindWindowEx(EditPro&, Marry&, "_AOL_Edit", vbNullString)
Computer& = FindWindowEx(EditPro&, Hobbies&, "_AOL_Edit", vbNullString)
Career& = FindWindowEx(EditPro&, Computer&, "_AOL_Edit", vbNullString)
Quote& = FindWindowEx(EditPro&, Career&, "_AOL_Edit", vbNullString)
Text1 = GetText(Member&)
Text2 = GetText(Location&)
Text3 = GetText(Birth&)
GenderM& = SendMessage(Male&, BM_GETCHECK, 0&, 0&)
GenderF& = SendMessage(Female&, BM_GETCHECK, 0&, 0&)
GenderN& = SendMessage(None&, BM_GETCHECK, 0&, 0&)
 If GenderM& = 1 Then
  Combo3.text = "Male"
 End If
 If GenderF& = 1 Then
  Combo3.text = "Female"
 End If
 If GenderN& = 1 Then
  Combo3.text = ""
 End If
Text4 = GetText(Marry&)
Text5 = GetText(Hobbies&)
Text6 = GetText(Computer&)
Text7 = GetText(Career&)
Text8 = GetText(Quote&)
Pause (0.5)
Call SendMessage(EditPro&, WM_CLOSE, 0&, 0&)
Call SendMessage(Search&, WM_CLOSE, 0&, 0&)
End Sub

Private Sub Command2_Click()
Dim aol As Long, mdi As Long, Quick As Long, Advanced As Long, Mine As Long, Search As Long
Dim Neighbor As Long, OkBox1 As Long, Check As Long, EditPro As Long, Update As Long, OKButton As Long
Dim Delete As Long, Site As Long, OkBox As Long, Yes As Long, No As Long, Ok As Long

Reply = MsgBox("Are you sure you want to delete your profile?", 36, "Delete Profile")
 If Reply = vbNo Then
  Exit Sub
 End If
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
EditPro& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit Your Online Profile")
 If EditPro& = 0 Then
  aol& = FindWindow("AOL Frame25", vbNullString)
  mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
  Search& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Directory")
   If Search& = 0 Then
    Call KeyWord("Profile")
    Do
     Search& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Directory")
    Loop Until Search& > 0
   End If
  Pause (0.7)
  Mine& = FindWindowEx(Search&, 0&, "_AOL_Icon", vbNullString)
  Call SendMessage(Mine&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Mine&, WM_LBUTTONUP, 0&, 0&)
  Call SendMessage(Mine&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Mine&, WM_LBUTTONUP, 0&, 0&)
  EditPro& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit Your Online Profile")
   Do
    EditPro& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit Your Online Profile")
   Loop Until EditPro& > 0
  Pause (1.3)
  Neighbor& = FindWindow("_AOL_Modal", vbNullString)
  Check& = FindWindowEx(Neighbor&, 0&, "_AOL_Checkbox", vbNullString)
  Ok& = FindWindowEx(Neighbor&, 0&, "_AOL_icon", vbNullString)
   If Neighbor& > 0 Then
    Call SendMessage(Check&, BM_SETCHECK, True, vbNullString)
    Call SendMessage(Ok&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Ok&, WM_LBUTTONUP, 0&, 0&)
    Pause (0.7)
   End If
 End If
Site& = FindWindowEx(EditPro&, 0&, "_AOL_Icon", vbNullString)
Update& = FindWindowEx(EditPro&, Site&, "_AOL_Icon", vbNullString)
Delete& = FindWindowEx(EditPro&, Update&, "_AOL_Icon", vbNullString)
Call SendMessage(Delete&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Delete&, WM_LBUTTONUP, 0&, 0&)
Call SendMessage(Delete&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Delete&, WM_LBUTTONUP, 0&, 0&)
 Do
  OkBox& = FindWindow("_AOL_Modal", vbNullString)
 Loop Until OkBox& > 0
Pause (0.5)
OkBox& = FindWindow("_AOL_Modal", vbNullString)
No& = FindWindowEx(OkBox&, 0&, "_AOL_Icon", vbNullString)
Yes& = FindWindowEx(OkBox&, No&, "_AOL_Icon", vbNullString)
Call SendMessage(Yes&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Yes&, WM_LBUTTONUP, 0&, 0&)
Call SendMessage(Yes&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Yes&, WM_LBUTTONUP, 0&, 0&)
 Do
  OkBox1& = FindWindow("#32770", "America Online")
 Loop Until OkBox1& > 0
Pause (0.5)
OkBox1& = FindWindow("#32770", "America Online")
OKButton& = FindWindowEx(OkBox1&, 0&, "Button", vbNullString)
Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Pause (0.5)
Call SendMessage(EditPro&, WM_CLOSE, 0&, 0&)
Call SendMessage(Search&, WM_CLOSE, 0&, 0&)
End Sub

Private Sub Command3_Click()
Dim aol As Long, mdi As Long, Quick As Long, Advanced As Long, Mine As Long, Search As Long, Female As Long
Dim Neighbor As Long, Ok As Long, Check As Long, EditPro As Long, Member As Long, MemString As String
Dim Birth As Long, Location As Long, Hobbies As Long, Computer As Long, Quote As Long, Male As Long
Dim Career As Long, Marry As Long, OkBox1 As Long, OKButton As Long, Site As Long, Update As Long, None As Long

If Text1 = "" And Text2 = "" And Text3 = "" And Text4 = "" And Text5 = "" And Text6 = "" And Text7 = "" And Text8 = "" And Text9 = "" Then
 MsgBox "There is no new profile to create. Please choose color, make a new profile, or copy your current one.", vbCritical, ""
End If
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
EditPro& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit Your Online Profile")
 If EditPro& = 0 Then
  aol& = FindWindow("AOL Frame25", vbNullString)
  mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
  Search& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Directory")
   If Search& = 0 Then
    Call KeyWord("Profile")
    Do
     Search& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Directory")
    Loop Until Search& > 0
   End If
  Pause (0.7)
  Mine& = FindWindowEx(Search&, 0&, "_AOL_Icon", vbNullString)
  Call SendMessage(Mine&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Mine&, WM_LBUTTONUP, 0&, 0&)
  Call SendMessage(Mine&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Mine&, WM_LBUTTONUP, 0&, 0&)
  EditPro& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit Your Online Profile")
   Do
    EditPro& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit Your Online Profile")
   Loop Until EditPro& > 0
  Pause (1.3)
  Neighbor& = FindWindow("_AOL_Modal", vbNullString)
  Check& = FindWindowEx(Neighbor&, 0&, "_AOL_Checkbox", vbNullString)
  Ok& = FindWindowEx(Neighbor&, 0&, "_AOL_icon", vbNullString)
   If Neighbor& > 0 Then
    Call SendMessage(Check&, BM_SETCHECK, True, vbNullString)
    Call SendMessage(Ok&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Ok&, WM_LBUTTONUP, 0&, 0&)
    Pause (0.5)
   End If
 End If
Member& = FindWindowEx(EditPro&, 0&, "_AOL_Edit", vbNullString)
Location& = FindWindowEx(EditPro&, Member&, "_AOL_Edit", vbNullString)
Birth& = FindWindowEx(EditPro&, Location&, "_AOL_Edit", vbNullString)
Male& = FindWindowEx(EditPro&, 0&, "_AOL_Checkbox", vbNullString)
Female& = FindWindowEx(EditPro&, Male&, "_AOL_Checkbox", vbNullString)
None& = FindWindowEx(EditPro&, Female&, "_AOL_Checkbox", vbNullString)
Marry& = FindWindowEx(EditPro&, Birth&, "_AOL_Edit", vbNullString)
Hobbies& = FindWindowEx(EditPro&, Marry&, "_AOL_Edit", vbNullString)
Computer& = FindWindowEx(EditPro&, Hobbies&, "_AOL_Edit", vbNullString)
Career& = FindWindowEx(EditPro&, Computer&, "_AOL_Edit", vbNullString)
Quote& = FindWindowEx(EditPro&, Career&, "_AOL_Edit", vbNullString)
If Text1 > "" And Text9 = "" Then
 Call SendMessageByString(Member&, WM_SETTEXT, 0&, Text1.text)
End If
If Text1 = " " And Text9 = "" Then
 Call SendMessageByString(Member&, WM_SETTEXT, 0&, "")
End If
If Text1 = " " And Text9 > "" Then
 Call SendMessageByString(Member&, WM_SETTEXT, 0&, Text9.text)
End If
If Text1 > "" And Text9 > "" Then
 Call SendMessageByString(Member&, WM_SETTEXT, 0&, Text9.text + Text1.text)
End If
If Text1 = "" And Text9 > "" Then
 MemString$ = GetText(Member&)
 Call SendMessageByString(Member&, WM_SETTEXT, 0&, Text9.text + MemString$)
End If
If Text2 = " " Then
 Call SendMessageByString(Location&, WM_SETTEXT, 0&, "")
End If
If Text2 > "" Then
 Call SendMessageByString(Location&, WM_SETTEXT, 0&, Text2.text)
End If
If Text3 = " " Then
 Call SendMessageByString(Birth&, WM_SETTEXT, 0&, "")
End If
If Text3 > "" Then
 Call SendMessageByString(Birth&, WM_SETTEXT, 0&, Text3.text)
End If
If Combo3.text = "Male" Then
 Call SendMessage(Male&, BM_SETCHECK, True, vbNullString)
End If
If Combo3.text = "Female" Then
 Call SendMessage(Female&, BM_SETCHECK, True, vbNullString)
End If
If Combo3.text = "" Then
 Call SendMessage(None&, BM_SETCHECK, True, vbNullString)
End If
If Text4 = " " Then
 Call SendMessageByString(Marry&, WM_SETTEXT, 0&, "")
End If
If Text4 > "" Then
 Call SendMessageByString(Marry&, WM_SETTEXT, 0&, Text4.text)
End If
If Text5 = " " Then
 Call SendMessageByString(Hobbies&, WM_SETTEXT, 0&, "")
End If
If Text5 > "" Then
 Call SendMessageByString(Hobbies&, WM_SETTEXT, 0&, Text5.text)
End If
If Text6 = " " Then
 Call SendMessageByString(Computer&, WM_SETTEXT, 0&, "")
End If
If Text6 > "" Then
 Call SendMessageByString(Computer&, WM_SETTEXT, 0&, Text6.text)
End If
If Text7 = " " Then
 Call SendMessageByString(Career&, WM_SETTEXT, 0&, "")
End If
If Text7 > "" Then
 Call SendMessageByString(Career&, WM_SETTEXT, 0&, Text7.text)
End If
If Text8 = " " Then
 Call SendMessageByString(Quote&, WM_SETTEXT, 0&, "")
End If
If Text8 > "" Then
 Call SendMessageByString(Quote&, WM_SETTEXT, 0&, Text8.text)
End If
Pause (0.8)
Site& = FindWindowEx(EditPro&, 0&, "_AOL_Icon", vbNullString)
Update& = FindWindowEx(EditPro&, Site&, "_AOL_Icon", vbNullString)
Call SendMessage(Update&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Update&, WM_LBUTTONUP, 0&, 0&)
Call SendMessage(Update&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Update&, WM_LBUTTONUP, 0&, 0&)
 Do
  OkBox1& = FindWindow("#32770", "America Online")
 Loop Until OkBox1& > 0
Pause (0.5)
OkBox1& = FindWindow("#32770", "America Online")
OKButton& = FindWindowEx(OkBox1&, 0&, "Button", vbNullString)
Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Pause (0.5)
Call SendMessage(EditPro&, WM_CLOSE, 0&, 0&)
Call SendMessage(Search&, WM_CLOSE, 0&, 0&)
End Sub

Private Sub Command4_Click()
 If Text1 = "" And Text2 = "" And Text3 = "" And Text4 = "" And Text5 = "" And Text6 = "" And Text7 = "" And Text8 = "" Then
  Exit Sub
 Else
  Reply = MsgBox("Are you sure you want to clear the copied profile?", 36, "Delete Profile")
   If Reply = vbNo Then
    Exit Sub
   Else
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text8 = ""
    Combo3.text = ""
   End If
 End If
End Sub

Private Sub Command5_Click()
 If LCase(Text1.text) Like LCase("<<u>*") Then
  MsgBox "Please delete current code from the member name field before using this option", vbCritical, "Code"
  Exit Sub
 End If
 If Text1 = "" And Text2 = "" And Text3 = "" And Text4 = "" And Text5 = "" And Text6 = "" And Text8 = "" Then
  Reply = MsgBox("There is no profile loaded. Would you like to add color without transfering?", 36, "Color")
   If Reply = vbNo Then
    MsgBox "Please load profile first. This will allow you to see the color of the background and text before applying changes", 0, "Load Profile"
    Exit Sub
   Else
    Color.Show
   End If
 End If
Color.Show
End Sub

Private Sub Command6_Click()
If Text10 = "" And Combo1.text = "" And Combo2.text = "" And Check1.Value = False And Check2.Value = False And Check3.Value = False Then
 MsgBox "There is no code to generate"
Else
 Text9 = ""
 Text9 = Text10
  If Combo1.text > "" Then
   Text9 = Text9 & "<<u>Font Face=" & Chr(34) & Combo1.text & Chr(34) & ">"
  End If
  If Combo2.text > "" Then
   Text9 = Text9 & "<<u>Font PTSize=" & Chr(34) & Combo2.text & Chr(34) & ">"
  End If
  If Check1.Value = 1 Then
   Text9 = Text9 & "<<u>b>"
  End If
  If Check2.Value = 1 Then
   Text9 = Text9 & "<<u>i>"
  End If
  If Check3.Value = 1 Then
   Text9 = Text9 & "<<u>u>"
  End If
 MsgBox "Code has been generated. Now just press Create New Profile to add coding", 0, "Coding Generated"
End If
End Sub

Private Sub Command7_Click()
MsgBox "This is quite simple really. All you have to do is copy your list then go to your new screen name and create a new. If you are not moving your list and want to add color, just go to the " & Chr(34) & "Add Color" & Chr(34) & " option and follow the simple instruction." & Chr(10) & Chr(13) & "There is only one catch to adding color. Some coding is quite long and can be complex. AOL often does not allow enough room for all these coding PLUS all the extra stuff. When adding color, make the member name fairly short. This will allow you to put in the ENITRE code PLUS all the extra. If the member name is not short enough, all the code should fit, but the name itself might not.   =)" & Chr(13) & Chr(10) & "If you want to leave a subject unchanged after copying the profile, Then either leave the subject field alone or leave it blank. If you want to delete a subject, put a single space in that field and the program will do the rest", 0, "Profile Editor/Transfer"
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Form_Load()
FormOnTop Me
Combo3.AddItem ""
Combo3.AddItem "Male"
Combo3.AddItem "Female"
Combo1.AddItem ""
Combo1.AddItem "Abadi MT Condensed"
Combo1.AddItem "Arial"
Combo1.AddItem "Arial Black"
Combo1.AddItem "Arial Narrow"
Combo1.AddItem "Bookman Old Style"
Combo1.AddItem "Calisto MT"
Combo1.AddItem "Comic Sans MS"
Combo1.AddItem "Courier New"
Combo1.AddItem "Garamond"
Combo1.AddItem "Geotype TT"
Combo1.AddItem "Impact"
Combo1.AddItem "Lucida Console"
Combo1.AddItem "Marlett"
Combo1.AddItem "MS Hei"
Combo1.AddItem "MS Song"
Combo1.AddItem "Symbol"
Combo1.AddItem "Tamoha"
Combo1.AddItem "Times New Roman"
Combo1.AddItem "Verdana"
Combo1.AddItem "Webdings"
Combo1.AddItem "Wingdings"
Combo2.AddItem ""
Combo2.AddItem "8"
Combo2.AddItem "10"
Combo2.AddItem "12"
Combo2.AddItem "14"
Combo2.AddItem "16"
Combo2.AddItem "18"
Combo2.AddItem "20"
Combo2.AddItem "22"
Combo2.AddItem "24"
Combo2.AddItem "26"
Combo2.AddItem "28"
Combo2.AddItem "30"
Combo2.AddItem "32"
Combo2.AddItem "34"
Combo2.AddItem "36"
Combo2.AddItem "38"
Combo2.AddItem "40"
Combo2.AddItem "42"
Combo2.AddItem "44"
Combo2.AddItem "46"
Combo2.AddItem "48"
Combo2.AddItem "50"
Combo2.AddItem "52"
Combo2.AddItem "54"
Combo2.AddItem "56"
Combo2.AddItem "58"
Combo2.AddItem "60"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
