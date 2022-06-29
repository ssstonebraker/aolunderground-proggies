VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buddy List Mover"
   ClientHeight    =   1200
   ClientLeft      =   1380
   ClientTop       =   1515
   ClientWidth     =   2730
   Height          =   1605
   Icon            =   "blistmove.frx":0000
   Left            =   1320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2730
   Top             =   1170
   Width           =   2850
   Begin VB.CommandButton Command5 
      Caption         =   "Clear this list->"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Help"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Buddy List"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Make new List"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.ListBox List1 
      ForeColor       =   &H00000000&
      Height          =   1230
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Buddy List"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Count As Long

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Butt% = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Call Click(Butt%)
TimeOut (1.3)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
but1% = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Call Click(but1%)
TimeOut (1.3)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
ourhandle& = FindWindowEx(OurParent&, 0, "_AOL_Listbox", vbNullString)
Count& = SendMessage(ourhandle&, LB_GETCOUNT, 0&, 0&)
Count = Count&

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
but2% = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Edit", vbNullString)
our% = FindWindowEx(OurParent&, Hand1&, "_AOL_Edit", vbNullString)
TimeOut (1)

For X = 0 To Count
Call Click(but2%)
addit = GetText(our%)
list1.AddItem addit
TimeOut (0.9)
Next X
TimeOut (0.5)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
ourparenta& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Call Window_Close(ourparenta&)
TimeOut (0.2)
i = list1.TopIndex
list1.RemoveItem i




End Sub


Private Sub Command2_Click()

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Butt% = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Call Click(Butt%)
TimeOut (1.3)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
but1% = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Call Click(but1%)
TimeOut (1.3)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Edit", vbNullString)
ourhandle& = FindWindowEx(OurParent&, Hand1&, "_AOL_Edit", vbNullString)

For X = 0 To list1.ListCount
Call SetFocusAPI(AORich%)
Call SendMessageByString(ourhandle&, WM_SETTEXT, 0, list1.List(X))
DoEvents
Call SendMessageByNum(ourhandle&, WM_CHAR, 13, 0)
TimeOut (0.6)
Next X
TimeOut (0.3)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
nig% = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Call Click(nig%)


End Sub


Private Sub Command3_Click()
Call RefreshBuddyList

End Sub


Private Sub Command4_Click()
MsgBox "Have you ever switched screen names and had to go through all kinds of trouble copying your old buddy list to your new one? Well this prog takes all the trouble out of that. You get on the screen name with the buddy list you want to copy, click 'copy buddy list' and it takes all the SN's and adds them to the list on this prog. Then go on your new SN and click 'Make new list' to add all the SN's from your origional buddy list to the new one. In order for it to work all your buddies must be in one category. (people who make all these seperate categories are lame) and you MUST have your buddy list window open.", 64, "Help"
End Sub


Private Sub Command5_Click()
List.Clear

End Sub

Private Sub Form_Load()
StayOnTop Me

End Sub


