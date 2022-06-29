VERSION 5.00
Begin VB.Form LB_Search 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ListBox By NightShade"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear ListBox"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoadList 
      Caption         =   "&Load List"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdFindString 
      Caption         =   "&Find String"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text to search for..."
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "LB_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoadList_Click()
    
    'loads text into your listbox

    Dim IndexNumber As Integer
    
    For IndexNumber% = 0 To 999
        List1.AddItem "Listbox entry #" & IndexNumber%
    Next IndexNumber%

End Sub

Private Sub cmdFindString_Click()

    'This example shows a fast search for a string in a
    'listbox.
    'NightShade 05/28/99

    Dim LBHandle As Long
    Dim ResultIndex As Long
    Dim LookinFor As String
    
    LookinFor$ = Text1.Text 'what text are you searching for?

    LBHandle& = List1.hwnd  'gets the listbox's handle
    
    'Varable     = API call  (Listbox id#, all lowercase,(search for string)
    ResultIndex& = SendMessageByString(LBHandle&, LCase(LB_FINDSTRINGEXACT), -1, LCase(LookinFor$))
             
        If ResultIndex& < 0 Then
            MsgBox "Your string was not found"
        Else
            MsgBox "ListBox entry found at index number: " & ResultIndex&
        End If
        

End Sub

Private Sub cmdClear_Click()

    List1.Clear 'clears the listbox

End Sub


Private Sub cmdExit_Click()

    End
    
End Sub

