VERSION 5.00
Begin VB.UserControl Chat 
   CanGetFocus     =   0   'False
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   930
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   255
   ScaleWidth      =   930
   Begin VB.Timer Room 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zer0 Effect"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event scan(ByVal sn As String, ByVal Msg As String)
Dim StopIt As Boolean
Dim TxtOld As String
Dim TxtOldShort As String
'okay let me start this off by saying thanx to
'the members of PAA and their chat scan example.
'i was able to see the errors in my last chat scan
'and i was able to see the errors in theirs to
'create this perfect chat scan.
'(perfect is probably too strong of a word, but o well)
'
'i cannot take full credit for this scan, because
'of the item mentioned above.
'
'this was beta tested by Akuma & me
'hopefully this will be the last version, but if there
'are any errors... i'll fix those and send it to pat

Private Sub UserControl_Initialize()
    UserControl.Width = Label1.Width
    UserControl.Height = Label1.Height
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = Label1.Width
    UserControl.Height = Label1.Height
End Sub

Private Sub Room_Timer()
            'get the new text
               If ChatBox = 0 Then Exit Sub
                'if the new text is like the old text then we don't need it, so exit the sub
            HandleLastLine
            'if the new text isn't like the old text then HandleLastLine will get what we need
End Sub


Private Sub HandleLastLine()
'i'm sure there's a way to speed up this process even more, but i can't think of it....
'this is where a mixture of my coding and PAA's combines
'i use a sub instead of a timer to sort the code
'i go through the lines 1 by 1 w/ a function from my module
'PAA's way was good kuz it didn't take all the chat
'but it had too many if then statements that slowed it down
'and on top of that they had all the coding in the timer
'even dos stated that running from a sub, rather than a timer,
'is much more effective.

        On Error GoTo Top:
       
    Room.Enabled = False
    Dim LineCount As Integer
    Dim CombnTxt As String
    Dim Text As String
    Dim ChatMsg As String
    Dim ChatSN As String
    Dim Txt As String
Top:
DoEvents
        Txt = GetLastChatLine
If TxtOld = Txt Then GoTo Top:
    
    
    LineCount = Line_Count(GetLastChatLine)
   
    If LineCount = 0 Then Exit Sub
  
        
        If StopIt = True Then Exit Sub
        
        Text = GetLastChatLine

        
        ChatMsg = GetLastMSG(GetLastChatLine)
       
        ChatSN = GetLastSN(GetLastChatLine)
      
        If ChatSN <> "" Then
        
             RaiseEvent scan(ChatSN, ChatMsg)
        
        End If
       
        TxtOld = Txt

    Room.Enabled = True
   
End Sub
Public Sub ScanOff()
TxtOld = ""
    StopIt = True
    Room.Enabled = False
End Sub

Public Sub ScanOn()
TxtOld = ""
    StopIt = False
    Room.Enabled = True
End Sub
Public Sub ChatSend(Message As String)
Call SendChat(Message)
End Sub
