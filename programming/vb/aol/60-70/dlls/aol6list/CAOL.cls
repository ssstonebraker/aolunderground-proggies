VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CAOL class object

'###################################################################
' This class object was created in Visual Basic 6 SP4 EE with a 1024 x 768 resolution monitor
' using the verdana font, size 9.
'
' This object handles a few basic functionalities for AOL
'
' This code and its entirety is provided "AS IS" with no warranties of any kind.
' If you would like to distribute this code, the class information as well as the top portion of this
' object must remain intact. I can be reached at h4ma@yahoo.com/http://welcome.to/huma/
'###################################################################
' Modifications:
'   1.00    10/11/00    Created CAOL class module
'   1.01    03/06/01    Added isOnline, GetAolVersion, GetScreenName
'   1.02    03/31/01    Most of the methods from the CAOL class object were removed and optimized for
'                                 public distribution
'###################################################################

Option Explicit

'class information
Private Const m_sName = "CAOL"
Private Const m_sVersion = "1.02"
Private Const m_sDate = "03/31/01"
Private Const m_sAuthor = "Nai better known as Huma"

'API Declarations
Private Const GW_HWNDNEXT = 2
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

'The findchildbytitle and findchildbyclass are wrapped from the FindWindowEx function so that
'it's shorter and i've grown accustomed to it from vb3

'finds a window handle by a parent handle and the title text from a class
Public Property Get FindChildByTitle(parenthWnd As Long, childhWnd As String) As Long
    FindChildByTitle = FindWindowEx(parenthWnd, 0, vbNullString, childhWnd)
End Property

'finds a window handle by a parent handle and the class name
Public Property Get FindChildByClass(parenthWnd As Long, childhWnd As String) As Long
    FindChildByClass = FindWindowEx(parenthWnd, 0, childhWnd, vbNullString)
End Property

'Goes to the next specified handle from the defined handle until the count has been met
Private Sub GethWndByNum(hWnd As Long, Optional Count As Long = 1)
    Dim dwEnum As Long
    Dim eventCnt As Long
    
    Do
        'we use a doevents every 8 times in the loop to prevent this application from looking as though it is frozen
        If eventCnt >= 8 Then
            DoEvents
            eventCnt = 0
        Else
            'increment the value
            eventCnt = eventCnt + 1
        End If
        
        dwEnum = dwEnum + 1
        'get the handle of the next window
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
        
    Loop Until dwEnum >= Count
End Sub

'retrieve the text from a window handle
Private Function GetWinText(hWnd As Long) As String
    Dim retString As String 'buffer for the window text
    
    On Error Resume Next
    
    'return the length of the string. then create a string buffer from the length
    retString = Space(SendMessage(hWnd, WM_GETTEXTLENGTH, 0, 0))
    
        'retrieve the text into the buffer
        Call SendMessageByString(hWnd, WM_GETTEXT, Len(retString) + 1, ByVal retString)
        
    GetWinText = vbNullString & retString
End Function

'returns the window handle from the current chat room even if the chatroom is not in focus
Public Function CurrentChatRoom() As Long
    'the variables below hold the handles to various windows on aol
    Dim AOL As Long
    Dim MdiClient As Long
    Dim AolChild As Long
    Dim AOLStatic As Long
    Dim AolListbox As Long

    Dim i As Long 'current index in a for/next statement
    Dim childCount As Long 'current index in a for/next statement
    Dim cStatic As String 'holds the text from an AOL Static class
       
    On Error GoTo Err_CurrentChatRoom
    
    'find the handle to an Aol Child class on AOL
    AOL = FindWindow("AOL Frame25", vbNullString)
    MdiClient = FindChildByClass(AOL, "MDIClient")
    
    AolChild = FindChildByClass(MdiClient, "AOL Child")
    
    'iteriate through the aolchild windows 25 times to look for the chat room incase
    'the chatroom isn't the window that isn't in focus. just hope the chatroom isn't more than 25 levels deep
    For childCount = 1 To 25
            
            'return the handle to the listbox
            AolListbox = FindChildByClass(AolChild, "_AOL_Listbox")
            
            If AolListbox > 0 Then
                
                'return the handle to the aol static class
                AOLStatic = FindChildByClass(AolChild, "_AOL_Static")
                
                'iteriate 25 times through the aol static classes
                For i = 1 To 25

                    If AOLStatic > 0 Then
                        
                        'return the text from the aol static class
                        cStatic = LCase$(GetWinText(AOLStatic))
                        
                        'make sure the aol static is "people here" which ensures that the aol child
                        'handle is correct
                        If InStr(cStatic, "people here") > 0 Then
                            'return the handle to the aol child class and exit this method
                            CurrentChatRoom = AolChild
                            GoTo Exit_CurrentChatRoom
                        End If
                    
                    Else
                        'since no aol static class was found, we longer need to iteriate through the
                        'for/next statement so we continue to the next aol child
                        Exit For
                    End If
                    
                    'go to the next static on the current aol child class
                    GethWndByNum AOLStatic
                    
                Next
                
            End If
        
        'go to the next aol child on AOL
        GethWndByNum AolChild
    Next
    
Exit_CurrentChatRoom:

    Exit Function

Err_CurrentChatRoom:

    If Err <> 0 Then
        MsgBox "An error has occurred. Error " & Err & ". " & Error$
    End If
    
End Function

'returns the current chatroom's name
Public Function CurrentChatroomName() As String
    On Error Resume Next
    
    CurrentChatroomName = GetWinText(CurrentChatRoom)
End Function
