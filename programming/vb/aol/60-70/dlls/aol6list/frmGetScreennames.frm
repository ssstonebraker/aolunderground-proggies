VERSION 5.00
Begin VB.Form frmGetScreennames 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Screennames Example Project"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Select an example"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2370
      Width           =   5535
      Begin VB.OptionButton optSelection 
         Caption         =   "Member Directory"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   7
         Top             =   240
         Width           =   1185
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Who's Chatting"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2430
         TabIndex        =   6
         Top             =   300
         Width           =   1635
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Chatroom"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   300
         Width           =   1185
      End
      Begin VB.OptionButton optSelection 
         Caption         =   "Sign On"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdGetScreennames 
      Caption         =   "Get Screennames"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   3150
      Width           =   5535
   End
   Begin VB.ListBox lstScreennames 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   240
      TabIndex        =   0
      Top             =   1050
      Width           =   5535
   End
   Begin VB.Label lblSnCount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4080
      TabIndex        =   9
      Top             =   3630
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "created by huma (h4ma@yahoo.com)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   3660
      Width           =   2925
   End
   Begin VB.Label lblStatus 
      Caption         =   "information about the selected example"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   150
      TabIndex        =   2
      Top             =   90
      Width           =   5715
   End
End
Attribute VB_Name = "frmGetScreennames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmGetScreennames

Option Explicit

'###################################################################
' This class object was created in Visual Basic 6 SP4 EE with a 1024 x 768 resolution monitor
' using the verdana font, size 9.
'
' This module demonstrates various usages for the AOL6LIST ActiveX DLL. This project requires a
' reference to the AOL6LIST by Huma ActiveX DLL (AOL6LIST.dll) before you can run this project because
' early binding to the object in the AOL6LIST.dll is used. The CAOL class object that was included with
' this project is also required.
'
' This code and its entirety is provided "AS IS" with no warranties of any kind.
' If you would like to distribute this code, the module information as well as the top portion of this
' object must remain intact. I can be reached at h4ma@yahoo.com/http://welcome.to/huma/
'###################################################################
' Modifications:
'
' 1.00  03/31/01    Created by Nai
'###################################################################

'module information
Private Const csName = "frmGetScreennames"
Private Const csVersion = "1.00"
Private Const csDate = "03/31/01"
Private Const csAuthor = "Nai better known as Huma"

Private AOL As CAOL 'binds to the CAOL object, module level usage.

'API Function
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Sub GetScreennamesfromChatroom()
    Dim clsAol As CAol6List 'early bind to the object in AOL6LIST
    Dim index As Long 'current item index to iteriate through the screen name list
    
    'the variables below hold the handles to various windows on aol
    Dim chathWnd As Long
    Dim aolListhWnd As Long
    
    'create the object from the aol6list.dll so we can work with it
    Set clsAol = New CAol6List
    
        With AOL
            'return the window handle to the current chatroom
            chathWnd = .CurrentChatRoom
            'return the window handle to the aol listbox from the chathwnd
            aolListhWnd = .FindChildByClass(chathWnd, "_AOL_Listbox")
            
            'grab the screen names. only continue if it is successful
            If clsAol.GetScreennamesFromList(aolListhWnd, lteListbox) Then
            
                'clear out the screennames from the listbox
                lstScreennames.Clear
                
                'iteriate through each item in the list and add them to the listbox
                For index = 0 To clsAol.ListCount - 1
                    lstScreennames.AddItem clsAol.Item(index)
                Next
                
            End If
            
            'display the number of screen names
            lblSnCount = "screen names: " & lstScreennames.ListCount
        End With
    
    'destroy the object since we're finished with it
    Set clsAol = Nothing
End Sub

Private Sub GetScreennamesfromSignon()
    Dim clsAol As CAol6List 'early bind to the object in AOL6LIST
    Dim index As Long 'current item index to iteriate through the screen name list
    
    'the variables below hold the handles to various windows on aol
    Dim aolhWnd As Long
    Dim mdihWnd As Long
    Dim signonhWnd As Long
    Dim aolcombohWnd As Long
    
    'create the object from the aol6list.dll so we can work with it
    Set clsAol = New CAol6List
    
        With AOL
        
            'find the handle to the combobox on the signon window
            aolhWnd = FindWindow("AOL Frame25", vbNullString)
            mdihWnd = .FindChildByClass(aolhWnd, "MDIClient")
            signonhWnd = .FindChildByTitle(mdihWnd, "Sign On")
            aolcombohWnd = .FindChildByClass(signonhWnd, "_AOL_Combobox")
        
            'grab the screen names from the combobox. only continue if it is successful
            If clsAol.GetScreennamesFromList(aolcombohWnd, lteCombobox) Then
            
                'clear out the screennames from the listbox
                lstScreennames.Clear
                
                'iteriate through each item in the list and add them to the listbox
                For index = 0 To clsAol.ListCount - 1
                    lstScreennames.AddItem clsAol.Item(index)
                Next

            End If
            
            'display the number of screen names
            lblSnCount = "screen names: " & lstScreennames.ListCount
                
        End With
    
    'destroy the object since we're finished with it
    Set clsAol = Nothing

End Sub

Private Sub GetScreennamesfromWhosChatting()
    Dim clsAol As CAol6List 'early bind to the object in AOL6LIST
    Dim index As Long 'current item index to iteriate through the screen name list
    
    'the variables below hold the handles to various windows on aol
    Dim aolhWnd As Long
    Dim mdihWnd As Long
    Dim chattinghWnd As Long
    Dim aolListhWnd As Long
    
    'create the object from the aol6list.dll so we can work with it
    Set clsAol = New CAol6List
    
        With AOL
        
            'find the handle to the combobox on the signon window
            aolhWnd = FindWindow("AOL Frame25", vbNullString)
            mdihWnd = .FindChildByClass(aolhWnd, "MDIClient")
            chattinghWnd = .FindChildByTitle(mdihWnd, "Who's Chatting")
            aolListhWnd = .FindChildByClass(chattinghWnd, "_AOL_Listbox")
        
            'grab the screen names. only continue if it is successful
            If clsAol.GetScreennamesFromList(aolListhWnd, lteListbox) Then
            
                'clear out the screennames from the listbox
                lstScreennames.Clear
                
                'iteriate through each item in the list and add them to the listbox
                For index = 0 To clsAol.ListCount - 1
                    lstScreennames.AddItem clsAol.Item(index)
                Next
                
            End If
            
            'display the number of screen names
            lblSnCount = "screen names: " & lstScreennames.ListCount
        
        End With
    
    'destroy the object since we're finished with it
    Set clsAol = Nothing

End Sub

Private Sub GetScreennamesfromMemberDirectory()
    Dim clsAol As CAol6List 'early bind to the object in AOL6LIST
    Dim index As Long 'current item index to iteriate through the screen name list
    Dim lpos As Integer 'position of the tab character in the returned screenname
    Dim tmpSn As String 'temporary variable to hold the screen name
    
    'the variables below hold the handles to various windows on aol
    Dim aolhWnd As Long
    Dim mdihWnd As Long
    Dim memdirSRhWnd As Long
    Dim aolListhWnd As Long
    
    'create the object from the aol6list.dll so we can work with it
    Set clsAol = New CAol6List
    
        With AOL
        
            'find the handle to the combobox on the signon window
            aolhWnd = FindWindow("AOL Frame25", vbNullString)
            mdihWnd = .FindChildByClass(aolhWnd, "MDIClient")
            memdirSRhWnd = .FindChildByTitle(mdihWnd, "Member Directory Search Results")
            aolListhWnd = .FindChildByClass(memdirSRhWnd, "_AOL_Listbox")
        
            'grab the screen names from the list. only continue if it is successful
            If clsAol.GetScreennamesFromList(aolListhWnd, lteListbox) Then
            
                'clear out the screennames from the listbox
                lstScreennames.Clear
                
                    'iteriate through each item in the list and add them to the listbox
                    For index = 0 To clsAol.ListCount - 1
                        'since the screen names that are returned from the member directory search results
                        'are not formatted, we must format it before displaying it on the list
                        
                        tmpSn = clsAol.Item(index) 'place a copy of the screen name so we can manipulate it
                    
                        'return the position of the first tab character from the current screen name in the array
                        lpos = InStr(tmpSn, vbTab)
                        If lpos > 0 Then
                            'set the current screenname without the first tab character we found
                            tmpSn = Mid$(tmpSn, lpos + 1)
                        End If
                        
                        'now we return the position of the next tab character from the current screen name in the array
                        lpos = InStr(tmpSn, vbTab)
                        If lpos > 0 Then
                            'set the current screenname with everything left of the tab character we found
                            tmpSn = Mid$(tmpSn, 1, lpos - 1)
                        End If
                        
                        'add the formatted screen name into the list
                        lstScreennames.AddItem tmpSn
                    Next
                
            End If
            
            'display the number of screen names
            lblSnCount = "screen names: " & lstScreennames.ListCount
        
        End With
    
    'destroy the object since we're finished with it
    Set clsAol = Nothing

End Sub

Private Sub cmdGetScreennames_Click()
    'calls a method based on the selected option
    
    If optSelection.Item(0).Value Then
    
        Call GetScreennamesfromSignon
            
    ElseIf optSelection.Item(1).Value Then
    
        Call GetScreennamesfromChatroom
    
    ElseIf optSelection.Item(2).Value Then
    
        Call GetScreennamesfromWhosChatting
    
    ElseIf optSelection.Item(3).Value Then
    
        Call GetScreennamesfromMemberDirectory

    End If
End Sub

Private Sub Form_Load()
    'create the object so we can work with it
    Set AOL = New CAOL
    
    Call optSelection_Click(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'cleanup by destroying the object before exiting
    Set AOL = Nothing
End Sub

Private Sub optSelection_Click(index As Integer)
    'display information about the selected option
    
    Select Case index
        Case 0
            lblStatus.Caption = "This option allows you to grab the screen names off AOL 6's Sign On window. " & _
                "Make sure AOL 6 is opened prior to running this example."
        Case 1
            lblStatus.Caption = "This option allows you to grab the screen names off the chat room you " & _
                "are currently in, on AOL 6. Make sure you are in a chat room prior to running this example."
        Case 2
            lblStatus.Caption = "This option allows you to grab the screen names off the 'Who's Chatting' window. " & _
                "Make sure you have the Who's Chatting window visible prior to running this example."
        Case 3
            lblStatus.Caption = "This option allows you to grab the screen names off the Member Directory Search " & _
                "Results window. Make sure the Member Directory Search Results window is visible prior to running " & _
                "this example."
    End Select
End Sub
