Option Explicit
'=====================================================================
'                  Copyright 1993-1996 by Teletech Systems, Inc. All rights reserved
'
'
'This source code may not be distributed in part or as a whole without
'express written permission from Teletech Systems.
'=====================================================================
'
'This source code contains the following routines:
'  o SetAppHelp() 'Called in the main Form_Load event to register your
'                 'program with WINHELP.EXE
'  o QuitHelp()    'Deregisters your program with WINHELP.EXE. Should
'                  'be called in your main Form_Unload event
'  o ShowHelpTopic(Topicnum) 'Brings up context sensitive help based on
'                  'any of the following CONTEXT IDs
'  o ShowContents  'Displays the startup topic
'********** Shameless Plug <g> **********
'The Standard and Professional editions of HelpWriter
' also include the following routines to add sizzle to your
' helpfile presentation...
'  o HelpWindowSize(x,y,dx,dy) ' Position help window in a screen
'                              ' independent manner
'  o SearchHelp()  'Brings up the windows help KEYWORD SEARCH dialog box
'***********************************************************************
'
'=====================================================================
'List of Context IDs for <MMERHLP>
'=====================================================================
Global Const Hlp_SoLuTioN_MaiLing = 10    'Main Help Window
Global Const Hlp_Menu_Driven = 20    'Main Help Window
Global Const Hlp_Room_Buster = 30    'Main Help Window
Global Const Hlp_SiGN_oN = 40    'Main Help Window
Global Const Hlp_MaiLing_Options = 50    'Main Help Window
Global Const Hlp_MaiLiNg_StaTus = 60    'Main Help Window
Global Const GLOS_Instant_Messages_OnxOff = 70
Global Const GLOS_AOL_ToolBar = 80
Global Const GLOS_Exit = 90
Global Const GLOS_Notify_Additions = 100
Global Const GLOS_IM_Full_Boxes = 110
Global Const GLOS_SignOff_After = 120
Global Const GLOS_Start_From_Bottom = 130
Global Const GLOS_Keep_As_New = 140
Global Const GLOS_Mail_Current = 150
Global Const GLOS_Mail_All = 160
Global Const GLOS_Mail_Range = 170
Global Const GLOS_Mail_Selected = 180
Global Const GLOS_Mail_Alphabetically = 190
Global Const GLOS_Mailing_Message = 200
Global Const GLOS_Pause = 210
Global Const GLOS_Resume = 220
Global Const GLOS_Stop = 230
Global Const GLOS_RexAdd_Screen_Names = 240
'=====================================================================
'
'
'  Help engine section.

' Commands to pass WinHelp()
'Global Const HELP_CONTEXT = &H1 '  Display topic in ulTopic
'Global Const HELP_QUIT = &H2    '  Terminate help
'Global Const HELP_INDEX = &H3   '  Display index
'Global Const HELP_HELPONHELP = &H4      '  Display help on using help
'Global Const HELP_SETINDEX = &H5        '  Set the current Index for multi index help
'Global Const HELP_KEY = &H101           '  Display topic for keyword in offabData
'Global Const HELP_MULTIKEY = &H201
Global Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
Global Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
Global Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
Global Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
Global Const HELP_COMMAND = &H102    ' Execute Help macro
Global Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
Global Const HELP_SETWINPOS = &H203  ' Display and position Help window

Type HELPWININFO
  wStructSize As Integer
  X As Integer
  Y As Integer
  dX As Integer
  dY As Integer
  wMax As Integer
  rgChMember As String * 2
End Type
    'Declare Function WinHelp Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As Any) As Integer
    Declare Function WinHelpByInfo Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As HELPWININFO) As Integer
    Declare Function WinHelpByStr Lib "User" Alias "Winhelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$) As Integer
    Declare Function WinHelpByNum Lib "User" Alias "Winhelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData&) As Integer

    Dim m_hWndMainWindow As Integer ' hWnd to tell WINHELP the helpfile owner
Dim MainWindowInfo As HELPWININFO

Sub DoHelpMacro (ByVal Macrostring As String)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result As Variant

    Result = WinHelpByStr(m_hWndMainWindow, APP.HelpFile, HELP_COMMAND, ByVal (Macrostring))

End Sub

Sub QuitHelp ()
    Dim Result As Variant
    Result = WinHelp(m_hWndMainWindow, APP.HelpFile, HELP_QUIT, Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0))
End Sub

Sub SetApphelp (ByVal hWndMainWindow)
'=====================================================================
'To use these subroutines to access WINHELP, you need to add
'at least this one subroutine call to your code
'     o  In the Form_Load event of your main Form enter:
'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
'         (If you are not interested in keyword searching or context
'         sensitive help, this is the only call you need to make!)
'=====================================================================
    m_hWndMainWindow = hWndMainWindow
    If Right$(Trim$(APP.Path), 1) = "\" Then
        APP.HelpFile = APP.Path + "MMERHLP.HLP"
    Else
        APP.HelpFile = APP.Path + "\MMERHLP.HLP"
    End If
MainWindowInfo.wStructSize = 14
    MainWindowInfo.X = 256
    MainWindowInfo.Y = 256
    MainWindowInfo.dX = 512
    MainWindowInfo.dY = 512
    MainWindowInfo.rgChMember = Chr$(0) + Chr$(0)
End Sub

Sub ShowGlossary ()
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXT, CLng(64000))

End Sub

Sub ShowHelpContents ()
'=====================================================================
'  DISPLAY HELP STARTUP TOPIC IN RESPONSE TO A COMMAND BUTTON or MENU ...
'=====================================================================
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTENTS, CLng(0))

End Sub

Sub ShowHelpOnHelp ()
'=====================================================================
'  DISPLAY HELP for WINHELP.EXE  ...
'=====================================================================
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_HELPONHELP, CLng(0))

End Sub

Sub ShowHelpTopic (ByVal ContextID As Long)
'=====================================================================
'  FOR CONTEXT SENSITIVE HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic(<any Hlpxxx entry above>)
'=====================================================================
'  TO ADD FORM LEVEL CONTEXT SENSITIVE HELP...
'=====================================================================
'     o  For FORM level context sensetive help, you should set each
'        Me.HelpContext=<any Hlp_xxx entry above>
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXT, CLng(ContextID))

End Sub

Sub ShowHelpTopic2 (ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 2 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic2(<any Hlpxxx entry above>)
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile & ">HlpWnd02", HELP_CONTEXT, CLng(ContextID))

End Sub

Sub ShowHelpTopic3 (ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 3 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic3(<any Hlpxxx entry above>)
'
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile & ">HlpWnd03", HELP_CONTEXT, CLng(ContextID))

End Sub

Sub ShowPopupHelp (ByVal ContextID As Long)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result As Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXTPOPUP, CLng(ContextID))

End Sub

