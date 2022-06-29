Attribute VB_Name = "ContextIDs"
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
'List of Context IDs for <Server>
'=====================================================================
Global Const Hlp_XTreme_Server = 10    'Main Help Window
Global Const Hlp_Bust_ = 30    'Main Help Window
Global Const Hlp_Warranty_Disclaimer = 1070    'Main Help Window
Global Const Hlp_Instruction = 1100    'Main Help Window
Global Const Hlp_New_Features = 1110    'Main Help Window
Global Const Hlp_Mail_Comments = 1140    'Main Help Window
Global Const Hlp_M_Chat = 1150    'Main Help Window
Global Const Hlp_Commands = 1160    'Main Help Window
Global Const Hlp_Advertise = 1170    'Main Help Window
Global Const Hlp_Preferences = 1180    'Main Help Window
Global Const Hlp_File = 1190    'Main Help Window
Global Const Hlp_Banned_People = 1200    'Main Help Window
Global Const Hlp_Idle_Bot = 1210    'Main Help Window
Global Const Hlp_Kill = 1220    'Main Help Window
Global Const Hlp_Scroll = 1230    'Main Help Window
Global Const Hlp_Notify = 1240    'Main Help Window
Global Const Hlp_Mail_Type = 1250    'Main Help Window
Global Const Hlp_Write_Sending = 1260    'Main Help Window
Global Const Hlp_Kill_Wait = 1270    'Main Help Window
Global Const Hlp_Send_Status = 1280    'Main Help Window
Global Const Hlp_Commands_Pending = 1290    'Main Help Window
Global Const GLOS_Close_After_Busting =  1300
Global Const GLOS_Minimize_After_Busting =  1310
Global Const GLOS_Send_Tryxsx_To_Chat =  1320
Global Const GLOS_2xKill_Wait_After_Clicking_Send =  1330
Global Const GLOS_3xScroll_Status_After_List_Is_Sentx =  1340
Global Const GLOS_IM_ON =  1350
Global Const GLOS_IM_OFF =  1360
Global Const GLOS_Yes =  1370
Global Const GLOS_Cancel =  1380
Global Const GLOS_Save_To_Log_File =  1390
Global Const GLOS_Save_Settings =  1400
Global Const GLOS_Save_Settings_At_Exit =  1410
Global Const GLOS_Server_Test =  1420
Global Const GLOS_IM =  1430
Global Const GLOS_CHAT =  1440
Global Const GLOS_New_Mails =  1450
Global Const GLOS_Serve_Old_Mails =  1460
Global Const GLOS_Serve_Sent_Mails =  1470
Global Const GLOS_Serve_IncommingxSaved_Mails =  1480
Global Const GLOS_The_Server_Has_No_Mails_Limits =  1490
Global Const GLOS_Complete_Setup_File =  1500
Global Const GLOS_Help_File =  1510
Global Const GLOS_The_List_Is_Draw_Out_A_File =  1520
Global Const GLOS_Serve_New_Mails =  1530
'=====================================================================
'
'
'  Help engine section.

' Commands to pass WinHelp()
Global Const HELP_CONTEXT = &H1 '  Display topic in ulTopic
Global Const HELP_QUIT = &H2    '  Terminate help
Global Const HELP_FINDER = &HB  '  Display Contents tab
Global Const HELP_INDEX = &H3   '  Display index
Global Const HELP_HELPONHELP = &H4      '  Display help on using help
Global Const HELP_SETINDEX = &H5        '  Set the current Index for multi index help
Global Const HELP_KEY = &H101           '  Display topic for keyword in offabData
Global Const HELP_MULTIKEY = &H201
Global Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
Global Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
Global Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
Global Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
Global Const HELP_COMMAND = &H102    ' Execute Help macro
Global Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
Global Const HELP_SETWINPOS = &H203  ' Display and position Help window

    Type HELPWININFO
      wStructSize As Long
      X As Long
      Y As Long
      dX As Long
      dY As Long
      wMax As Long
      rgChMember As String * 2
    End Type
    Declare Function WinHelp Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
    Declare Function WinHelpByInfo Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As HELPWININFO) As Long
    Declare Function WinHelpByStr Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData$) As Long
    Declare Function WinHelpByNum Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData&) As Long
    Dim m_hWndMainWindow as Long ' hWnd to tell WINHELP the helpfile owner

Dim MainWindowInfo as HELPWININFO
Sub SetAppHelp (ByVal hWndMainWindow)
'=====================================================================
'To use these subroutines to access WINHELP, you need to add
'at least this one subroutine call to your code
'     o  In the Form_Load event of your main Form enter:
'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
'         (If you are not interested in keyword searching or context
'         sensitive help, this is the only call you need to make!)
'=====================================================================
    m_hWndMainWindow = hWndMainWindow
    If Right$(Trim$(App.Path),1) = "\" then
        App.HelpFile = App.Path + "Server.HLP"
    else
        App.HelpFile = App.Path + "\Server.HLP"
    end if
    MainWindowInfo.wStructSize = 26
    MainWindowInfo.X=256
    MainWindowInfo.Y=256
    MainWindowInfo.dX=512
    MainWindowInfo.dY=512
    MainWindowInfo.rgChMember=Chr$(0)+Chr$(0)
End Sub
Sub QuitHelp ()
    Dim Result as Variant
    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_QUIT, Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0))
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
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXT, Clng(ContextID))

End Sub
Sub ShowHelpTopic2 (ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 2 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic2(<any Hlpxxx entry above>)
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile &">HlpWnd02", HELP_CONTEXT, Clng(ContextID))

End Sub
Sub ShowHelpTopic3 (ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 3 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic3(<any Hlpxxx entry above>)
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile &">HlpWnd03", HELP_CONTEXT, Clng(ContextID))

End Sub
Sub ShowGlossary ()
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXT, Clng(64000))

End Sub
Sub ShowPopupHelp (ByVal ContextID As Long)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXTPOPUP, Clng(ContextID))

End Sub
Sub DoHelpMacro (ByVal MacroString As String)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result as Variant

    Result = WinHelpByStr(m_hWndMainWindow, APP.HelpFile, HELP_COMMAND, ByVal(Macrostring))

End Sub
Sub ShowHelpContents ()
'=====================================================================
'  DISPLAY STARTUP TOPIC IN RESPONSE TO A COMMAND BUTTON or MENU ...
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTENTS, Clng(0))

End Sub
Sub ShowContentsTab ()
'=====================================================================
'  DISPLAY Contents tab (*.CNT)
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_FINDER, Clng(0))

End Sub
Sub ShowHelpOnHelp ()
'=====================================================================
'  DISPLAY HELP for WINHELP.EXE  ...
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_HELPONHELP, Clng(0))

End Sub
