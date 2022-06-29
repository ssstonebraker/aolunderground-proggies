Option Explicit

'----------------------------------------------------------------------------------
'   Name    :   MsgBeep.Bas
'   Author  :   Peter Wright
'   Date    :   21 September 1994
'
'   Notice  :   This code is freely distributable
'           :   Peter Wright, Psynet Ltd
'           :   peter@gendev.demon.co.uk
'----------------------------------------------------------------------------------

' The Windows MessageBeep API call plays the sound associated with the various msgbox
' icons. The sounds are set up using Control Panel.
Declare Sub MessageBeep Lib "User" (ByVal BeepType As Integer)

Function MsgBoxF (sMessage As String, nType As Integer, sTitle As String) As Integer

'--------------------------------------------------------------------------------------
'   Funcname:   msgboxF
'   Author  :   Peter Wright
'   Date    :   21 September 1993
'
'   Params  :   sMessage - the standard message you might pass to a MsgBox routine
'           :   nType    - the type code, again much the same as you would pass to MsgBox
'           :   sTitle   - the text to appear in the title bar of the messagebox
'
'   Return  :   Integer - same value returned by VB msgbox function.
'
'   Notes   :   This code emulates the standard MsgBox function, but plays the appropriate
'           :   windows sound depending on the type of message box being displayed.
'
'--------------------------------------------------------------------------------------
'                           C H A N G E    H I S T O R Y
'   [Date]      [Description]                                                   [Who]
'
'   20/6/94     Comments added to the code for Beginners Guide To VB            PJW
'
'--------------------------------------------------------------------------------------


    ' Declare a variable to hold the icon code, for the icon displayed in the box
    Dim nIcon As Integer


    ' Calculate which icon is being used based on the nType parameter passed into the code
    nIcon = nType And 112

    ' Call the MessageBeep subproc to play the required sound
    Call MessageBeep(nIcon)

    ' Run up the MsgBox as normal, returning the value returned by MsgBox function
    MsgBoxF = MsgBox(sMessage, nType, sTitle)

End Function




Sub MsgBoxS (sMessage As String, nType As Integer, sTitle As String)

'--------------------------------------------------------------------------------------
'   Subname:   msgboxS
'   Author  :   Peter Wright
'   Date    :   21 September 1993
'
'   Params  :   sMessage - the standard message you might pass to a MsgBox routine
'           :   nType    - the type code, again much the same as you would pass to MsgBox
'           :   sTitle   - the text to appear in the title bar of the messagebox
'
'   Notes   :   This code emulates the standard MsgBox subprocedure, but plays the appropriate
'           :   windows sound depending on the type of message box being displayed.
'
'--------------------------------------------------------------------------------------
'                           C H A N G E    H I S T O R Y
'   [Date]      [Description]                                                   [Who]
'
'   20/6/94     Comments added to the code for Beginners Guide To VB            PJW
'
'--------------------------------------------------------------------------------------


    ' Declare a variable to hold the value of the icon used in the message box
    Dim nIcon As Integer

    ' Calculate which icon was based, from the nType parameter
    nIcon = nType And 112

    ' First we call the MessageBeep subproc to play the required sound
    Call MessageBeep(nIcon)

    ' Now we can run up the MsgBox as normal
    MsgBox sMessage, nType, sTitle

End Sub

