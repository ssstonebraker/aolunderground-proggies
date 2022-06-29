Option Explicit
'----------------------------------------------------
'   Name    : VB3D.BAS
'   Author  : Peter Wright
'   Date    : 5 June 1994
'
'   Notes   : This module allows your VB code to use the 3 main functions
'           : contained in CTL3D.DLL, Microsofts Windows 3D effects library
'           : necessary to give the built in dialogs a 3D look and feel identical to
'           : that used in Access 2, Excel 5, Word 6 and Powerpoint.
'
'   Funcs   : Start3D   - Turns the 3d effects on
'           : End3D     - Turns the 3d effects off
'
'   Notice  : CTL3DV2.DLL is a Microsoft DLL.
'           : (c) 1992-1994 Microsoft, All rights reserved.
'------------------------------------------------------

Declare Function GetModuleHandle Lib "Kernel" (ByVal lpModuleName As String) As Integer
Declare Function Ctl3dRegister Lib "CTL3DV2.DLL" (ByVal hInstance As Integer) As Integer
Declare Function Ctl3dAutoSubClass Lib "CTL3dV2.DLL" (ByVal hInstance As Integer) As Integer
Declare Function Ctl3dUnRegister Lib "Ctl3d.DLL" (ByVal hInstance As Integer) As Integer

Sub End3d ()

'--------------------------------------------------------------------------------------
'   Subname :   End3d
'   Author  :   Peter Wright
'   Date    :   5 June 1994
'
'   Notes   :   This code unregisters the application with CTL3DV2.DLL, allowing the
'           :   application to exit gracefully, ie without crashing.
'
'--------------------------------------------------------------------------------------
'                           C H A N G E    H I S T O R Y
'   [Date]      [Description]                                                   [Who]
'
'   20/6/94     Comments added to the code for Beginners Guide To VB            PJW
'
'-------------------------------------------------------------------------------------

    Dim hInstance As Integer
    Dim nReturn As Integer

    ' Get the handle to this applications instance again, as in the Start3D code
    hInstance = GetModuleHandle(App.Exename)

    ' Unregister the application with CTL3DV2.DLL
    nReturn = Ctl3dUnRegister(hInstance)

End Sub

Sub Start3d ()

'--------------------------------------------------------------------------------------
'   Subname :   Start3d
'   Author  :   Peter Wright
'   Date    :   5 June 1994
'
'   Notes   :   This code registers the application with the CTL3DV2.DLL, and then starts
'           :   of the subclassing procedure which turns all the dialogs into 3d dialogs.
'
'--------------------------------------------------------------------------------------
'                           C H A N G E    H I S T O R Y
'   [Date]      [Description]                                                   [Who]
'
'   20/6/94     Comments added to the code for Beginners Guide To VB            PJW
'
'-------------------------------------------------------------------------------------


    ' Declare a variable to a hold a "handle" to this programs instance. This is used
    ' CTL3DV2.DLL to register the application
    Dim hInstance As Integer

    ' Define a generic integer variable to hold unused return values from thge CTL3DV2.DLL
    ' functions
    Dim nReturn As Integer

    ' Get a handle to the application's instance
    hInstance = GetModuleHandle(App.Exename)

    ' Register this application with the VBCtl3D.DLL
    nReturn = Ctl3dRegister(hInstance)

    ' Start the graphics subclassing off  (all standard windows dialogs used within this
    ' application will now have a 3D appearance.
    nReturn = Ctl3dAutoSubClass(hInstance)

End Sub

