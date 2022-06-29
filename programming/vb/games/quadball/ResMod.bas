Attribute VB_Name = "ResMod"
' I Got This Sub From Planet-Source-Code
' It Was Written By Mike Miller.
' I added the ability to return a true or false
' so the programmer knows if the user's computer
' supports the display.
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Public Const CCDEVICENAME = 312
Public Const CCFORMNAME = 32
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
                

Public Function ChangeRes(iWidth As Single, iHeight As Single) As Boolean

Dim DevM As DEVMODE
Dim a As Boolean
Dim b&
Dim i&

i = 0
 Do
  a = EnumDisplaySettings(0&, i&, DevM)
  i = i + 1
 Loop Until (a = False)
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    DevM.dmPelsWidth = iWidth
    DevM.dmPelsHeight = iHeight
    b = ChangeDisplaySettings(DevM, 0)
    If b = 0 Then
     ChangeRes = True
    Else
     ChangeRes = False
    End If
End Function



