Attribute VB_Name = "converter_mod"
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long

Public curicon
Public iconcount
Public lastlen
Public watchedit
Public mshop_frmleft
Public mshop_frmtop
Public scan_mode As Integer
Public old_x As Long
Public old_y As Long
Public new_x As Long
Public new_y As Long
Public scan_active As Integer
Public scan_active2 As Integer
Public filename As String
Public scan_letter As String
Public scan_font As String
Public scan_resx As Long
Public scan_timeout As Double
Public scanning As Integer
Public scan_offset
Public scan_bold As Integer
Public scan_underline As Integer
Public scan_italic As Integer
Public info_times As Integer
Public scan_mix
Public orig_offsetH
Public orig_offsetW
Public macro_name As String
Public Type convertdata
imagedrive As String
imagedir As String
imagefile As String
cscan_letter As String
cscan_font As String
cscan_timeout As Double
cscan_mix As Integer
cscan_bold As Integer
cscan_italics As Integer
cscan_underline As Integer
cscan_resx As Integer
cscan_offset As Double
End Type
Public currentsettings As convertdata


Public Function cmix(color1, color2) As String
Dim mixed As Long
Red1 = val("&h" & Mid(color1, 1, 2))
Green1 = val("&h" & Mid(color1, 3, 2))
Blue1 = val("&h" & Mid(color1, 5, 2))
Red2 = val("&h" & Mid(color2, 1, 2))
Green2 = val("&h" & Mid(color2, 3, 2))
Blue2 = val("&h" & Mid(color2, 5, 2))
red = (Red1 + Red2) / 2
green = (Green1 + Green2) / 2
blue = (Blue1 + Blue2) / 2
mixed = RGB(red, green, blue)
cmix = VbtoAol(mixed)
End Function
Public Function cmix2(color1, color2) As Long
Dim mixed As Long
Red1 = val("&h" & Mid(color1, 1, 2))
Green1 = val("&h" & Mid(color1, 3, 2))
Blue1 = val("&h" & Mid(color1, 5, 2))
Red2 = val("&h" & Mid(color2, 1, 2))
Green2 = val("&h" & Mid(color2, 3, 2))
Blue2 = val("&h" & Mid(color2, 5, 2))
red = (Red1 + Red2) / 2
green = (Green1 + Green2) / 2
blue = (Blue1 + Blue2) / 2
mixed = RGB(red, green, blue)
cmix2 = mixed
End Function
