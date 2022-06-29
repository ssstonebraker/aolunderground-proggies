VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRounder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Rounded Border Example
'Written by Chris Neuner
'in Visual Basic 6.0

'Description:
' Rounds the edges of a form

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Dim ret As Long 'Handle to the new region

Public Sub RoundedBorder(ByVal Frm As Form)
    Dim PriorScaleMode As Long 'Holds the initial ScaleMode of the form
    
     'Store the prior scalemode of the form
    PriorScaleMode = Frm.ScaleMode
    
    'If the prior scalemode isn't in Pixels, it is set to Pixels
    If PriorScaleMode <> vbPixels Then Frm.ScaleMode = vbPixels
    
    'Creates a rounded rectangualar region
    ret = CreateRoundRectRgn((Frm.ScaleWidth / 60), (Frm.ScaleHeight / 60), Frm.ScaleWidth - 4, Frm.ScaleHeight, (Frm.ScaleWidth / 20), (Frm.ScaleHeight / 20))
    
    'Selects the region into the form's device context
    SelectClipRgn Frm.hdc, ret
    
    'Copy the current clipping region into the temporary region
    GetClipRgn Frm.hdc, ret
    
    'Set the new window region
    SetWindowRgn Frm.hWnd, ret, True
    
    'Restore the prior scalemode
    If PriorScaleMode <> vbPixels Then Frm.ScaleMode = PriorScaleMode
End Sub


Private Sub Class_Terminate()
    'Deletes the region
    Call DeleteObject(ret)
End Sub
