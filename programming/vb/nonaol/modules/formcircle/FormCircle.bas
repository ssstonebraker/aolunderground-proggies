Attribute VB_Name = "FormCircle"
'This bas was PUT TOGETHER by VeGa....i did not
'create or in any way figure this code out...
'so i am not taking credit for it, the only reason
'i created this bas was to make it easier to use the
'circle form code...so don't go sayin I didnt make
'the code...BECAUSE I KNOW THAT!!!

'For help: X_VeGa_X@yahoo.com

'Please read the instructions on how to use the code
'by going to the code and reading the info









Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Public Function CircleForm(Frm As Form)
'To use this...Put "CircleForm Me" in Form: Load
SetWindowRgn Frm.hWnd, _
  CreateEllipticRgn(0, 0, 200, 300), True
'To make the circle WIDER(left to right)..change the
'300 on the previous line to a larger number
'To make the circle LONGER..change the 200 to a
'larger number...to make it smaller, make the 200
'or 300 a smaller number

'if the full circle isnt showing up...make the form
'bigger wider and longer..the form looks best when
'the borderstyle is set to none
End Function





