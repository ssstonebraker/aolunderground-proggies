Attribute VB_Name = "FormFlash"
'Hey I thought this was neat.
'Flashes the title bar like aim does.
'Handy little tool I found. Feel free to use
'No I did not write this. I found the code
'In the book that came with vb6, but I put
'it all together into a bas form.
'****Important!****
'Look at the propertys and timer on the form
'They are nesscesary parts of code.
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal binvert As Long) As Long
Sub FlashSlow()
Me.Rate = 1
Me.Flash = True
End Sub
Sub FlashFast()
frmFlash.Rate = 3
frmFlash.Flash = True
End Sub

