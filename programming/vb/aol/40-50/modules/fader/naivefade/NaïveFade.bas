Attribute VB_Name = "NaïveFade"
'|!¯¯¯¯\   |¯¯¯¯| '/¯¯¯¯/\¯¯¯¯\‚ |¯¯¯¯| |¯¯¯¯'|    '|¯¯¯¯'| '/¯¯¯¯/\¯¯¯¯\‚
'|:·.·.·:|\'\ '|:·.·.·:| |:·.·.·:|_|:·.·.·:'| |¯¯¯¯| |:·.·.·:'|    '|:·.·.·:'| |:·.·.·:|:/____/|
'|:·.·.·:|:'\'\|:·.·.·:| |:·.·.·:|¯|:·.·.·:'| |:·.·.·:| |\:·.·.·:\   /:·.·.·:/| |:·.·.·:|'|/¯¯¯¯/|‚
'|____|\::\____!| |____|:‚|____'| |____| |:'\____\/____/':| |____|/____/':|'
'|¯`·v·´| '\:'|'¯`v.´'| |¯`·v·´|¯|¯`·v·´'| |¯`·v·´| '\:'|'¯`·v´'||'¯`·v´|':/' |'¯`·v´·||'¯`·v´|':/‚
'|L,__'|   \|L,__,| |L,__'|  |'L,__‚| |L,__‚|   '\|L,__'||L,__|/   |L,__,||L,__'|/'‚
' (made with ribbits jungle font in macro machine 3 by builder)

' set your editor font to arial to see the macro

' Mail me at
'naivexistence@iname.com
'-or-
'GutterTrash@beer.com
'
'
'My name is Naïve, and I wrote this bas. Thats good for starters.
'Please, if you include this code or an alteration of this code, just
'put my name in you bas or whatever, its only fair. Since this bas
'is the only one that fades this way, we'll all know if you copied the
'code because the actual line equation is complex, I got it from my
'fav. VBgraphics proggin book, with the obvious name
' "visual basic graphics programming" by rod stevens. Get it,
'and have more originality than the guy who named that book,
' dont steal my code.
'
' Check this code out, it uses api to draw the lines which takes a little
' than 1/7th of the time that frm.line does
'
'The reason there are no other types of fader subs is that theres
'monkefade for that, no one needs another standard fade bas with the
'same subs.
'
'



    Declare Function MoveToEx Lib "gdi32" ( _
        ByVal hdc As Long, ByVal x As Long, _
        ByVal y As Long, lpPoint As Any) _
        As Long
    
    Declare Function LineTo Lib "gdi32" ( _
        ByVal hdc As Long, ByVal x As Long, _
        ByVal y As Long) As Long



Sub WaveFade(frm As Form)
'call in form_paint event

'makes the form blue in kind of a wave pattern
'you can easily change the colors
'this may take a while
frm.BackColor = vbBlack
Const Amp = 3
Const PI = 3.14159
Const Per = 4 * PI
Dim Red As Integer

Dim i As Single
Dim j As Single
Dim hgt As Single
Dim wID As Single
Dim TheColor As Long
Dim ClrStr As String
Dim OldColor As Long

OldColor = frm.ForeColor
    
    Red = 255
    frm.ScaleMode = 3   ' Pixel.
    
    frm.Cls     ' Clear the form

H = frm.hdc
    For i = 0 To frm.ScaleHeight Step 1
        s = MoveToEx(H, 0, i, ByVal 0&)
        If Red >= 0 Then
                TheColor = RGB(0, 0, Red)
        Else
            TheColor = vbBlack
        End If
        frm.ForeColor = TheColor
        For j = 0 To frm.ScaleWidth Step 4
            s = LineTo(H, j + 2, i + Amp * Sin(j / Per))
            s = LineTo(H, j + 2, i + Amp * Sin(j / Per))
            s = LineTo(H, j + 4, i + Amp * Sin(j / Per))
            s = LineTo(H, j + 4, i + Amp * Sin(j / Per))
        Next j
        Red = Red - 1
    Next i
frm.ForeColor = OldColor

End Sub

Sub CrazyColors(frm As Form)
'call in form_paint event

'kind of a patriotic feel :-)
frm.BackColor = vbBlack
Const Amp = 3
Const PI = 3.14159
Const Per = 4 * PI
Dim Red As Integer

Dim i As Single
Dim j As Single
Dim hgt As Single
Dim wID As Single
Dim TheColor As Long
Dim ClrStr As String
Dim OldColor As Long

OldColor = frm.ForeColor
    
    Red = 255
    frm.ScaleMode = 3   ' Pixel.
    
    frm.Cls     ' Clear the form

H = frm.hdc
    For i = 0 To frm.ScaleHeight Step 1
    
            p = i

p = Right(p, 1)
If p = "0" Then TheColor = vbPurple
If p = "1" Then TheColor = DarkBlue
If p = "2" Then TheColor = vbBlue
If p = "3" Then TheColor = LightBlue
If p = "4" Then TheColor = vbWhite
If p = "5" Then TheColor = LightRed
If p = "6" Then TheColor = vbRed
If p = "7" Then TheColor = vbBlack
If p = "8" Then TheColor = vbBlack
If p = "9" Then TheColor = vbBlack


        s = MoveToEx(H, 0, i, ByVal 0&)
        
        frm.ForeColor = TheColor
        For j = 0 To frm.ScaleWidth Step 4
            s = LineTo(H, j + 2, i + Amp * Sin(j / Per))
            s = LineTo(H, j + 2, i + Amp * Sin(j / Per))
            s = LineTo(H, j + 4, i + Amp * Sin(j / Per))
            s = LineTo(H, j + 4, i + Amp * Sin(j / Per))
        Next j
        Red = Red - 1
    Next i
frm.ForeColor = OldColor

End Sub


