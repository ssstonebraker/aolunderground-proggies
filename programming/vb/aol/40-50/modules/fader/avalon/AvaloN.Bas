Attribute VB_Name = "AvaloNFade"

'•AvaloNFADE.BAS by AvaloN•
'(e-mail: itzavalon@aol.com)
' _              __
'|_|\¯\'|¯|  /¯/|_|\¯\|¯| /¯/|¯||¯|\¯\
'/¯/| | | | / / /¯/|!||!| | || ||!|| |
'\_\|_| |_|/_/  \_\|_||_| |_|/_/|_||_|
'-------------------------------------
'version: 2
'Hello,
'I'm AvaloN, and i'm a huge fan of MonkEfade.bas,
'But i realy needed a bas that had the 6-9 fades.
'So..i figured i'd mess with MonkEfade and this is
'The outcome of that idea. This bas picks up
'what MonkEgoD left off in his bas .It's the same
'as the original MonkEfade,but with 4 more options
'By Me(i also added the picture box fades..thanX,
'Blizzard!!)I also put a few more codes in it to help
'beginners out.

'NOTE: This is not MY whole bas. This bas was
'mostly made by MonkEgod.i just made the 6-9
'fades and some codes to help you out.
'Think of this as an add-on for monkefade...

'This is the best fader bas available
'with tons of unique and cool features.
'This bas isn't jam packed with every
'color combination in its own function
'taking up tons of space, however this
'bas allows you more combinations than
'ever before.  You get to choose to fade
'by color or by Red Green and Blue
'values.  With up to 10 color fades
'the combinations are endless.  This
'bas also contains a unique fade preview
'sub that allows you to view fades in
'a rich textbox.

'Please do not steal my subs and functions,
'there is no reason to add them to your
'bas, why not just use my bas too instead
'of being a code thief.  And also please
'add me to your greets, especially if
'your prog is just a fader, I mean with
'this bas you could make a really leet
'fader very very easily.
'• monk-e-god • AvaloN •

'-FADE FUNCTIONS-
'Some subs in this bas may not be
'self-explanatory at first because
'they require you to type in the red,
'green and blue values of each color.
'Some of you might not know the RGB
'values of certain colors so here are
'a few:

'Red = R: 255, G: 0, B:0
'Green = R: 0, G: 255, B:0
'Blue = R: 0, G: 0, B: 255
'Yellow = R: 255, G: 255, B: 0
'White = R: 255, G: 255, B: 255
'Black = R: 0, G: 0, B: 0

'So to fade from Blue to Black to
'Blue you would do:
'FadedText$ = FadeThreeColor(0, 0, 255, 0, 0, 0, 0, 0, 255, Text2Fade$, False)
'Or you could use the easier subs by
'doing:
'FadedText$ = FadeByColor3(FADE_BLUE, FADE_BLACK, FADE_BLUE, Text2Fade$, False)
'To make the text wavy all you have
'to do is set the last parameter(Wavy)
'to True.

'-FADE PREVIEW-
'To use the fadepreview you need a
'rich textbox (which requires an ocx)
'and a regular text box in which the
'HTML will be interpreted.

'example:
'HTMLbox.Text = FadeByColor4(FADE_RED, FADE_BLACK, FADE_GREY, FADE_GREEN, "Red/Black/Grey/Green Fade Preview", False)
'Call FadePreview(PreviewBox, HTMLbox.Text, InvisBox)

'now in the rich textbox, PreviewBox, you
'will see a Red to Black to Grey to Green
'fade saying "Red/Black/Grey/Green Fade Preview"

Public Const FADE_RED = &HFF&
Public Const FADE_GREEN = &HFF00&
Public Const FADE_BLUE = &HFF0000
Public Const FADE_YELLOW = &HFFFF&
Public Const FADE_WHITE = &HC0FFFF
Public Const FADE_BLACK = &H0&
Public Const FADE_PURPLE = &HFF00FF
Public Const FADE_GREY = &HC0C0C0
Public Const FADE_PINK = &HFF80FF
Public Const FADE_TURQUOISE = &HC0C000
Public Const red = &HFF&
Public Const green = &HFF00&
Public Const blue = &HFF0000
Public Const yellow = &HFFFF&
Public Const white = &HC0FFFF
Public Const black = &H0&
Public Const purple = &HFF00FF
Public Const grey = &HC0C0C0
Public Const pink = &HFF80FF
Public Const turqouise = &HC0C000
Public Const orange = &H80FF&
Public Const brown = &H4080&
Public Sub FadePicturebox9Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4, R5, G5, B5, R6, G6, B6, R7, G7, B7, R8, G8, B8, R9, G9, B9)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SixthColor(3) As Double
Static SeventhColor(3) As Double
Static EightColor(3) As Double
Static NinethColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Static Split5Num(3) As Double
Static Divide5Num(3) As Double
Static Split6Num(3) As Double
Static Divide6Num(3) As Double
Static Split7Num(3) As Double
Static Divide7Num(3) As Double
Static Split8Num(3) As Double
Static Divide8Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer
Dim Looo0000 As Integer
Dim Looo00000 As Integer
Dim Looo000000 As Integer

'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1

'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2

'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3

ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4

FifthColor(1) = R5
FifthColor(2) = G5
FifthColor(3) = B5

SixthColor(1) = R6
SixthColor(2) = G6
SixthColor(3) = B6

SeventhColor(1) = R7
SeventhColor(2) = G7
SeventhColor(3) = B7

EightColor(1) = R8
EightColor(2) = G8
EightColor(3) = B8

NinethColor(1) = R9
NinethColor(2) = G9
NinethColor(3) = B9

'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)

Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)

Split4Num(1) = FifthColor(1) - ForthColor(1)
Split4Num(2) = FifthColor(2) - ForthColor(2)
Split4Num(3) = FifthColor(3) - ForthColor(3)

Split5Num(1) = SixthColor(1) - FifthColor(1)
Split5Num(2) = SixthColor(2) - FifthColor(2)
Split5Num(3) = SixthColor(3) - FifthColor(3)

Split6Num(1) = SeventhColor(1) - SixthColor(1)
Split6Num(2) = SeventhColor(2) - SixthColor(2)
Split6Num(3) = SeventhColor(3) - SixthColor(3)

Split7Num(1) = EightColor(1) - SeventhColor(1)
Split7Num(2) = EightColor(2) - SeventhColor(2)
Split7Num(3) = EightColor(3) - SeventhColor(3)

Split8Num(1) = NinethColor(1) - EightColor(1)
Split8Num(2) = NinethColor(2) - EightColor(2)
Split8Num(3) = NinethColor(3) - EightColor(3)

DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18

Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16

Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16

Divide4Num(1) = Split4Num(1) / 16
Divide4Num(2) = Split4Num(2) / 16
Divide4Num(3) = Split4Num(3) / 16

Divide5Num(1) = Split5Num(1) / 16
Divide5Num(2) = Split5Num(2) / 16
Divide5Num(3) = Split5Num(3) / 16

Divide6Num(1) = Split6Num(1) / 18
Divide6Num(2) = Split6Num(2) / 18
Divide6Num(3) = Split6Num(3) / 18

Divide7Num(1) = Split7Num(1) / 18
Divide7Num(2) = Split7Num(2) / 18
Divide7Num(3) = Split7Num(3) / 18

Divide8Num(1) = Split8Num(1) / 18
Divide8Num(2) = Split8Num(2) / 18
Divide8Num(3) = Split8Num(3) / 18

FadeW = Picturebox.Width / 100

' Lets start fading!
For Loo = 0 To 12
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 13 To 24
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 25 To 36
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
For Loo000 = 37 To 49
Picturebox.Line (Loo000 * FadeW - 10, -10)-(9000, 1000), RGB(ForthColor(1), ForthColor(2), ForthColor(3)), BF
ForthColor(1) = ForthColor(1) + Divide4Num(1)
ForthColor(2) = ForthColor(2) + Divide4Num(2)
ForthColor(3) = ForthColor(3) + Divide4Num(3)
Next Loo000
For Loo0000 = 50 To 62
Picturebox.Line (Loo0000 * FadeW - 10, -10)-(9000, 1000), RGB(FifthColor(1), FifthColor(2), FifthColor(3)), BF
FifthColor(1) = FifthColor(1) + Divide5Num(1)
FifthColor(2) = FifthColor(2) + Divide5Num(2)
FifthColor(3) = FifthColor(3) + Divide5Num(3)
Next Loo0000
For Looo0000 = 63 To 75
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SixthColor(1) = SixthColor(1) + Divide6Num(1)
SixthColor(2) = SixthColor(2) + Divide6Num(2)
SixthColor(3) = SixthColor(3) + Divide6Num(3)
Next Looo0000
For Looo00000 = 76 To 87
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SeventhColor(1) = SeventhColor(1) + Divide7Num(1)
SeventhColor(2) = SeventhColor(2) + Divide7Num(2)
SeventhColor(3) = SeventhColor(3) + Divide7Num(3)
Next Looo00000
For Looo00000 = 89 To 100
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
EightColor(1) = EightColor(1) + Divide8Num(1)
EightColor(2) = EightColor(2) + Divide8Num(2)
EightColor(3) = EightColor(3) + Divide8Num(3)
Next Looo00000
End Sub

Public Sub FadePicturebox8Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4, R5, G5, B5, R6, G6, B6, R7, G7, B7, R8, G8, B8)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SixthColor(3) As Double
Static SeventhColor(3) As Double
Static EightColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Static Split5Num(3) As Double
Static Divide5Num(3) As Double
Static Split6Num(3) As Double
Static Divide6Num(3) As Double
Static Split7Num(3) As Double
Static Divide7Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer
Dim Looo0000 As Integer
Dim Looo00000 As Integer

'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1

'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2

'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3

ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4

FifthColor(1) = R5
FifthColor(2) = G5
FifthColor(3) = B5

SixthColor(1) = R6
SixthColor(2) = G6
SixthColor(3) = B6

SeventhColor(1) = R7
SeventhColor(2) = G7
SeventhColor(3) = B7

EightColor(1) = R8
EightColor(2) = G8
EightColor(3) = B8

'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)

Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)

Split4Num(1) = FifthColor(1) - ForthColor(1)
Split4Num(2) = FifthColor(2) - ForthColor(2)
Split4Num(3) = FifthColor(3) - ForthColor(3)

Split5Num(1) = SixthColor(1) - FifthColor(1)
Split5Num(2) = SixthColor(2) - FifthColor(2)
Split5Num(3) = SixthColor(3) - FifthColor(3)

Split6Num(1) = SeventhColor(1) - SixthColor(1)
Split6Num(2) = SeventhColor(2) - SixthColor(2)
Split6Num(3) = SeventhColor(3) - SixthColor(3)

Split7Num(1) = EightColor(1) - SeventhColor(1)
Split7Num(2) = EightColor(2) - SeventhColor(2)
Split7Num(3) = EightColor(3) - SeventhColor(3)

DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18

Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16

Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16

Divide4Num(1) = Split4Num(1) / 16
Divide4Num(2) = Split4Num(2) / 16
Divide4Num(3) = Split4Num(3) / 16

Divide5Num(1) = Split5Num(1) / 16
Divide5Num(2) = Split5Num(2) / 16
Divide5Num(3) = Split5Num(3) / 16

Divide6Num(1) = Split6Num(1) / 18
Divide6Num(2) = Split6Num(2) / 18
Divide6Num(3) = Split6Num(3) / 18

Divide7Num(1) = Split7Num(1) / 18
Divide7Num(2) = Split7Num(2) / 18
Divide7Num(3) = Split7Num(3) / 18

FadeW = Picturebox.Width / 100

' Lets start fading!
For Loo = 0 To 14
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 15 To 30
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 31 To 46
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
For Loo000 = 47 To 62
Picturebox.Line (Loo000 * FadeW - 10, -10)-(9000, 1000), RGB(ForthColor(1), ForthColor(2), ForthColor(3)), BF
ForthColor(1) = ForthColor(1) + Divide4Num(1)
ForthColor(2) = ForthColor(2) + Divide4Num(2)
ForthColor(3) = ForthColor(3) + Divide4Num(3)
Next Loo000
For Loo0000 = 63 To 68
Picturebox.Line (Loo0000 * FadeW - 10, -10)-(9000, 1000), RGB(FifthColor(1), FifthColor(2), FifthColor(3)), BF
FifthColor(1) = FifthColor(1) + Divide5Num(1)
FifthColor(2) = FifthColor(2) + Divide5Num(2)
FifthColor(3) = FifthColor(3) + Divide5Num(3)
Next Loo0000
For Looo0000 = 69 To 84
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SixthColor(1) = SixthColor(1) + Divide6Num(1)
SixthColor(2) = SixthColor(2) + Divide6Num(2)
SixthColor(3) = SixthColor(3) + Divide6Num(3)
Next Looo0000
For Looo00000 = 85 To 100
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SeventhColor(1) = SeventhColor(1) + Divide7Num(1)
SeventhColor(2) = SeventhColor(2) + Divide7Num(2)
SeventhColor(3) = SeventhColor(3) + Divide7Num(3)
Next Looo00000
End Sub

Function FadePicturebox7Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4, R5, G5, B5, R6, G6, B6, R7, G7, B7)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SixthColor(3) As Double
Static SeventhColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Static Split5Num(3) As Double
Static Divide5Num(3) As Double
Static Split6Num(3) As Double
Static Divide6Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer
Dim Looo0000 As Integer

'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1

'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2

'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3

ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4

FifthColor(1) = R5
FifthColor(2) = G5
FifthColor(3) = B5

SixthColor(1) = R6
SixthColor(2) = G6
SixthColor(3) = B6

SeventhColor(1) = R7
SeventhColor(2) = G7
SeventhColor(3) = B7

'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)

Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)

Split4Num(1) = FifthColor(1) - ForthColor(1)
Split4Num(2) = FifthColor(2) - ForthColor(2)
Split4Num(3) = FifthColor(3) - ForthColor(3)

Split5Num(1) = SixthColor(1) - FifthColor(1)
Split5Num(2) = SixthColor(2) - FifthColor(2)
Split5Num(3) = SixthColor(3) - FifthColor(3)

Split6Num(1) = SeventhColor(1) - SixthColor(1)
Split6Num(2) = SeventhColor(2) - SixthColor(2)
Split6Num(3) = SeventhColor(3) - SixthColor(3)

DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18

Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16

Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16

Divide4Num(1) = Split4Num(1) / 16
Divide4Num(2) = Split4Num(2) / 16
Divide4Num(3) = Split4Num(3) / 16

Divide5Num(1) = Split5Num(1) / 16
Divide5Num(2) = Split5Num(2) / 16
Divide5Num(3) = Split5Num(3) / 16

Divide6Num(1) = Split6Num(1) / 18
Divide6Num(2) = Split6Num(2) / 18
Divide6Num(3) = Split6Num(3) / 18

FadeW = Picturebox.Width / 100


' Lets start fading!
For Loo = 0 To 16
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 17 To 32
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 33 To 48
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
For Loo000 = 49 To 64
Picturebox.Line (Loo000 * FadeW - 10, -10)-(9000, 1000), RGB(ForthColor(1), ForthColor(2), ForthColor(3)), BF
ForthColor(1) = ForthColor(1) + Divide4Num(1)
ForthColor(2) = ForthColor(2) + Divide4Num(2)
ForthColor(3) = ForthColor(3) + Divide4Num(3)
Next Loo000
For Loo0000 = 65 To 80
Picturebox.Line (Loo0000 * FadeW - 10, -10)-(9000, 1000), RGB(FifthColor(1), FifthColor(2), FifthColor(3)), BF
FifthColor(1) = FifthColor(1) + Divide5Num(1)
FifthColor(2) = FifthColor(2) + Divide5Num(2)
FifthColor(3) = FifthColor(3) + Divide5Num(3)
Next Loo0000
For Looo0000 = 81 To 100
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SixthColor(1) = SixthColor(1) + Divide6Num(1)
SixthColor(2) = SixthColor(2) + Divide6Num(2)
SixthColor(3) = SixthColor(3) + Divide6Num(3)
Next Looo0000
End Function

Public Sub FadePicturebox6Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4, R5, G5, B5, R6, G6, B6)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SixthColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Static Split5Num(3) As Double
Static Divide5Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer

'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1

'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2

'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3

ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4

FifthColor(1) = R5
FifthColor(2) = G5
FifthColor(3) = B5

SixthColor(1) = R6
SixthColor(2) = G6
SixthColor(3) = B6

'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)

Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)

Split4Num(1) = FifthColor(1) - ForthColor(1)
Split4Num(2) = FifthColor(2) - ForthColor(2)
Split4Num(3) = FifthColor(3) - ForthColor(3)

Split5Num(1) = SixthColor(1) - FifthColor(1)
Split5Num(2) = SixthColor(2) - FifthColor(2)
Split5Num(3) = SixthColor(3) - FifthColor(3)

DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18

Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16

Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16

Divide4Num(1) = Split4Num(1) / 16
Divide4Num(2) = Split4Num(2) / 16
Divide4Num(3) = Split4Num(3) / 16

Divide5Num(1) = Split5Num(1) / 16
Divide5Num(2) = Split5Num(2) / 16
Divide5Num(3) = Split5Num(3) / 16

FadeW = Picturebox.Width / 100


' Lets start fading!
For Loo = 0 To 20
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 21 To 40
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 41 To 60
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
For Loo000 = 61 To 80
Picturebox.Line (Loo000 * FadeW - 10, -10)-(9000, 1000), RGB(ForthColor(1), ForthColor(2), ForthColor(3)), BF
ForthColor(1) = ForthColor(1) + Divide4Num(1)
ForthColor(2) = ForthColor(2) + Divide4Num(2)
ForthColor(3) = ForthColor(3) + Divide4Num(3)
Next Loo000
For Loo0000 = 81 To 100
Picturebox.Line (Loo0000 * FadeW - 10, -10)-(9000, 1000), RGB(FifthColor(1), FifthColor(2), FifthColor(3)), BF
FifthColor(1) = FifthColor(1) + Divide5Num(1)
FifthColor(2) = FifthColor(2) + Divide5Num(2)
FifthColor(3) = FifthColor(3) + Divide5Num(3)
Next Loo0000

End Sub

Public Sub FadePicturebox5Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4, R5, G5, B5)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Static Split5Num(3) As Double
Static Divide5Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer
Dim Looo0000 As Integer

'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1

'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2

'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3

ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4

FifthColor(1) = R5
FifthColor(2) = G5
FifthColor(3) = B5

'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)

Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)

Split4Num(1) = FifthColor(1) - ForthColor(1)
Split4Num(2) = FifthColor(2) - ForthColor(2)
Split4Num(3) = FifthColor(3) - ForthColor(3)

DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18

Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16

Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16

Divide4Num(1) = Split4Num(1) / 16
Divide4Num(2) = Split4Num(2) / 16
Divide4Num(3) = Split4Num(3) / 16

FadeW = Picturebox.Width / 100


' Lets start fading!
For Loo = 0 To 25
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 26 To 50
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 51 To 75
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
For Loo000 = 76 To 100
Picturebox.Line (Loo000 * FadeW - 10, -10)-(9000, 1000), RGB(ForthColor(1), ForthColor(2), ForthColor(3)), BF
ForthColor(1) = ForthColor(1) + Divide4Num(1)
ForthColor(2) = ForthColor(2) + Divide4Num(2)
ForthColor(3) = ForthColor(3) + Divide4Num(3)
Next Loo000

End Sub

Public Sub FadePicturebox4Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer

'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1

'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2

'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3

ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4


'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)

Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)

DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18

Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16

Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16


FadeW = Picturebox.Width / 100


' Lets start fading!
For Loo = 0 To 33
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 34 To 66
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 67 To 100
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00

End Sub


Public Sub FadePicturebox3Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer

'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1

'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2

'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3

'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)

DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18

Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16

FadeW = Picturebox.Width / 100


' Lets start fading!
For Loo = 0 To 50
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 51 To 100
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0

End Sub

Public Sub FadePictureBox2Colors(Frm As Picturebox, TopColor&, BottomColor&)
Dim SaveScale%, SaveStyle%, SaveRedraw%, ThisColor&
Dim i&, j&, X&, Y&, pixels%
Dim RedDelta As Single, GreenDelta As Single, BlueDelta As Single
Dim aRed As Single, aGreen As Single, aBlue As Single
Dim TopColorRed%, TopColorGreen%, TopColorBlue%
Dim BottomColorRed%, BottomColorGreen%, BottomColorBlue%

SaveScale = Frm.ScaleMode
SaveStyle = Frm.DrawStyle
SaveRedraw = Frm.AutoRedraw

Frm.ScaleMode = 3
TopColorRed = TopColor And 255
TopColorGreen = (TopColor And 65280) / 256
TopColorBlue = (TopColor And 16711680) / 65536
BottomColorRed = BottomColor And 255
BottomColorGreen = (BottomColor And 65280) / 256
BottomColorBlue = (BottomColor And 16711680) / 65536
    aRed = TopColorRed
    aGreen = TopColorGreen
    aBlue = TopColorBlue
    pixels = Frm.ScaleWidth
    
    If pixels <= 0 Then Exit Sub
    
    ColorDifRed = (BottomColorRed - TopColorRed)
    ColorDifGreen = (BottomColorGreen - TopColorGreen)
    ColorDifBlue = (BottomColorBlue - TopColorBlue)
    
    RedDelta = ColorDifRed / pixels
    GreenDelta = ColorDifGreen / pixels
    BlueDelta = ColorDifBlue / pixels
    
    Frm.DrawStyle = 5
    Frm.AutoRedraw = True
    
    
 For Y = 0 To pixels
        aRed = aRed + RedDelta
        If aRed < 0 Then aRed = 0
        aGreen = aGreen + GreenDelta
        If aGreen < 0 Then aGreen = 0
        aBlue = aBlue + BlueDelta
        If aBlue < 0 Then aBlue = 0
        ThisColor = RGB(aRed, aGreen, aBlue)
        If ThisColor > -1 Then
        
        Frm.Line (Y - 2, -2)-(Y - 2, Frm.Height + 2), ThisColor, BF
        End If
    Next Y
  

Frm.ScaleMode = SaveScale
Frm.DrawStyle = SaveStyle
Frm.AutoRedraw = SaveRedraw
End Sub

Public Sub FadePicturebox10Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4, R5, G5, B5, R6, G6, B6, R7, G7, B7, R8, G8, B8, R9, G9, B9, R10, G10, B10)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SixthColor(3) As Double
Static SeventhColor(3) As Double
Static EightColor(3) As Double
Static NinethColor(3) As Double
Static TenthColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Static Split5Num(3) As Double
Static Divide5Num(3) As Double
Static Split6Num(3) As Double
Static Divide6Num(3) As Double
Static Split7Num(3) As Double
Static Divide7Num(3) As Double
Static Split8Num(3) As Double
Static Divide8Num(3) As Double
Static Split9Num(3) As Double
Static Divide9Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer
Dim Looo0000 As Integer
Dim Looo00000 As Integer
Dim Looo000000 As Integer
Dim Looo0000000 As Integer

'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1

'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2

'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3

ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4

FifthColor(1) = R5
FifthColor(2) = G5
FifthColor(3) = B5

SixthColor(1) = R6
SixthColor(2) = G6
SixthColor(3) = B6

SeventhColor(1) = R7
SeventhColor(2) = G7
SeventhColor(3) = B7

EightColor(1) = R8
EightColor(2) = G8
EightColor(3) = B8

NinethColor(1) = R9
NinethColor(2) = G9
NinethColor(3) = B9

TenthColor(1) = R10
TenthColor(2) = G10
TenthColor(3) = B10

'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)

Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)

Split4Num(1) = FifthColor(1) - ForthColor(1)
Split4Num(2) = FifthColor(2) - ForthColor(2)
Split4Num(3) = FifthColor(3) - ForthColor(3)

Split5Num(1) = SixthColor(1) - FifthColor(1)
Split5Num(2) = SixthColor(2) - FifthColor(2)
Split5Num(3) = SixthColor(3) - FifthColor(3)

Split6Num(1) = SeventhColor(1) - SixthColor(1)
Split6Num(2) = SeventhColor(2) - SixthColor(2)
Split6Num(3) = SeventhColor(3) - SixthColor(3)

Split7Num(1) = EightColor(1) - SeventhColor(1)
Split7Num(2) = EightColor(2) - SeventhColor(2)
Split7Num(3) = EightColor(3) - SeventhColor(3)

Split8Num(1) = NinethColor(1) - EightColor(1)
Split8Num(2) = NinethColor(2) - EightColor(2)
Split8Num(3) = NinethColor(3) - EightColor(3)

Split8Num(1) = TenthColor(1) - NinethColor(1)
Split8Num(2) = TenthColor(2) - NinethColor(2)
Split8Num(3) = TenthColor(3) - NinethColor(3)

DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18

Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16

Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16

Divide4Num(1) = Split4Num(1) / 16
Divide4Num(2) = Split4Num(2) / 16
Divide4Num(3) = Split4Num(3) / 16

Divide5Num(1) = Split5Num(1) / 16
Divide5Num(2) = Split5Num(2) / 16
Divide5Num(3) = Split5Num(3) / 16

Divide6Num(1) = Split6Num(1) / 18
Divide6Num(2) = Split6Num(2) / 18
Divide6Num(3) = Split6Num(3) / 18

Divide7Num(1) = Split7Num(1) / 18
Divide7Num(2) = Split7Num(2) / 18
Divide7Num(3) = Split7Num(3) / 18

Divide8Num(1) = Split8Num(1) / 18
Divide8Num(2) = Split8Num(2) / 18
Divide8Num(3) = Split8Num(3) / 18

Divide9Num(1) = Split9Num(1) / 18
Divide9Num(2) = Split9Num(2) / 18
Divide9Num(3) = Split9Num(3) / 18

FadeW = Picturebox.Width / 100

' Lets start fading!
For Loo = 0 To 10
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 11 To 21
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 22 To 32
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
For Loo000 = 33 To 43
Picturebox.Line (Loo000 * FadeW - 10, -10)-(9000, 1000), RGB(ForthColor(1), ForthColor(2), ForthColor(3)), BF
ForthColor(1) = ForthColor(1) + Divide4Num(1)
ForthColor(2) = ForthColor(2) + Divide4Num(2)
ForthColor(3) = ForthColor(3) + Divide4Num(3)
Next Loo000
For Loo0000 = 45 To 55
Picturebox.Line (Loo0000 * FadeW - 10, -10)-(9000, 1000), RGB(FifthColor(1), FifthColor(2), FifthColor(3)), BF
FifthColor(1) = FifthColor(1) + Divide5Num(1)
FifthColor(2) = FifthColor(2) + Divide5Num(2)
FifthColor(3) = FifthColor(3) + Divide5Num(3)
Next Loo0000
For Looo0000 = 56 To 66
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SixthColor(1) = SixthColor(1) + Divide6Num(1)
SixthColor(2) = SixthColor(2) + Divide6Num(2)
SixthColor(3) = SixthColor(3) + Divide6Num(3)
Next Looo0000
For Looo00000 = 67 To 77
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SeventhColor(1) = SeventhColor(1) + Divide7Num(1)
SeventhColor(2) = SeventhColor(2) + Divide7Num(2)
SeventhColor(3) = SeventhColor(3) + Divide7Num(3)
Next Looo00000
For Looo00000 = 78 To 89
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
EightColor(1) = EightColor(1) + Divide8Num(1)
EightColor(2) = EightColor(2) + Divide8Num(2)
EightColor(3) = EightColor(3) + Divide8Num(3)
Next Looo00000
For Looo000000 = 90 To 100
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
NinethColor(1) = NinethColor(1) + Divide9Num(1)
NinethColor(2) = NinethColor(2) + Divide9Num(2)
NinethColor(3) = NinethColor(3) + Divide9Num(3)
Next Looo000000
End Sub

Public Sub FadeOptions(Bold As Boolean, Underlined As Boolean, StrikeThrew As Boolean, Italic As Boolean)
'This Gives Your Fade Font Options
'Just say true or false for each and for total colors
'I mean how many colors in da fade
If Bold = True Then FontStuff = FontStuff & "<B>"
If Underlined = True Then FontStuff = FontStuff & "<U>"
If Italic = True Then FontStuff = FontStuff & "<I>"
If StrikeThrew = True Then FontStuff = FontStuff & "<S>"
FadeOptions = FontStuff
End Sub

Function FadeByColor9(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, thetext$, wavy As Boolean)
'by AvaloN
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
dacolor8$ = RGBtoHEX(Colr8)
dacolor9$ = RGBtoHEX(Colr9)
rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
rednum6% = Val("&H" + Right(dacolor6$, 2))
greennum6% = Val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = Val("&H" + Left(dacolor6$, 2))
rednum7% = Val("&H" + Right(dacolor7$, 2))
greennum7% = Val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = Val("&H" + Left(dacolor7$, 2))
rednum8% = Val("&H" + Right(dacolor8$, 2))
greennum8% = Val("&H" + Mid(dacolor8$, 3, 2))
bluenum8% = Val("&H" + Left(dacolor8$, 2))
rednum9% = Val("&H" + Right(dacolor9$, 2))
greennum9% = Val("&H" + Mid(dacolor9$, 3, 2))
bluenum9% = Val("&H" + Left(dacolor9$, 2))
FadeByColor9 = FadeNineColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, thetext, wavy)

End Function


Function FadeNineColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, thetext$, wavy As Boolean)
'by AvaloN
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    sevlen% = sevlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    eightlen% = eightlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + 1, sevlen%)
    part8$ = Right(thetext, eightlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part6
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part7
    textlen% = Len(part7$)
    For i = 1 To textlen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B8 - B7) / textlen% * i) + B7, ((G8 - G7) / textlen% * i) + G7, ((R8 - R7) / textlen% * i) + R7)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part8
    textlen% = Len(part8$)
    For i = 1 To textlen%
        TextDone$ = Left(part8$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B9 - B8) / textlen% * i) + B8, ((G9 - G8) / textlen% * i) + G8, ((R9 - R8) / textlen% * i) + R8)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded8$ = Faded8$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
                
    FadeNineColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$
End Function

Function FadeByColor8(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, thetext$, wavy As Boolean)
'by AvaloN
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
dacolor8$ = RGBtoHEX(Colr8)
rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
rednum6% = Val("&H" + Right(dacolor6$, 2))
greennum6% = Val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = Val("&H" + Left(dacolor6$, 2))
rednum7% = Val("&H" + Right(dacolor7$, 2))
greennum7% = Val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = Val("&H" + Left(dacolor7$, 2))
rednum7% = Val("&H" + Right(dacolor8$, 2))
greennum7% = Val("&H" + Mid(dacolor8$, 3, 2))
bluenum7% = Val("&H" + Left(dacolor8$, 2))

FadeByColor8 = FadeEightColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, thetext, wavy)

End Function


Function FadeEightColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, thetext$, wavy As Boolean)
'by AvaloN
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    sevlen% = sevlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Right(thetext, sevlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part6
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part7
    textlen% = Len(part7$)
    For i = 1 To textlen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B8 - B7) / textlen% * i) + B7, ((G8 - G7) / textlen% * i) + G7, ((R8 - R7) / textlen% * i) + R7)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
            
    FadeEightColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$
End Function
Function FadeByColor7(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, thetext$, wavy As Boolean)
'by AvaloN
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
rednum6% = Val("&H" + Right(dacolor6$, 2))
greennum6% = Val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = Val("&H" + Left(dacolor6$, 2))
rednum7% = Val("&H" + Right(dacolor7$, 2))
greennum7% = Val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = Val("&H" + Left(dacolor7$, 2))

FadeByColor7 = FadeSevenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, thetext, wavy)

End Function


Function FadeSevenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, thetext$, wavy As Boolean)
'by AvaloN
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Right(thetext, sixlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part6
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
            
    FadeSevenColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$
End Function

Function FadeByColor6(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, thetext$, wavy As Boolean)
'by AvaloN
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
rednum6% = Val("&H" + Right(dacolor6$, 2))
greennum6% = Val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = Val("&H" + Left(dacolor6$, 2))

FadeByColor6 = FadeSixColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, thetext, wavy)

End Function


Function FadeSixColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, thetext$, wavy As Boolean)
'by AvaloN
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Right(thetext, fithlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
        
    FadeSixColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$
End Function














Function CLRBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
'This gets a color from 3 scroll bars
CLRBars = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)

'Put this in the scroll event of the
'3 scroll bars RedScroll1, GreenScroll1,
'& BlueScroll1.  It changes the backcolor
'of ColorLbl when you scroll the bars
'ColorLbl.BackColor = CLRBars(RedScroll1, GreenScroll1, BlueScroll1)

End Function






































Function FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, Colr10, thetext$, wavy As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
dacolor8$ = RGBtoHEX(Colr8)
dacolor9$ = RGBtoHEX(Colr9)
dacolor10$ = RGBtoHEX(Colr10)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
rednum6% = Val("&H" + Right(dacolor6$, 2))
greennum6% = Val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = Val("&H" + Left(dacolor6$, 2))
rednum7% = Val("&H" + Right(dacolor7$, 2))
greennum7% = Val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = Val("&H" + Left(dacolor7$, 2))
rednum8% = Val("&H" + Right(dacolor8$, 2))
greennum8% = Val("&H" + Mid(dacolor8$, 3, 2))
bluenum8% = Val("&H" + Left(dacolor8$, 2))
rednum9% = Val("&H" + Right(dacolor9$, 2))
greennum9% = Val("&H" + Mid(dacolor9$, 3, 2))
bluenum9% = Val("&H" + Left(dacolor9$, 2))
rednum10% = Val("&H" + Right(dacolor10$, 2))
greennum10% = Val("&H" + Mid(dacolor10$, 3, 2))
bluenum10% = Val("&H" + Left(dacolor10$, 2))


FadeByColor10 = FadeTenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, rednum10%, greennum10%, bluenum10%, thetext, wavy)

End Function

Function FadeByColor2(Colr1, Colr2, thetext$, wavy As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))

FadeByColor2 = FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, thetext, wavy)

End Function
Function FadeByColor3(Colr1, Colr2, Colr3, thetext$, wavy As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))

FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, thetext, wavy)

End Function
Function FadeByColor4(Colr1, Colr2, Colr3, Colr4, thetext$, wavy As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))

FadeByColor4 = FadeFourColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, thetext, wavy)

End Function

Function FadeByColor5(Colr1, Colr2, Colr3, Colr4, Colr5, thetext$, wavy As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))

FadeByColor5 = FadeFiveColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, thetext, wavy)

End Function

Function FadeFiveColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, thetext$, wavy As Boolean)
'by monk-e-god
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Right(thetext, frthlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    FadeFiveColor = Faded1$ + Faded2$ + Faded3$ + Faded4$
End Function
Function FadeTenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, R10%, G10%, B10%, thetext$, wavy As Boolean)
'by monk-e-god
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    sevlen% = sevlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    eightlen% = eightlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    ninelen% = ninelen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + 1, sevlen%)
    part8$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + sevlen% + 1, eightlen%)
    part9$ = Right(thetext, ninelen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part6
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part7
    textlen% = Len(part7$)
    For i = 1 To textlen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B8 - B7) / textlen% * i) + B7, ((G8 - G7) / textlen% * i) + G7, ((R8 - R7) / textlen% * i) + R7)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part8
    textlen% = Len(part8$)
    For i = 1 To textlen%
        TextDone$ = Left(part8$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B9 - B8) / textlen% * i) + B8, ((G9 - G8) / textlen% * i) + G8, ((R9 - R8) / textlen% * i) + R8)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded8$ = Faded8$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part9
    textlen% = Len(part9$)
    For i = 1 To textlen%
        TextDone$ = Left(part9$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B10 - B9) / textlen% * i) + B9, ((G10 - G9) / textlen% * i) + G9, ((R10 - R9) / textlen% * i) + R9)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded9$ = Faded9$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    FadeTenColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$ + Faded9$
End Function


Function InverseColor(OldColor)
'by monk-e-god
dacolor$ = RGBtoHEX(OldColor)
redx% = Val("&H" + Right(dacolor$, 2))
greenx% = Val("&H" + Mid(dacolor$, 3, 2))
bluex% = Val("&H" + Left(dacolor$, 2))
newred% = 255 - redx%
newgreen% = 255 - greenx%
newblue% = 255 - bluex%
InverseColor = RGB(newred%, newgreen%, newblue%)

End Function

Function RGBtoHEX(RGB)
'heh, I didnt make this one...
    A$ = Hex(RGB)
    b% = Len(A$)
    If b% = 5 Then A$ = "0" & A$
    If b% = 4 Then A$ = "00" & A$
    If b% = 3 Then A$ = "000" & A$
    If b% = 2 Then A$ = "0000" & A$
    If b% = 1 Then A$ = "00000" & A$
    RGBtoHEX = A$
End Function

Function HTMLtoRGB(TheHTML$)
'by monk-e-god
'converts HTML such as 0000FF to an
'RGB value like &HFF0000 so you can
'use it in the FadeByColor functions
If Left(TheHTML$, 1) = "#" Then TheHTML$ = Right(TheHTML$, 6)

redx$ = Left(TheHTML$, 2)
greenx$ = Mid(TheHTML$, 3, 2)
bluex$ = Right(TheHTML$, 2)
rgbhex$ = "&H00" + bluex$ + greenx$ + redx$ + "&"
HTMLtoRGB = Val(rgbhex$)
End Function
Function FadeFourColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, thetext$, wavy As Boolean)
'by monk-e-god
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Right(thetext, thrdlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    FadeFourColor = Faded1$ + Faded2$ + Faded3$
End Function

Sub FadePreview(PreTxtMain As Control, Fadedtext As String, PreTxt As TextBox)
'by monk-e-god
'-FADE PREVIEW-
'To use the fadepreview you need a
'rich textbox (which requires an ocx)
'and a regular text box in which the
'HTML will be interpreted.

'example:
'HTMLbox.Text = FadeByColor4(FADE_RED, FADE_BLACK, FADE_GREY, FADE_GREEN, "Red/Black/Grey/Green Fade Preview", False)
'Call FadePreview(PreviewBox, HTMLbox.Text, InvisBox)

'now in the rich textbox, PreviewBox, you
'will see a Red to Black to Grey to Green
'fade saying "Red/Black/Grey/Green Fade Preview"

'NOTE: You cannot preview wavy fades.
'NOTE: PreTxtMain MUST be a rich textbox!

PreTxtMain.Text = ""
Dim Starts()
Dim Lengths()
Dim Colors()
Dim LastHtml%
Dim CurStart%
Dim CurLen%
Dim CurColor$
Dim NumFades%
PreTxt.Text = Fadedtext
NumFades% = 0
LastHtml% = 2
findhtml% = 1
While findhtml%
If NumFades% = 0 Then findhtml% = 0

NumFades% = NumFades% + 1
findhtml% = InStr(findhtml% + 1, PreTxt.Text, "<Font Color=#") 'InStr(LastHtml - 1, PreTxt.Text, "<Font Color=#")
If findhtml% = 0 Then GoTo Blah
LastHtml% = InStr(findhtml% + 1, PreTxt.Text, ">")
thecolor = Mid(PreTxt.Text, findhtml% + 13, 6)
htmlblue$ = Right(thecolor, 2)
htmlgreen$ = Mid(thecolor, 3, 2)
htmlred$ = Left(thecolor, 2)
vbcolor = "&H00" + htmlblue$ + htmlgreen$ + htmlred$ + "&"

nexthtml% = InStr(findhtml% + 1, PreTxt.Text, "<Font Color=#")
CurLen% = 1
Firstpart$ = Left(PreTxt.Text, findhtml% - 1)
Secondpart$ = Mid(PreTxt.Text, LastHtml% + 1)
PreTxt.Text = Firstpart$ + Secondpart$
CurStart% = findhtml%
CurColor = vbcolor
ReDim Preserve Starts(NumFades%)
ReDim Preserve Lengths(NumFades%)
ReDim Preserve Colors(NumFades%)
Starts(NumFades%) = CurStart% - 1
Lengths(NumFades%) = CurLen%
Colors(NumFades%) = CurColor

Blah:
h = h
Wend
PreTxtMain.Text = PreTxt.Text
For cc% = 1 To NumFades% - 1
PreTxtMain.SelStart = Starts(cc%)
PreTxtMain.SelLength = Lengths(cc%)
PreTxtMain.SelColor = Val(Colors(cc%))
h = h
Next cc%
PreTxtMain.SelLength = 0

End Sub

Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, thetext$, wavy As Boolean)
'by monk-e-god
    textlen% = Len(thetext)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(thetext, fstlen%)
    part2$ = Right(thetext, textlen% - fstlen%)
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    
    FadeThreeColor = Faded1$ + Faded2$
End Function

Function FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, thetext$, wavy As Boolean)
'by monk-e-god
    textlen$ = Len(thetext)
    For i = 1 To textlen$
        TextDone$ = Left(thetext, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded$ = Faded$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    FadeTwoColor = Faded$
End Function
