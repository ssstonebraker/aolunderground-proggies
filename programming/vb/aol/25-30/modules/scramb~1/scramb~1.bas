Attribute VB_Name = "Module1"
' This Function is made 100% by me, PkY
' I don't ask for money, but for credit ONLY.
' That means if you decide to distribute a program
' that uses this function, please give me credit through
' the program (as "PkY").  And please leave this module
' unrevised when distributing as a module.  Contact me
' at PkY@Juno.com with any other questions/comments.

' Example of using this Function:

'     Result$ = Scrambler("This is easy")

' * An example return value = "hiTs si saye"

' This will return a string with the three words scrambled.
' You can easily revise whole phrase to be in lower/upper
' capitalization in the code (or even trim the phrase beforehand).
' It does not matter how many spaces, words, characters
' are trying to be scrambled.  It will scramble it the same way.
' All you have to do is add this module to your project and that's it!

Option Explicit
Function ScrambleIt(txt As String) As String    ' By: PkY


Dim Word$, Buff$
Dim Random%, I%, a%

Separate:

Do: DoEvents
    a% = InStr(txt$, " ")
    If a% = 0 Then
        Buff$ = txt$
        txt$ = ""
        Exit Do
    End If
    If a% = 1 Then
        ScrambleIt$ = ScrambleIt$ & " "
        txt$ = Right$(txt$, Len(txt$) - 1)
    End If
    If a% > 1 Then
        Buff$ = Left$(txt$, a% - 1)
        txt$ = Right$(txt$, Len(txt$) - a% + 1)
        Exit Do
    End If
Loop Until a% = 0

Word$ = ""

' Scrambles the word/section that was just separated.

For I% = 1 To Len(Buff$) - 1
    Random% = Int(Len(Buff$) * Rnd + 1)
    Word$ = Word$ & Mid$(Buff$, Random%, 1)
    Buff$ = Left$(Buff$, Random% - 1) & Right$(Buff$, Len(Buff$) - Random%)
Next I%

Word$ = Word$ & Buff$
ScrambleIt$ = ScrambleIt$ & Word$

If txt$ <> "" Then GoTo Separate


End Function
