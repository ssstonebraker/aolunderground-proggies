Attribute VB_Name = "modMacroFont"
Option Explicit

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Sub TextToPictureBox(txtSource As TextBox, picDestination As PictureBox)
    'this sub will take the text in our textbox and draw it to our
    'picturebox. first you will see that i reset the backcolor of
    'the picturebox. i do this so that our picturebox is "erased"
    'before we draw to it instead of drawing on top of our old text.
    'after using the textout api, notice that i will refresh the
    'picturebox. this is necessary since i have the picturebox's
    'autoredraw property set to true.
    picDestination.BackColor = vbWhite
    If txtSource.Text = "" Then
        picDestination.Refresh
    Else
        Call TextOut(picDestination.hdc, 0&, 0&, txtSource.Text, Len(txtSource.Text))
        picDestination.Refresh
    End If
End Sub

Public Function Convert(picSource As PictureBox) As String
    'this is the function which makes this all work. because this
    'function is more complicated than the rest, i will step through
    'it instead of explaining it all hear.
    Dim lngDoWidth As Long, lngDoHeight As Long
    Dim lngTop As Long, lngBottom As Long
    Dim strChar As String, strLine As String
    Dim strMacro As String, lngFix As Long, blnFix As Boolean
    Dim strFinal As String, strTmp As String
    
    For lngDoHeight& = 1 To picSource.ScaleHeight Step 2
    'we are starting our first for/next loop here. notice that this
    'loop is according to the height of the picturebox in pixels. we
    'are stepping through every other line here (hence the step 2).
    'we do this since when converting to our ascii art we must consider
    'two lines at a time as we will see later.
        strLine$ = ""
        For lngDoWidth& = 1 To picSource.ScaleWidth Step 1
        'here we are starting our loop which is according to the
        'picturebox's width. this loop will continue through until
        'we reach the end of the picturebox. at that point, we will
        'go back to our height loop. in other words, we are looping
        'from left to right, two lines of pixels at a time, over and
        'over again until we reach the extent of the picturebox's
        'surface.
            lngTop& = GetPixel(picSource.hdc, lngDoWidth&, lngDoHeight&)
            'first we retreive the long color value of our pixel.
            'i did try to use the point property of the picturebox,
            'but it proved to be slower than the getpixel api.
            lngBottom& = GetPixel(picSource.hdc, lngDoWidth&, lngDoHeight& + 1&)
            'again we're getting the long color value of a pixel,
            'except this time we're getting the pixel below the last.
            'we do this because we are looping through two lines at
            'a time (remember the step 2). we're going two lines at
            'a time since to create a smooth ascii image, we must
            'account for these two lines. you should be able to figure
            'out why below.
            If lngTop& = vbWhite And lngBottom& = vbWhite Then
            'here we check to see if both pixels are white. if so,
            'we know it is safe to use a space.
                strChar$ = " "
            End If
            If lngTop& <> vbWhite And lngBottom& <> vbWhite Then
            'if both pixels are not white, we will fill our "space"
            'with our ascii.
                strChar$ = ";"
            End If
            If lngTop& = vbWhite And lngBottom& <> vbWhite Then
                'if the top pixel is white, and the bottom is not,
                'then we will use a character which gives use the
                'appearance of a space on the top line and a fill
                'on the bottom line.
                strChar$ = ","
            End If
            If lngTop& <> vbWhite And lngBottom& = vbWhite Then
                'this is the opposite of the last if/then we just had.
                'here we are reacting to the top pixel not being white
                'and the bottom pixel being white.
                strChar$ = "´"
            End If
            If lngTop& = -1 And lngBottom& = -1 Then
                'this is just here to account for an odd number of
                'pixels. if there is no pixel there, we get a -1 return
                'from getpixel. we use this since the return will not
                'be white and we don't want to end up with ascii characters
                'when they're not wanted.
                strChar$ = " "
            End If
            strLine$ = strLine$ & strChar$
            'here we add our character to our current line.
        Next
        'in the following lines, if we have characters (not just a
        'long line of spaces) we will trim the spaces off the right
        'end of the string before adding them to our macro string.
        If Trim(strLine$) <> "" Then
            strLine$ = RTrim(strLine$)
        End If
        If strMacro$ = "" Then
            strMacro$ = strLine$
        Else
            strMacro$ = strMacro$ & vbCrLf & strLine$
        End If
    Next
    'the following code is not necessary, but i felt it was important
    'to do since we would end up with an awful lot of spaces which
    'were not needed. these two loops simply trim off the leading and
    'trailing lines which are filled with spaces only.
    blnFix = True
    For lngFix& = 1 To LineCount(strMacro$)
        strLine$ = LineFromString(strMacro$, lngFix&)
        strTmp$ = Replace(strLine$, " ", "")
        strTmp$ = Replace(strTmp$, vbCrLf, "")
        If strTmp$ <> "" Then
            blnFix = False
        End If
        If blnFix = False Then
            If strFinal$ = "" Then
                strFinal$ = strLine$
            Else
                strFinal$ = strFinal$ & vbCrLf & strLine$
            End If
        End If
    Next
    blnFix = True
    strMacro$ = ""
    For lngFix& = LineCount(strFinal$) To 1 Step -1
        strLine$ = LineFromString(strFinal$, lngFix&)
        strTmp$ = Replace(strLine$, " ", "")
        strTmp$ = Replace(strTmp$, vbCrLf, "")
        If strTmp$ <> "" Then
            blnFix = False
        End If
        If blnFix = False Then
            If strMacro$ = "" Then
                strMacro$ = strLine$
            Else
                strMacro$ = strLine$ & vbCrLf & strMacro$
            End If
        End If
    Next
    Convert$ = strMacro$
End Function

'the following functions should look familiar if you have dos32.bas.
'their purpose is to retreive the linecount from a string as well as
'extract a specific line from a string.

Public Function LineCount(strlngCount As String) As Long
    Dim lngPos As Long, lngCount As Long
    If Len(strlngCount$) < 1 Then
        LineCount& = 0&
        Exit Function
    End If
    lngPos& = InStr(strlngCount$, Chr(13))
    If lngPos& <> 0& Then
        LineCount& = 1
        Do
            lngPos& = InStr(lngPos + 1, strlngCount$, Chr(13))
            If lngPos& <> 0& Then
                LineCount& = LineCount& + 1
            End If
        Loop Until lngPos& = 0&
    End If
    LineCount& = LineCount& + 1
End Function

Public Function LineFromString(strSearch As String, lngLine As Long) As String
    Dim strCurLine As String, lngCount As Long
    Dim lngPos As Long, lngPosB As Long, lngDo As Long
    lngCount& = LineCount(strSearch$)
    If lngLine& > lngCount& Then
        Exit Function
    End If
    If lngLine& = 1 And lngCount& = 1 Then
        LineFromString$ = strSearch$
        Exit Function
    End If
    If lngLine& = 1 Then
        strCurLine$ = Left(strSearch$, InStr(strSearch$, Chr(13)) - 1)
        strCurLine$ = Replace(strCurLine$, Chr(13), "")
        strCurLine$ = Replace(strCurLine$, Chr(10), "")
        LineFromString$ = strCurLine$
        Exit Function
    Else
        lngPos& = InStr(strSearch$, Chr(13))
        For lngDo& = 1 To lngLine& - 1
            lngPosB& = lngPos&
            lngPos& = InStr(lngPos& + 1, strSearch$, Chr(13))
        Next lngDo
        If lngPos = 0 Then
            lngPos = Len(strSearch$)
        End If
        strCurLine$ = Mid(strSearch$, lngPosB&, lngPos& - lngPosB& + 1)
        strCurLine$ = Replace(strCurLine$, Chr(13), "")
        strCurLine$ = Replace(strCurLine$, Chr(10), "")
        LineFromString$ = strCurLine$
    End If
End Function
