VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mp3info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'written by smaq ;D
Private mvarFilePath As String 'local copy
Public Function Artist() As String
On Error Resume Next
Dim strFileExtract As String
Dim strExtract As String
Dim lngFileLength As Long
Dim lngPos As Long
    Open mvarFilePath$ For Binary As #1
        lngFileLength& = LOF(1)
        lngPos& = lngFileLength& - 128&
        strFileExtract$ = Space(128&)
        Get #1, lngPos& + 1&, strFileExtract$
    Close #1
    If Left(strFileExtract$, 3) <> "TAG" Then
        Artist$ = "no tag"
        Exit Function
    End If
    strFileExtract$ = Right(strFileExtract$, Len(strFileExtract$) - 3&)
    strFileExtract$ = Left(strFileExtract$, 60&)
    strFileExtract$ = Right(strFileExtract$, 30&)
    strFileExtract$ = Trim$(strFileExtract$)
    If strFileExtract$ = "" Then
        strFileExtract$ = "n/a"
    End If
    Artist$ = strFileExtract$
End Function

Public Function SongLength(media As MediaPlayer) As String
Dim lngSeconds As Long
Dim lngMinutes As Long
Dim strSeconds As String
Dim strMinutes As String
Dim lngPos As Long
    strSeconds$ = media.Duration
    If InStr(strSeconds$, ".") <> 0& Then
        strSeconds$ = Right(strSeconds$, Len(strSeconds$) - InStr(strSeconds$, "*"))
    End If
    lngSeconds& = strSeconds$
    Do Until lngSeconds& <= 59
        DoEvents
        lngSeconds& = lngSeconds& - 60
        lngMinutes& = lngMinutes& + 1
    Loop
    strSeconds$ = Format$(lngSeconds&, "00")
    strMinutes$ = Format$(lngMinutes&, "00")
    SongLength$ = strMinutes$ & ":" & strSeconds$
End Function

Public Property Let FilePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an
'assignment.  Syntax: X.FilePath = 5
    mvarFilePath = vData
End Property


Public Property Get FilePath() As String
'used when retrieving value of a property, on the right side of an
'assignment.  Syntax: Debug.Print X.FilePath
    FilePath = mvarFilePath
End Property
Public Function Comments() As String
On Error Resume Next
Dim strFileExtract As String
Dim strExtract As String
Dim lngFileLength As Long
Dim lngPos As Long
    Open mvarFilePath$ For Binary As #1
        lngFileLength& = LOF(1)
        lngPos& = lngFileLength& - 128&
        strFileExtract$ = Space(128&)
        Get #1, lngPos& + 1&, strFileExtract$
    Close #1
    If Left(strFileExtract$, 3) <> "TAG" Then
        Comments$ = "no tag"
        Exit Function
    End If
    strFileExtract$ = Right(strFileExtract$, Len(strFileExtract$) - 3&)
    strFileExtract$ = Right(strFileExtract$, 31&)
    strFileExtract$ = Left(strFileExtract$, 30&)
    strFileExtract$ = Trim$(strFileExtract$)
    If strFileExtract$ = "" Then
        strFileExtract$ = "n/a"
    End If
Comments$ = strFileExtract$
End Function
Public Function Year() As String
On Error Resume Next
Dim strFileExtract As String
Dim strExtract As String
Dim lngFileLength As Long
Dim lngPos As Long
    Open mvarFilePath$ For Binary As #1
        lngFileLength& = LOF(1)
        lngPos& = lngFileLength& - 128&
        strFileExtract$ = Space(128&)
        Get #1, lngPos& + 1&, strFileExtract$
    Close #1
    If Left(strFileExtract$, 3) <> "TAG" Then
        Year$ = "no tag"
        Exit Function
    End If
    strFileExtract$ = Right(strFileExtract$, Len(strFileExtract$) - 3&)
    strFileExtract$ = Left(strFileExtract$, 94&)
    strFileExtract$ = Right(strFileExtract$, 4&)
    strFileExtract$ = Trim$(strFileExtract$)
    If strFileExtract$ = "" Then
        strFileExtract$ = "n/a"
    End If
    Year$ = strFileExtract$
End Function
Public Function AlbumName() As String
On Error Resume Next
Dim strFileExtract As String
Dim strExtract As String
Dim lngFileLength As Long
Dim lngPos As Long
    Open mvarFilePath$ For Binary As #1
        lngFileLength& = LOF(1)
        lngPos& = lngFileLength& - 128&
        strFileExtract$ = Space(128&)
        Get #1, lngPos& + 1&, strFileExtract$
    Close #1
    If Left(strFileExtract$, 3) <> "TAG" Then
        AlbumName$ = "no tag"
        Exit Function
    End If
    strFileExtract$ = Right(strFileExtract$, Len(strFileExtract$) - 3&)
    strFileExtract$ = Left(strFileExtract$, 90&)
    strFileExtract$ = Right(strFileExtract$, 30&)
    strFileExtract$ = Trim$(strFileExtract$)
    If strFileExtract$ = "" Then
        strFileExtract$ = "n/a"
    End If
    AlbumName$ = strFileExtract$
End Function
Public Function SongTitle() As String
On Error Resume Next
Dim strFileExtract As String
Dim strExtract As String
Dim lngFileLength As Long
Dim lngPos As Long
    Open mvarFilePath$ For Binary As #1
        lngFileLength& = LOF(1)
        lngPos& = lngFileLength& - 128&
        strFileExtract$ = Space(128&)
        Get #1, lngPos& + 1&, strFileExtract$
    Close #1
    If Left(strFileExtract$, 3) <> "TAG" Then
        SongTitle$ = "no tag"
        Exit Function
    End If
    strFileExtract$ = Right(strFileExtract$, Len(strFileExtract$) - 3&)
    strFileExtract$ = Left(strFileExtract$, 30&)
    strFileExtract$ = Trim$(strFileExtract$)
    If strFileExtract$ = "" Then
        strFileExtract$ = "n/a"
    End If
    SongTitle = strFileExtract$
End Function

