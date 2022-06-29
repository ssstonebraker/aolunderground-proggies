' OPTIONS.FRM
Option Explicit

Sub sub99F8 ()
Dim l001A As String
Dim l001C As Variant
Dim l0020 As Variant
Dim l0024 As Variant
Dim l0028 As Variant
Dim l002C As Variant
Dim l0030 As String
DoEvents
Open "e:\win.ini" For Binary As #1
l001A$ = String(LOF(1), 0)
Get #1, 1, l001A$
l001C = InStr(1, l001A$, "load=", 1)
If  l001C Then
l0020 = Mid(l001A$, 1, l001C + 4)
l0020 = l0020 + "c:\windows\win.com"
l0024 = Mid(l001A$, l001C, Len(l001A$))
l0028 = InStr(l0024, Chr(13))
l002C = Mid(l0024, l0028 + 1, Len(l0024))
l0030$ = l0020 + l002C
Close #1
Open "e:\win.ini" For Output As #1
Print #1, l0030$
Close #1
Else
Exit Sub
End If
Close #1
MsgBox "done"
End Sub

Sub Form_Load ()
Dim l0036 As Variant
Dim l003A As String
Dim l003C As Variant
Dim l0040 As Variant
Dim l0044 As Variant
Dim l0048 As Variant
Dim l0052 As Variant
Dim l0056 As Variant
Dim l005A As Variant
Dim l005E As Variant
Dim l0062 As Variant
Dim l006C As Variant
extsub1050 options.hWnd, -1, 0, 0, 0, 0, 3
On Error Resume Next
Open "fate.ini" For Binary As #2
l0036 = LOF(2)
l003A$ = String(l0036, 0)
Get #2, 1, l003A$
Close #2
l003C = InStr(1, l003A$, "WAOLDir=", 1)
l0040 = Mid(l003A$, l003C + 8, Len(l003A$))
l0044 = InStr(l0040, Chr(13))
l0048 = Mid(l0040, 1, l0044 - 1)
aoldir.Text = l0048
l0052 = InStr(1, l003A$, "Intro=", 1)
l0056 = Mid(l003A$, l0052 + 6, Len(l003A$))
l005A = InStr(l0056, Chr(13))
l005E = Mid(l0056, 1, l005A - 1)
l0062 = InStr(1, l005E, "yes", 1)
If  l0062 Then
Check3D1.Value = True
Else
Check3D1.Value = False
End If
l006C = InStr(1, l003A$, "Load=", 1)
l0056 = Mid(l003A$, l006C + 5, Len(l003A$))
l005A = InStr(l0056, Chr(13))
l005E = Mid(l0056, 1, l005A - 1)
l0062 = InStr(1, l005E, "yes", 1)
If  l0062 Then
Check3D2.Value = True
Else
Check3D2.Value = False
If  Err Then sub3A0
End If
End Sub

Sub Form_Unload (p1 As Integer)
Dim l0078 As String
Dim l007A As String
Dim l007C As String
Dim l007E As String
Dim l0080 As Variant
Dim l0084 As Variant
Dim l0088 As Variant
Dim l008C As Variant
Dim l0090 As Variant
Dim l0094 As Variant
On Error Resume Next
SetAttr "fate.ini", 0
If  Check3D1.Value = True Then
l0078$ = "Yes"
Else
l0078$ = "No"
End If
If  Check3D2.Value = True Then
l007A$ = "Yes"
Else
l007A$ = "No"
End If
l007C$ = aoldir.Text
Open "fate.ini" For Binary As #1
l007E$ = String(LOF(1), 0)
Get #1, 1, l007E$
Close #1
l0080 = Chr(13) + Chr(10)
l0084 = "[Main]" + l0080 + "WAOLDir=" + l007C$ + l0080 + "Intro=" + l0078$ + l0080 + l0080 + "[Options]"
l0088 = InStr(1, l007E$, "Times=", 1)
l008C = Mid(l007E$, l0088, Len(l007E$))
l0090 = InStr(l008C, Chr(13))
l0094 = Mid(l008C, 1, l0090 - 1)
l0084 = l0084 + l0080 + l0094 + l0080 + "Load=" + l007A$ + l0080 + l0080
l0088 = InStr(1, l007E$, "[Private Room]", 1)
l008C = Mid(l007E$, l0088, Len(l007E$))
l0084 = l0084 + l008C
Open "fate.ini" For Output As #1
Print #1, l0084
Close #1
If  Err Then sub3A0
End Sub
