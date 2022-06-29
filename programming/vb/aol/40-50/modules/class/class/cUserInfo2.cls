VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------
' By Ðigital-Flame(DigitalFlame@email.com)
' Created on: 11/19/98
'
' This is a small Class Module example to show
'     you what is on the inside and get
'        an ideah on how to do it.
'
' Ps: Please don't still my shit! write your
'    own code, or by a book or two like I did.
'-------------------------------------------
Option Explicit
Private Declare Function CompName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function UserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Function ComputerName() As String
Attribute ComputerName.VB_UserMemId = 0
Dim Bufstr As String
Bufstr = Space$(50)
If CompName(Bufstr, 50) > 0 Then
  ComputerName = Bufstr
  ComputerName = RTrim(ComputerName)
  ComputerName = StripTerm(ComputerName)
Else
  ComputerName = ""
End If
End Function

Private Function StripTerm(ByVal stString As String) As String
Dim ZPos As Integer
ZPos = InStr(stString, Chr$(0))
If ZPos > 0 Then
    StripTerm = Left$(stString, ZPos - 1)
    Else
        StripTerm = stString
        End If
End Function

Public Function UserNamez() As String
Dim Bufstr As String
Bufstr = Space$(50)
If UserName(Bufstr, 50) > 0 Then
    UserNamez = Bufstr
    UserNamez = RTrim(UserNamez)
    UserNamez = StripTerm(UserNamez)
Else
    UserNamez = "*Could not get username*"
End If
End Function
