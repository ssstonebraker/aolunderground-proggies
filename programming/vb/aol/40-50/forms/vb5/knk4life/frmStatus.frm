VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   2865
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   2865
   Begin VB5Chat2.Chat Chat1 
      Left            =   120
      Top             =   120
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
'Dim SN As String, Said As String
'SN$ = LCase(ReplaceString(GetUser$, " ", ""))
'Said$ = LCase(ReplaceString(What_Said$, " ", ""))
'If InStr(Said$, SN$) And InStr(Said$, "list") Then
'    Text1.SelStart = Len(Text1.Text)
'    Text1.SelText = "" & Screen_Name & " Said he has sent you the list" & vbCrLf & "---------------------------------------------------"
'End If


If What_Said Like "*" & GetUser & "*" & "list" & "*" Then
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = "" & Screen_Name & " Said he has sent you the list" & vbCrLf & "---------------------------------------------------"
End If

If What_Said Like "*" & LCase(GetUser) & "*" & "list" & "*" Then
    Text1.SelStart = Len(Text1.Text) '
    Text1.SelText = "" & Screen_Name & " Said he has sent you the list" & vbCrLf & "---------------------------------------------------"
End If

User$ = GetUser
User$ = ReplaceString(User$, " ", "")
If What_Said Like "*" & User$ & "*" & "list" & "*" Then
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = "" & Screen_Name & " Said he has sent you the list" & vbCrLf & "---------------------------------------------------"
End If

If What_Said Like "*" & LCase(User$) & "*" & "list" & "*" Then
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = "" & Screen_Name & " Said he has sent you the list" & vbCrLf & "---------------------------------------------------"
End If

If What_Said Like "*" & GetUser & "*" & "List" & "*" Then
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = "" & Screen_Name & " Said he has sent you the list" & vbCrLf & "---------------------------------------------------"
End If

If What_Said Like "*" & LCase(GetUser) & "*" & "List" & "*" Then
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = "" & Screen_Name & " Said he has sent you the list" & vbCrLf & "---------------------------------------------------"
End If

User$ = GetUser
User$ = ReplaceString(User$, " ", "")
If What_Said Like "*" & User$ & "*" & "List" & "*" Then
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = "" & Screen_Name & " Said he has sent you the list" & vbCrLf & "---------------------------------------------------"
End If

If What_Said Like "*" & LCase(User$) & "*" & "List" & "*" Then
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = "" & Screen_Name & " Said he has sent you the list" & vbCrLf & "---------------------------------------------------"
End If

End Sub

Private Sub Form_Load()
FormOnTop Me
Chat1.ScanOn
End Sub
