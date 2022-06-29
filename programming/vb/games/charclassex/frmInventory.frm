VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3090
      Width           =   5445
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2685
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4736
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Who"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Attribute"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Info"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "What attribute means:"
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   2760
      Width           =   2145
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'  Set ItmX = ListView1.ListItems.Add(, , Text1.Text)
' ItmX.SubItems(1) = Text2.Text

Private Sub Form_Load()

End Sub



Private Sub ListView1_Click()
  Dim itemType As Integer
  Dim itemSelect As Integer
  Dim currentValue As Integer
  Dim currentString As String
  Dim strTarget As String
  itemSelect = ListView1.SelectedItem.Index - 1
  Text1.Text = ""
  
  Select Case itemSelect
    Case 0
      Text1.Text = "Empty"
    Case 1 To 11
      currentValue = itemList(itemSelect).Get_HPadd
      If (currentValue) Then
        Text1.Text = "Adds " & CStr(currentValue) & " Hit points" & strTarget & vbCrLf
      End If
      currentValue = itemList(itemSelect).Get_MPadd
      If (currentValue) Then
        Text1.Text = Text1.Text & "Adds " & currentValue & " Magic points" & strTarget & vbCrLf
      End If
      currentValue = itemList(itemSelect).Get_MPPercent
      If (currentValue) Then
        Text1.Text = Text1.Text & "Adds " & currentValue & " percent of total Magic points to MP" & strTarget & vbCrLf
      End If
      currentValue = itemList(itemSelect).Get_HPPercent
      If (currentValue) Then
        Text1.Text = Text1.Text & "Adds " & currentValue & " percent of total Hit points to HP" & strTarget & vbCrLf
      End If
      currentString = itemList(itemSelect).Get_Cure
      If Not currentString = "0" Then
         Text1.Text = Text1.Text & "Cures Status: " & GetStatus(currentString) & vbCrLf
      End If
    Case 12 To 51
      itemSelect = itemSelect - 12
      Text1.Text = "Attack value:   " & weaponList(itemSelect).Get_ATTACK & vbCrLf
      Text1.Text = Text1.Text & "Attack Percent: " & weaponList(itemSelect).Get_ATTPER & vbCrLf
      
      Text1.Text = Text1.Text & "Element Attack: " & weaponList(itemSelect).Get_ELEMENT & vbCrLf
      Text1.Text = Text1.Text & "Enemy Attack:   " & weaponList(itemSelect).Get_ENEMY & vbCrLf
    Case 52 To 80
      itemSelect = itemSelect - 52
      Text1.Text = "Armor defense:   " & armorList(itemSelect).Get_DEFENSE & vbCrLf
      Text1.Text = Text1.Text & "Armor defense %: " & armorList(itemSelect).Get_DEFPER & vbCrLf
      Text1.Text = Text1.Text & "Magic defense:   " & armorList(itemSelect).Get_MAGICDEF & vbCrLf
      Text1.Text = Text1.Text & "Magic defense %: " & armorList(itemSelect).Get_MAGPER & vbCrLf
      Text1.Text = Text1.Text & "Strong against element type: " & GetElement(armorList(itemSelect).Get_ELEMENT) & vbCrLf
      Text1.Text = Text1.Text & "Stronge against enemy type:   " & GetEnemy(armorList(itemSelect).Get_ENEMY)
    
    Case Else
      Text1.Text = "Error" 'Not really, but there is nothing in my data that
                           'is greater than #80
  End Select
End Sub
