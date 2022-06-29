VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1725
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgDelete 
      Height          =   240
      Left            =   3600
      Picture         =   "Form1.frx":0000
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgExit 
      Height          =   240
      Left            =   2640
      Picture         =   "Form1.frx":0242
      Top             =   960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgCaution 
      Height          =   240
      Left            =   2640
      Picture         =   "Form1.frx":0944
      Top             =   600
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image imgYield 
      Height          =   240
      Left            =   1800
      Picture         =   "Form1.frx":0B86
      Top             =   960
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image imgStop 
      Height          =   240
      Left            =   1800
      Picture         =   "Form1.frx":0D88
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "&Data"
      Begin VB.Menu mnuDataDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuDataOptions 
         Caption         =   "&Options"
         Begin VB.Menu mnuDataOptionsStop 
            Caption         =   "Stop"
         End
         Begin VB.Menu mnuDataOptionsYield 
            Caption         =   "Yield"
         End
         Begin VB.Menu mnuDataOptionsCaution 
            Caption         =   "Caution"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API stuff for putting bitmaps in menus.
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Const MF_BITMAP = &H4&
Private Const MFT_BITMAP = MF_BITMAP
Private Const MIIM_TYPE = &H10

Private Sub Form_Load()
    ' Set the menu bitmaps.
    SetMenuBitmap Me, Array(0, 0), imgExit.Picture
    SetMenuBitmap Me, Array(1, 0), imgDelete.Picture
    SetMenuBitmap Me, Array(1, 1, 0), imgStop.Picture
    SetMenuBitmap Me, Array(1, 1, 1), imgYield.Picture
    SetMenuBitmap Me, Array(1, 1, 2), imgCaution.Picture
End Sub
' Put a bitmap in a menu item.
Public Sub SetMenuBitmap(ByVal frm As Form, ByVal item_numbers As Variant, ByVal pic As Picture)
Dim menu_handle As Long
Dim i As Integer
Dim menu_info As MENUITEMINFO

    ' Get the menu handle.
    menu_handle = GetMenu(frm.hwnd)
    For i = LBound(item_numbers) To UBound(item_numbers) - 1
        menu_handle = GetSubMenu(menu_handle, item_numbers(i))
    Next i

    ' Initialize the menu information.
    With menu_info
        .cbSize = Len(menu_info)
        .fMask = MIIM_TYPE
        .fType = MFT_BITMAP
        .dwTypeData = pic
    End With

    ' Assign the picture.
    SetMenuItemInfo menu_handle, _
        item_numbers(UBound(item_numbers)), _
        True, menu_info
End Sub


Private Sub mnuDataDelete_Click()
    MsgBox "Delete"
End Sub

Private Sub mnuDataOptionsCaution_Click()
    MsgBox "Caution"
End Sub

Private Sub mnuDataOptionsStop_Click()
    MsgBox "Stop"
End Sub


Private Sub mnuDataOptionsYield_Click()
    MsgBox "Yield"
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub


