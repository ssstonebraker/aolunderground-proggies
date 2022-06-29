VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Menubmp sample by Matt Hart - vbhelp@matthart.com"
   ClientHeight    =   3405
   ClientLeft      =   1140
   ClientTop       =   1800
   ClientWidth     =   7335
   Height          =   4095
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   7335
   Top             =   1170
   Width           =   7455
   Begin VB.TextBox Text1 
      Height          =   1875
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "menubmp.frx":0000
      Top             =   180
      Width           =   7215
   End
   Begin VB.Image imCopy 
      Height          =   195
      Left            =   2760
      Picture         =   "menubmp.frx":02B3
      Top             =   2640
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imPrintSetup 
      Height          =   225
      Left            =   2400
      Picture         =   "menubmp.frx":07C5
      Top             =   2640
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imPrint 
      Height          =   210
      Left            =   1920
      Picture         =   "menubmp.frx":0CBB
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imSave 
      Height          =   210
      Left            =   1440
      Picture         =   "menubmp.frx":11DD
      Top             =   2640
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imOpen 
      Height          =   195
      Left            =   960
      Picture         =   "menubmp.frx":16FF
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuPrintSetup 
         Caption         =   "Print &Setup"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuExtra 
         Caption         =   "&Extra Sub Menu"
         Begin VB.Menu mnuCopy 
            Caption         =   "&Copy"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Option Explicit

' MenuBmp sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' This sample shows you how to add bitmaps to your
' menu items.  First, you must retrieve the VB menu handle
' with the GetMenu API call.  Then, you set the Unchecked and
' Checked bitmaps (I don't differentiate between the two
' with this example - usually you would put a checkmark in
' the bitmap and use that for Checked).
'
' Note that the Picture property of an Image control (or
' a Picture control) is the Bitmap Handle, this the Image
' controls can be used like a bitmap resource.
'
' I also have one sub menu so that you can see how getting
' submenu handles and item positions works.

Private Declare Function GetMenu Lib "user32" _
   (ByVal hwnd As Long) As Long

Private Declare Function GetSubMenu Lib "user32" _
   (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function SetMenuItemBitmaps Lib "user32" _
   (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
    ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Const MF_BYPOSITION = &H400&

Private Sub Form_Load()
    Dim mHandle As Long, lRet As Long, sHandle As Long, sHandle2 As Long
    mHandle = GetMenu(hwnd)
    sHandle = GetSubMenu(mHandle, 0)
    lRet = SetMenuItemBitmaps(sHandle, 0, MF_BYPOSITION, imOpen.Picture, imOpen.Picture)
    lRet = SetMenuItemBitmaps(sHandle, 1, MF_BYPOSITION, imSave.Picture, imSave.Picture)
    lRet = SetMenuItemBitmaps(sHandle, 3, MF_BYPOSITION, imPrint.Picture, imPrint.Picture)
    lRet = SetMenuItemBitmaps(sHandle, 4, MF_BYPOSITION, imPrintSetup.Picture, imPrintSetup.Picture)
    sHandle = GetSubMenu(mHandle, 1)
    sHandle2 = GetSubMenu(sHandle, 0)
    lRet = SetMenuItemBitmaps(sHandle2, 0, MF_BYPOSITION, imCopy.Picture, imCopy.Picture)
End Sub
