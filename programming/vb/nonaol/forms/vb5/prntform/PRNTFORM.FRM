VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Prntform sample by Matt Hart - vbhelp@matthart.com"
   ClientHeight    =   2415
   ClientLeft      =   1125
   ClientTop       =   1920
   ClientWidth     =   7455
   Height          =   2820
   Left            =   1065
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   7455
   Top             =   1575
   Width           =   7575
   Begin VB.CommandButton cmdPrintScreen 
      Caption         =   "Print the Screen!"
      Height          =   495
      Left            =   4140
      TabIndex        =   2
      Top             =   1620
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrintForm 
      Caption         =   "Print the Form!"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1620
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   1035
      Left            =   420
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Prntform.frx":0000
      Top             =   180
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

' Prntform sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
'
' This sample utilizes a better method than "PrintForm" to print the contents of
' a Form. PrintForm sometimes excludes grid and other controls. This method simulates
' pressing the PrintScrn key, which copies the image of either the form or screen to
' the clipboard. Note that I call the .Clear method of the Clipboard first - that's because
' it might already have text or something on it.
'
' When printing, note that I scale the picture's height to proportionally fit the
' printer's resolution. The procedure would need to be adjusted if the Height of the
' Form/Screen was greater than Printer.ScaleHeight, or if the Height was greater than
' the Width and the Width was greater than Printer.ScaleWidth.

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const VK_SNAPSHOT As Byte = &H2C

Private Sub cmdPrintForm_Click()
    Dim lWidth As Long, lHeight As Long
    Clipboard.Clear
    Call keybd_event(VK_SNAPSHOT, 0, 0, 0)
    DoEvents
    Printer.Print
    If Width > Printer.ScaleWidth Then
        lWidth = Printer.ScaleWidth
        lHeight = (Printer.ScaleWidth / Width) * Height
    Else
        lWidth = Width
        lHeight = Height
    End If
    Printer.PaintPicture Clipboard.GetData, 0, 0, lWidth, lHeight
    Printer.EndDoc
    ' SavePicture Clipboard.GetData, Me.Name & ".BMP"
End Sub

Private Sub cmdPrintScreen_Click()
    Dim lHeight As Long
    Clipboard.Clear
    Call keybd_event(VK_SNAPSHOT, 1, 0, 0)
    DoEvents
    Printer.Print
    lHeight = (Printer.ScaleWidth / Screen.Width) * Screen.Height
    Printer.PaintPicture Clipboard.GetData, 0, 0, Printer.ScaleWidth, lHeight
    Printer.EndDoc
    ' SavePicture Clipboard.GetData, Me.Name & ".BMP"
End Sub
