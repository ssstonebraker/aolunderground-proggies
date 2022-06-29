VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "macro font converter"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7575
      Begin VB.TextBox txtMacro 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   7575
      Begin VB.CommandButton cmdConvert 
         Caption         =   "convert"
         Height          =   330
         Left            =   6480
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cmbFonts 
         Height          =   330
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   8
         Text            =   "arial"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.VScrollBar vsbFontSize 
         Height          =   255
         Left            =   7080
         Max             =   4
         Min             =   140
         TabIndex        =   6
         Top             =   1320
         Value           =   20
         Width           =   375
      End
      Begin VB.CheckBox chkItalic 
         BackColor       =   &H00C0C0C0&
         Caption         =   "italic"
         Height          =   210
         Left            =   4920
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox chkBold 
         BackColor       =   &H00C0C0C0&
         Caption         =   "bold"
         Height          =   210
         Left            =   4320
         TabIndex        =   4
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox chkUnderline 
         BackColor       =   &H00C0C0C0&
         Caption         =   "underline"
         Height          =   210
         Left            =   5520
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.PictureBox picMacro 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         ScaleHeight     =   121
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   273
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   4320
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblFontSize 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         Height          =   255
         Left            =   6600
         TabIndex        =   7
         Top             =   1290
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'5.1.99 - after several requests for this, i decided that even though
'i felt something like this wasn't worth while, i decided to go ahead
'and do it anyway. honestly something like this is a novelty, but like
'i said, you guys wanted this. i guess its kind of cool. well it was
'for about 5 or 10 min anyway. honestly, i thought it would've been
'a little harder. oh well. either way, enjoy. if you have any questions
'or comments, feel free to contact me.

'dos
'dos@hider.com
'www.hider.com/dos

Private lngFontSize As Long

Private Sub chkBold_Click()
    'in this sub we're checking to see if the bold checkbox is
    'checked or unchecked. we will change the fontbold properties
    'of our textbox and picturebox as well as redraw our text to
    'the picturebox.
    If chkBold.Value = vbChecked Then
        txtText.FontBold = True
        picMacro.FontBold = True
    Else
        txtText.FontBold = False
        picMacro.FontBold = False
    End If
    Call TextToPictureBox(txtText, picMacro)
End Sub

Private Sub chkItalic_Click()
    'this is the same as the bold checkbox. this time we're setting
    'the italic properties.
    If chkItalic.Value = vbChecked Then
        txtText.FontItalic = True
        picMacro.FontItalic = True
    Else
        txtText.FontItalic = False
        picMacro.FontItalic = False
    End If
    Call TextToPictureBox(txtText, picMacro)
End Sub

Private Sub chkUnderline_Click()
    'again, like the bold and italic checkboxes, we're setting another
    'font property, this time the undline property.
    If chkUnderline.Value = vbChecked Then
        txtText.FontUnderline = True
        picMacro.FontUnderline = True
    Else
        txtText.FontUnderline = False
        picMacro.FontUnderline = False
    End If
    Call TextToPictureBox(txtText, picMacro)
End Sub

Private Sub cmbFonts_Click()
    'in this sub, we're responding to the change of the font name
    'in our combobox. a lot of this code looks unecessary, but it
    'isn't. when different fonts are loaded, they may have a different
    'size, be bolded, italic, etc. we must account for this and
    'adjust our controls accordingly.
    txtText.FontName = cmbFonts.Text
    picMacro.FontName = cmbFonts.Text
    vsbFontSize.Value = txtText.FontSize
    lblFontSize.Caption = vsbFontSize.Value
    If txtText.FontBold = True Then
        chkBold.Value = vbChecked
    Else
        chkBold.Value = vbUnchecked
    End If
    If txtText.FontItalic = True Then
        chkItalic.Value = vbChecked
    Else
        chkItalic.Value = vbUnchecked
    End If
    If txtText.FontUnderline = True Then
        chkUnderline.Value = vbChecked
    Else
        chkUnderline.Value = vbUnchecked
    End If
    Call TextToPictureBox(txtText, picMacro)
End Sub

Private Sub cmdConvert_Click()
    'in this button's event, we're calling to our convert function
    'which converts the picture box to our ascii art.
    txtMacro.Text = Convert(picMacro)
End Sub

Private Sub Form_Load()
    'in our form load, we're loading fonts to our combobox and setting
    'our initial font to arial and it's size to 20. oh, and the lcase
    'on the fonts isn't necessary, i'm just an lcase kind of guy =).
    Dim intFonts As Integer
    cmbFonts.Clear
    For intFonts% = 0 To Screen.FontCount - 1
        cmbFonts.AddItem LCase(Screen.Fonts(intFonts%))
    Next
    cmbFonts.Text = "arial"
    lngFontSize& = 20
End Sub

Private Sub txtText_Change()
    'we draw to the picturebox when the text changes in an effort
    'to update the picturebox as the textbox is being updated.
    Call TextToPictureBox(txtText, picMacro)
End Sub

Private Sub vsbFontSize_Change()
    'with our scroll bar, we're adjusting the size of our font for
    'both the textbox and the picturebox. we will also display that
    'value in our label. and again, we must redraw our text to the
    'picturebox.
    lngFontSize& = vsbFontSize.Value
    lblFontSize.Caption = lngFontSize&
    txtText.FontSize = lngFontSize&
    picMacro.FontSize = lngFontSize&
    Call TextToPictureBox(txtText, picMacro)
End Sub
