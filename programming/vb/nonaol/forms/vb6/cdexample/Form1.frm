VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "CommonDialog Example : JBF"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   1920
      Width           =   7215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Open &Multi"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Open"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Printer"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Font"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Color"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://tznproggin.cjb.net"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Let the user change the ForeColor of the Text1 control.
With CommonDialog1
    'Prevent display of the custom color section
    'of the dialog.
    .Flags = cdlCCPreventFullOpen Or cdlCCRGBInit
    .Color = Text1.ForeColor
        .CancelError = False
        .ShowColor
        Text1.ForeColor = .Color
End With
End Sub

Private Sub Command2_Click()
With CommonDialog1
    .Flags = cdlCFScreenFonts Or cdlCFForceFontExist Or cdlCFEffects Or cdlCFLimitSize
    .Min = 8
    .Max = 80
    .FontName = Text1.FontName
    .FontSize = Text1.FontSize
    .FontBold = Text1.FontBold
    .FontItalic = Text1.FontItalic
    .FontUnderline = Text1.FontUnderline
    .FontStrikethru = Text1.FontStrikethru
    .CancelError = False
    .ShowFont
    Text1.FontName = .FontName
    Text1.FontSize = .FontSize
    Text1.FontBold = .FontBold
    Text1.FontItalic = .FontItalic
    Text1.FontUnderline = .FontUnderline
    Text1.FontStrikethru = .FontStrikethru
End With
End Sub

Private Sub Command3_Click()
On Error Resume Next
With CommonDialog1
    'Prepare to print using the Printer object.
    .PrinterDefault = True
    'Disable printing to file and individual page printing.
    .Flags = cdlPDDiablePrintToFile Or cdlPDNoPageNums
    If Text1.SelLength = 0 Then
        'Hide Selection button if tere is no selected text.
        .Flags = .Flags Or cdlPDNoSelection
    Else
        'Else enable the Selection button and make it the default
        'choice.
        .Flags = .Flags Or cdlPDSelection
    End If
    'We need to know whether the user decided to print.
    .CancelError = True
    .ShowPrinter
    If Err = 0 Then
        If .Flags And cdlPDSelection Then
            Printer.Print Text1.SelText
        Else
            Printer.Print Text1.Text
        End If
    End If
End With
End Sub

Private Sub Command4_Click()
Dim Filename As String
If SaveTextControl(Text1, CommonDialog1, Filename) Then
    MsgBox "Text has been saved to file " & Filename
End If
End Sub

Private Sub Command5_Click()
Dim Filename As String
If LoadTextControl(Text1, CommonDialog1, Filename) Then
    MsgBox "Text has been loaded from path " & vbNewLine & " & filename "
End If
End Sub

Private Sub Command6_Click()
Dim Filenames() As String, i As Integer
If SelectMultipleFiles(CommonDialog1, "", Filenames()) Then
    If UBound(Filenames) = 0 Then
        'The Filename property contained only one element.
        Print "Selected file: " & Filenames(0)
    Else
        'The Filename property contained multiple elements.
        Print "Directory name: " & Filenames(0)
        For i = 1 To UBound(Filenames)
            Print "File #" & i & ": " & Filenames(i)
        Next
    End If
End If
End Sub
