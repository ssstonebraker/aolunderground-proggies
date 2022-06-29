VERSION 2.00
Begin Form Form6 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "EMail TOS"
   ClientHeight    =   2205
   ClientLeft      =   3495
   ClientTop       =   1575
   ClientWidth     =   2085
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Height          =   2610
   Left            =   3435
   LinkTopic       =   "Form6"
   ScaleHeight     =   2205
   ScaleWidth      =   2085
   Top             =   1230
   Width           =   2205
   Begin TextBox Text2 
      Height          =   1095
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3240
      Width           =   1575
   End
   Begin SSFrame Frame3D2 
      Caption         =   "Subject Of Mail:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   1580
      Width           =   2055
      Begin TextBox Text1 
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Text            =   "I Need Help!"
         Top             =   240
         Width           =   1815
      End
   End
   Begin SSCommand Command3D8 
      Caption         =   "Exit"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   1200
      Width           =   525
   End
   Begin SSCommand Command3D7 
      Caption         =   "Done"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin CommonDialog CMDialog1 
      Filename        =   "TOSSNs.TES"
      Left            =   480
      Top             =   2640
   End
   Begin SSFrame Frame3D1 
      Caption         =   "TOSEmail SN's:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin SSCommand Command3D6 
         Caption         =   "Load"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   525
      End
      Begin SSCommand Command3D5 
         Caption         =   "Save"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin SSCommand Command3D4 
         Caption         =   "Clr"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   960
         Width           =   375
      End
      Begin SSCommand Command3D3 
         Caption         =   "ñ"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   255
      End
      Begin SSCommand Command3D2 
         Caption         =   "+"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   255
      End
      Begin SSCommand Command3D1 
         Caption         =   "Add Normal TOS SN's"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1845
      End
      Begin SSPanel Panel3D1 
         BevelInner      =   1  'Inset
         BevelWidth      =   2
         BorderWidth     =   1
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1845
         Begin ComboBox Combo2 
            BackColor       =   &H00000000&
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   75
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   75
            Width           =   1700
         End
      End
   End
End
' EMAIL.FRM
Option Explicit
Const c0034 = 1 ' &H1%
Const c005E = 2 ' &H2%
Const c0064 = 12 ' &HC%
Const c0074 = 1 ' &H1%
Const c0096 = 14 ' &HE%
Const c009C = 13 ' &HD%

Sub Command3D1_Click ()
Combo2.AddItem "TOSIM1"
Combo2.AddItem "TOSEmail1"
Combo2.AddItem "TOSEmail2"
Combo2.AddItem "TOSEmail3"
Combo2.AddItem "VLA"
Combo2.AddItem "CGU France"
Combo2.AddItem "COSRoomsDE"
Combo2.AddItem "CosUK"
Combo2.AddItem "HotWheels"
Combo2.AddItem "Lotsen"
Combo2.AddItem "Lotsen AM"
Combo2.AddItem "Lotsen AW"
Combo2.AddItem "Lotsen HO"
Combo2.AddItem "Lotsen MI"
Combo2.AddItem "Lotsen ON"
Combo2.AddItem "TLA CML"
Combo2.AddItem "TLA DJP"
Combo2.AddItem "TLA ECA"
Combo2.AddItem "TLA KNG"
Combo2.AddItem "TLA KRA"
Combo2.AddItem "TLA OWL"
Combo2.AddItem "TLA PEN"
Combo2.AddItem "TLA RMB"
Combo2.AddItem "TLA SIR"
Combo2.AddItem "TLA TAM"
Combo2.AddItem "TOSDonc"
Combo2.AddItem "TOSIRC"
Combo2.AddItem "TOSKids"
Combo2.AddItem "TOSChrisD"
Combo2.AddItem "TOSSpam"
Combo2.AddItem "TROUBLE"
End Sub

Sub Command3D2_Click ()
Dim l0020 As String
Dim l0022 As Variant
l0020$ = InputBox("Please enter what you wanna add")
If  l0020$ = "" Then Exit Sub
For l0022 = 0 To Combo2.ListCount - 1
If  UCase$(Combo2.List(l0022)) Like UCase$("*" + l0020$ + "*") Then Exit Sub
Next
If  l0020$ = "" Then Exit Sub
Combo2.AddItem l0020$
End Sub

Sub Command3D3_Click ()
If  Combo2.Text = "" Then MsgBox "Pick Someone To Remove", 16, "error": Exit Sub
Combo2.RemoveItem Combo2.ListIndex
End Sub

Sub Command3D4_Click ()
Combo2.Clear
End Sub

Sub Command3D5_Click ()
Dim l002C As Integer
Dim l0030 As Variant
' Const c0034 = 1 ' &H1%
If  Combo2.ListCount = 0 Then MsgBox "What you gonna save?", 16, "error": Exit Sub
On Error GoTo L2438A
CMDialog1.DialogTitle = "Save Executions Check List"
CMDialog1.Filter = "List Files (*.Ecl)|*.Ecl| Execution Check List (*.Ecl)|*.Ecl"
CMDialog1.Flags = &H1000&
CMDialog1.Action = 2
If  (CMDialog1.Filetitle <> "") Then
l002C = 2
Open CMDialog1.Filename For Output As l002C
l0030 = 0
Do While l0030 < Combo2.ListCount
Print #l002C, Combo2.List(l0030)
l0030 = l0030 + 1
Loop
Close l002C
End If
Combo2.Clear
L2438A:
Exit Sub
End Sub

Sub Command3D6_Click ()
Dim l0036 As Variant
Dim l003A As Variant
On Error GoTo L244F0
CMDialog1.DialogTitle = "Load Execution Check List"
CMDialog1.Filter = "List Files (*.Ecl)|*.Ecl| Execution Check List Files  (*.Ecl)|*.Ecl"
CMDialog1.Flags = &H1000&
CMDialog1.Action = c0034
l0036 = 1
If  (CMDialog1.Filetitle <> "") Then
Open CMDialog1.Filetitle For Input As l0036
While (EOF(l0036) = False)
Line Input #l0036, l003A
Combo2.AddItem l003A
Wend
Close l0036
End If
L244F0:
Exit Sub
End Sub

Sub Command3D7_Click ()
Dim l0048 As Variant
Dim l004C As Variant
Dim l0050 As Variant
Dim l0056 As Integer
Dim l005A As Integer
Dim l005C As Integer
' Const c005E = 2 ' &H2%
Dim l0062 As Integer
' Const c0064 = 12 ' &HC%
Dim l0068 As Integer
Dim l006A As Integer
Dim l006C As Integer
' Const c0074 = 1 ' &H1%
Dim l0076 As Integer
Dim l007A As Integer
Dim l007C As Integer
Dim l007E As Integer
Dim l0080 As Integer
Dim l0082 As Integer
Dim l0084 As Variant
Dim l0088 As Variant
Dim l008C As Integer
Dim l008E As Integer
Dim l0090 As Integer
Dim l0092 As Integer
Dim l0094 As Integer
' Const c0096 = 14 ' &HE%
Dim l009A As String
' Const c009C = 13 ' &HD%
Dim l009E As Variant
Dim l00A2 As Variant
Dim l00A6 As Variant
Dim l00AA As Integer
Dim l00AC As String
Dim l00AE As Variant
Dim l00B2 As Integer
Dim l00B4 As Variant
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C0 As Variant
Dim l00C4 As Variant
Dim l00C8 As Integer
Dim l00CA As Integer
If  TOSer.Combo1.Text = "" Then MsgBox "Error: Who You Gonna TOS?", 16, "Error": Exit Sub
If  TOSer.Text1.Text = "" Then MsgBox "Error: What Phrase You Gonna Use?", 16, "Error": Exit Sub
If  TOSer.Label7.Caption = "S" Then GoTo L24616:
If  TOSer.Label7.Caption = "s" Then GoTo L2517C:
Exit Sub
L24616:
Call sub950("MM")
Text2.Text = ""
L24654:
Let l004C = Combo2.List(l0048)
l0048 = l0048 + 1
l0050 = Trim(l004C)
If  l0050 = "" Then GoTo L246EC
Text2.Text = Text2.Text + "(" + l0050 + "), "
GoTo L24654

L246EC:
l0056% = extfn1A60("aol frame25", 0&)
Call sub790("Compose Mail")
Do
DoEvents
l005A% = extfn1A98(l0056%, "Compose Mail")
Loop Until l005A% <> 0
l005C% = extfn1A98(l005A%, "To:")
l0062% = extfn1B40(l005C%, c005E)
l0068% = extfn19B8(l0062%, c0064, 0, Text2)
l006A% = extfn1A98(l005A%, "Subject:")
l006C% = extfn1B40(l006A%, c005E)
l0068% = extfn19B8(l006C%, c0064, 0, Text1)
l005C% = extfn1A98(l005A%, "To:")
l0076% = extfn1B40(l005C%, c0074)
l0068% = extfn19B8(l0076%, c0064, 0, TOSer.Combo1 + ":        " + TOSer.Text1)
l007A% = extfn1AD0(l005A%, "_AOL_ICON")
sub368 (l007A%)

L24998:
Do
DoEvents
l0056% = extfn1A60("AOL FRAME25", 0&)
l007C% = extfn1AD0(l0056%, "MDIClient")
l007E% = extfn1AD0(l007C%, "AOL Child")
l0080% = extfn1A98(l007C%, "Error")
l0082% = extfn1A60("#32770", "America Online")
subD08 (.15#)
Loop Until l0080% <> 0 Or l0082% <> 0
If  l0082% <> 0 Then
SendKeys Chr$(32)
L24B2E:
Let l0084 = Combo2.List(l0048)
l0048 = l0048 + 1
l0088 = Trim(l0084)
If  UCase$(l0088) = UCase$(TOSer.Combo1.Text) Then Exit Sub
GoTo L24B2E
End If
If  l0080% <> 0 Then
l0090% = extfn1A98(l007C%, "Error")
l0092% = extfn1AD0(l0090%, "_AOL_View")
l0094% = extfn1A98(l0090%, "Ok")
l008C = extfn1980(l0092%, c0096, 0&, 0&)
l009A$ = String$(l008C + 1, 0)
l008E = extfn19B8(l0092%, c009C, 0&, l009A$)
l009E = Trim$(l009A$)
l00A2 = InStr(1, l009A$, " - ")
If  l00A2 <> 0 Then
l00A6 = Left$(l009A$, l00A2 - 1)
For l00AA% = 1 To Len(l00A6)
sub368 (l0094%)
l00AC$ = Right$(l00A6, l00AA%)
If  Left$(l00AC$, 1) = Chr(10) Or Left$(l00AC$, 1) = Chr(13) Then l00AC$ = Right$(l00A6, l00AA% - 1): Exit For
Next l00AA%
l00AE = UCase(Text2.Text)
If  InStr(l00AE, UCase(l00AC$)) Then
l00B2% = InStr(l00AE, UCase(l00AC$))
l00B4 = Left(l00AE, l00B2% - 1)
l00B8 = Right(l00AE, (Len(l00AE) - (Len(l00AC$) + l00B2% - 1)))
Text2.Text = ""
            For l00BC = 0 To Combo2.ListCount - 1
                    If  InStr(UCase$(l00AC$), UCase$(Combo2.List(l00BC))) <> 0 Then Combo2.RemoveItem (l00BC): GoTo L24EC6
            Next
L24EC6:
For l00C0 = 0 To Combo2.ListCount - 1
l00C4 = Combo2.List(l00C0)
Text2.Text = Text2.Text + "(" + l00C4 + "), "
Next
l005A% = extfn1A98(l0056%, "Compose Mail")
l005C% = extfn1A98(l005A%, "To:")
l0062% = extfn1B40(l005C%, c005E)
l0068% = extfn19B8(l0062%, c0064, 0, Text2)
l006A% = extfn1A98(l005A%, "Subject:")
l006C% = extfn1B40(l006A%, c005E)
l0068% = extfn19B8(l006C%, c0064, 0, Text1)
l005C% = extfn1A98(l005A%, "To:")
l0076% = extfn1B40(l005C%, c0074)
l0068% = extfn19B8(l0076%, c0064, 0, TOSer.Combo1 + ":   " + TOSer.Text1)
l007A% = extfn1AD0(l005A%, "_AOL_ICON")
sub368 (l007A%)
GoTo L24998:
End If
End If
End If

L2517C:
TOSer.Command15.Value = True
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Now TOSing: " + TOSer.Combo1 + "§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Method: EMail - Phrase " + TOSer.Label3 + "§]v^ñï")
Call sub950("MM")
Text2.Text = ""
L252FC:
Let l004C = Combo2.List(l0048)
l0048 = l0048 + 1
l0050 = Trim(l004C)
If  l0050 = "" Then GoTo L25394
Text2.Text = Text2.Text + "(" + l0050 + "), "
GoTo L252FC

L25394:
l0056% = extfn1A60("aol frame25", 0&)
Call sub790("Compose Mail")
Do
DoEvents
l005A% = extfn1A98(l0056%, "Compose Mail")
Loop Until l005A% <> 0
l005C% = extfn1A98(l005A%, "To:")
l0062% = extfn1B40(l005C%, c005E)
l0068% = extfn19B8(l0062%, c0064, 0, Text2)
l006A% = extfn1A98(l005A%, "Subject:")
l006C% = extfn1B40(l006A%, c005E)
l0068% = extfn19B8(l006C%, c0064, 0, Text1)
l005C% = extfn1A98(l005A%, "To:")
l0076% = extfn1B40(l005C%, c0074)
l0068% = extfn19B8(l0076%, c0064, 0, TOSer.Combo1 + ":        " + TOSer.Text1)
l007A% = extfn1AD0(l005A%, "_AOL_ICON")
sub368 (l007A%)

L25640:
Do
DoEvents
l0056% = extfn1A60("AOL FRAME25", 0&)
l007C% = extfn1AD0(l0056%, "MDIClient")
l007E% = extfn1AD0(l007C%, "AOL Child")
l0080% = extfn1A98(l007C%, "Error")
l0082% = extfn1A60("#32770", "America Online")
subD08 (.15#)
Loop Until l0080% <> 0 Or l0082% <> 0
If  l0082% <> 0 Then
subE20
Call subBF0("ïñ^v[§" + TOSer.Combo1 + ", You Have Been EXeCuTeD§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§" + TOSer.Label5 + "§]v^ñï")
L25892:
Let l0084 = Combo2.List(l0048)
l0048 = l0048 + 1
l0088 = Trim(l0084)
If  l0088 = "" Then GoTo L25916
If  UCase$(l0088) = UCase$(TOSer.Combo1.Text) Then Exit Sub
GoTo L25892
L25916:
TOSer.Combo2.AddItem TOSer.Combo1
End If
If  l0080% <> 0 Then
l0090% = extfn1A98(l007C%, "Error")
l0092% = extfn1AD0(l0090%, "_AOL_View")
l0094% = extfn1A98(l0090%, "Ok")
l00C8 = extfn1980(l0092%, c0096, 0&, 0&)
l009A$ = String$(l00C8 + 1, 0)
l00CA = extfn19B8(l0092%, c009C, 0&, l009A$)
l009E = Trim$(l009A$)
l00A2 = InStr(1, l009A$, " - ")
If  l00A2 <> 0 Then
l00A6 = Left$(l009A$, l00A2 - 1)
For l00AA% = 1 To Len(l00A6)
sub368 (l0094%)
l00AC$ = Right$(l00A6, l00AA%)
If  Left$(l00AC$, 1) = Chr(10) Or Left$(l00AC$, 1) = Chr(13) Then l00AC$ = Right$(l00A6, l00AA% - 1): Exit For
Next l00AA%
l00AE = UCase(Text2.Text)
If  InStr(l00AE, UCase(l00AC$)) Then
l00B2% = InStr(l00AE, UCase(l00AC$))
l00B4 = Left(l00AE, l00B2% - 1)
l00B8 = Right(l00AE, (Len(l00AE) - (Len(l00AC$) + l00B2% - 1)))
Text2.Text = ""
            For l00BC = 0 To Combo2.ListCount - 1
                    If  InStr(UCase$(l00AC$), UCase$(Combo2.List(l00BC))) <> 0 Then Combo2.RemoveItem (l00BC): GoTo L25C7C
            Next
L25C7C:
For l00C0 = 0 To Combo2.ListCount - 1
l00C4 = Combo2.List(l00C0)
Text2.Text = Text2.Text + "(" + l00C4 + "), "
Next
l005A% = extfn1A98(l0056%, "Compose Mail")
l005C% = extfn1A98(l005A%, "To:")
l0062% = extfn1B40(l005C%, c005E)
l0068% = extfn19B8(l0062%, c0064, 0, Text2)
l006A% = extfn1A98(l005A%, "Subject:")
l006C% = extfn1B40(l006A%, c005E)
l0068% = extfn19B8(l006C%, c0064, 0, Text1)
l005C% = extfn1A98(l005A%, "To:")
l0076% = extfn1B40(l005C%, c0074)
l0068% = extfn19B8(l0076%, c0064, 0, TOSer.Combo1 + ":   " + TOSer.Text1)
l007A% = extfn1AD0(l005A%, "_AOL_ICON")
sub368 (l007A%)
GoTo L25640:
End If
End If
End If
End Sub

Sub Command3D8_Click ()
Me.Hide
End Sub

Sub Form_Load ()
subCD0 Me
End Sub

Sub sub33D8 ()
Combo2.AddItem fn720()
End Sub
