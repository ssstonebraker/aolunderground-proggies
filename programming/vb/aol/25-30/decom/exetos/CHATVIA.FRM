VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Chat Violation Option:"
   ClientHeight    =   2730
   ClientLeft      =   2475
   ClientTop       =   1890
   ClientWidth     =   4530
   ControlBox      =   0   'False
   Height          =   3135
   Left            =   2415
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   4530
   Top             =   1545
   Width           =   4650
   Begin SSPanel Panel3D5 
      BevelInner      =   1  'Inset
      BevelWidth      =   2
      BorderWidth     =   1
      ForeColor       =   &H00000000&
      Height          =   575
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   4485
      Begin SSCommand Command3D1 
         Caption         =   "Cancel"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3300
         TabIndex        =   15
         Top             =   285
         Width           =   1095
      End
      Begin SSCommand Command3D11 
         Caption         =   "Done"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   230
         Left            =   3300
         TabIndex        =   14
         Top             =   75
         Width           =   1095
      End
      Begin SSPanel Panel3D6 
         BevelInner      =   1  'Inset
         BevelWidth      =   2
         BorderWidth     =   1
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   75
         TabIndex        =   10
         Top             =   75
         Width           =   3225
         Begin HScrollBar HScroll1 
            Height          =   240
            Left            =   75
            Max             =   8
            TabIndex        =   11
            Top             =   75
            Width           =   3060
         End
      End
   End
   Begin SSFrame Frame3D1 
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin SSPanel Panel3D4 
         BevelInner      =   1  'Inset
         BevelWidth      =   2
         BorderWidth     =   1
         Caption         =   "Panel3D1"
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   4245
         Begin TextBox Text4 
            BackColor       =   &H00000000&
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   75
            TabIndex        =   8
            Top             =   75
            Width           =   4100
         End
      End
      Begin SSPanel Panel3D3 
         BevelInner      =   1  'Inset
         BevelWidth      =   2
         BorderWidth     =   1
         Caption         =   "Panel3D1"
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   4245
         Begin TextBox Text3 
            BackColor       =   &H00000000&
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   75
            TabIndex        =   6
            Top             =   75
            Width           =   4100
         End
      End
      Begin SSPanel Panel3D2 
         BevelInner      =   1  'Inset
         BevelWidth      =   2
         BorderWidth     =   1
         Caption         =   "Panel3D1"
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4245
         Begin TextBox Text2 
            BackColor       =   &H00000000&
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   75
            TabIndex        =   4
            Top             =   75
            Width           =   4100
         End
      End
      Begin SSPanel Panel3D1 
         BevelInner      =   1  'Inset
         BevelWidth      =   2
         BorderWidth     =   1
         Caption         =   "Panel3D1"
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4245
         Begin TextBox Text1 
            BackColor       =   &H00000000&
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   75
            TabIndex        =   2
            Top             =   75
            Width           =   4100
         End
      End
   End
   Begin Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   2880
      Width           =   1095
   End
   Begin Label Label1 
      Caption         =   "0"
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
End
' CHATVIA.FRM
Option Explicit
Const c0042 = 2 ' &H2%
Const c0056 = 12 ' &HC%
Const c006A = 16 ' &H10%

Sub Command2_Click ()
Me.Hide
Exit Sub
End Sub

Sub Command3D1_Click ()
Me.Hide
Exit Sub
End Sub

Sub Command3D11_Click ()
Dim l0038 As Integer
Dim l003C As Integer
Dim l0040 As Integer
' Const c0042 = 2 ' &H2%
Dim l0046 As Integer
Dim l0048 As Integer
Dim l004A As Variant
Dim l004E As Integer
Dim l0050 As Integer
Dim l0052 As Integer
Dim l0054 As Integer
' Const c0056 = 12 ' &HC%
Dim l0060 As Integer
Dim l0062 As Integer
Dim l0064 As Integer
Dim l0066 As Integer
Dim l0068 As Integer
' Const c006A = 16 ' &H10%
Dim l006E As Integer
Dim l0070 As Variant
Dim l0074 As Variant
Dim l0078 As Variant
If  Text1.Text = "" Then MsgBox "What Phrase?", 16, "error": Exit Sub
If  Text2.Text = "" Then MsgBox "What Phrase?", 16, "error": Exit Sub
If  Text3.Text = "" Then MsgBox "What Phrase?", 16, "error": Exit Sub
If  Text4.Text = "" Then MsgBox "What Phrase?", 16, "error": Exit Sub
If  TOSer.Combo1.Text = "" Then MsgBox "Error: Who You Gonna TOS?", 16, "Error": Exit Sub
If  TOSer.Label7.Caption = "S" Then GoTo L27226:
If  TOSer.Label7.Caption = "s" Then GoTo L27804:
Exit Sub
L27226:
Call sub870("GuidePager")
l0038% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l003C% = extfn1A98(l0038%, "I Need Help!")
Loop Until l003C% <> 0
l0040% = extfn1AD0(l003C%, "_AOL_ICON")
l0046% = extfn1B40(l0040%, c0042)
sub368 (l0046%)
Do
DoEvents
l0048% = extfn1A98(l0038%, "Notify AOL")
Loop Until l0048% <> 0
l004A = Int(Rnd * 99 + 1)
l004E% = extfn1A98(l0048%, "Time and Date:")
l0050% = extfn1B40(l004E%, c0042)
l0052% = extfn1AD0(l0048%, "_AOL_EDIT")
Call subC28(l0052%, "Lobby " + l004A)
l0054% = extfn1B40(l0052%, c0042)
l0060% = extfn19B8(l0054%, c0056, 0, Time + "  " + Label2)
l0062% = extfn1B40(l0054%, c0042)
l0064% = extfn1B40(l0062%, c0042)
l0066% = extfn19B8(l0064%, c0056, 0, TOSer.Combo1)
l0068% = extfn1B40(l0064%, c0042)
l0068% = extfn1B40(l0068%, c0042)
l0068% = extfn1B40(l0068%, c0042)
Call subC28(l0068%, TOSer.Combo1 + ":    " + Text1.Text + Chr(13) + TOSer.Combo1 + ":    " + Text2.Text + Chr(13) + TOSer.Combo1 + ":    " + Text3.Text + Chr(13) + TOSer.Combo1 + ":    " + Text4.Text)
subD08 .7#
SendKeys "{tab}" + "{tab}" + "{tab}" + "{ }"
subD08 .7#
l006E% = extfn1A28(l003C%, c006A, 0, 0&)
l006E% = extfn1A28(l0048%, c006A, 0, 0&)
subE20
Me.Hide
L27740:
Let l0074 = TOSer.Combo2.List(l0070)
l0070 = l0070 + 1
l0078 = Trim(l0074)
If  l0078 = "" Then GoTo L277C6
If  UCase$(l0078) = UCase$(TOSer.Combo1) Then Exit Sub
GoTo L27740:
L277C6:
TOSer.Combo2.AddItem TOSer.Combo1
Exit Sub

L27804:
TOSer.Command15.Value = True
subD08 .4#
Call subBF0("•–^v[¤EXeCuTioN¹·º TOSer ßy YußaN¤]v^–•")
subD08 1E-99#
Call subBF0("•–^v[¤Now TOSing: " + TOSer.Combo1 + "¤]v^–•")
subD08 1E-99#
Call subBF0("•–^v[¤Method: Chat Violation - Chat Phrase " + Label1 + "¤]v^–•")
subD08 1E-99#
Call sub870("GuidePager")
l0038% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l003C% = extfn1A98(l0038%, "I Need Help!")
Loop Until l003C% <> 0
l0040% = extfn1AD0(l003C%, "_AOL_ICON")
l0046% = extfn1B40(l0040%, c0042)
sub368 (l0046%)
Do
DoEvents
l0048% = extfn1A98(l0038%, "Notify AOL")
Loop Until l0048% <> 0
l004A = Int(Rnd * 99 + 1)
l004E% = extfn1A98(l0048%, "Time and Date:")
l0050% = extfn1B40(l004E%, c0042)
l0052% = extfn1AD0(l0048%, "_AOL_EDIT")
Call subC28(l0052%, "Lobby " + l004A)
l0054% = extfn1B40(l0052%, c0042)
l0060% = extfn19B8(l0054%, c0056, 0, Time + "  " + Label2)
l0062% = extfn1B40(l0054%, c0042)
l0064% = extfn1B40(l0062%, c0042)
l0066% = extfn19B8(l0064%, c0056, 0, TOSer.Combo1)
l0068% = extfn1B40(l0064%, c0042)
l0068% = extfn1B40(l0068%, c0042)
l0068% = extfn1B40(l0068%, c0042)
Call subC28(l0068%, TOSer.Combo1 + ":    " + Text1.Text + Chr(13) + TOSer.Combo1 + ":    " + Text2.Text + Chr(13) + TOSer.Combo1 + ":    " + Text3.Text + Chr(13) + TOSer.Combo1 + ":    " + Text4.Text)
                        subD08 .7#
SendKeys "{tab}" + "{tab}" + "{tab}" + "{ }"
            subD08 .7#
l006E% = extfn1A28(l003C%, c006A, 0, 0&)
l006E% = extfn1A28(l0048%, c006A, 0, 0&)
subE20
Me.Hide
L27EAA:
Let l0074 = TOSer.Combo2.List(l0070)
l0070 = l0070 + 1
l0078 = Trim(l0074)
If  l0078 = "" Then GoTo L27F30
If  UCase$(l0078) = UCase$(TOSer.Combo1) Then Exit Sub
GoTo L27EAA:
L27F30:
TOSer.Combo2.AddItem TOSer.Combo1
Exit Sub
End Sub

Sub Form_Load ()
Text1.Text = "•-^v•HaZe-X²·² • TOSing • " + fn720()
Text2.Text = "•-^v•HaZe-X²·² • Method • New KO Help"
Text3.Text = "•-^v•HaZe-X²·² • Phrase • Gen."
Text4.Text = "•-^v•HaZe-X²·² • Hush lil fuck."
Label2.Caption = Date
subCD0 Me
End Sub

Sub HScroll1_Change ()
Label1.Caption = HScroll1
End Sub

Sub HScroll1_Scroll ()
Label1.Caption = HScroll1.Value
End Sub

Sub Label1_Change ()
If  Label1.Caption < 1 Then Label1.Caption = "1"
If  Label1.Caption = "1" Then
Text1.Text = "•-^v•HaZe-X²·² • TOSing • " + fn720()
Text2.Text = "•-^v•HaZe-X²·² • Method • New KO Help"
Text3.Text = "•-^v•HaZe-X²·² • Phrase • Gen."
Text4.Text = "•-^v•HaZe-X²·² • Hush lil fuck."
End If
If  Label1.Caption = "2" Then
Text1.Text = "‹\•–¤Unbroken³ TOSer³ ßy: CeD¤–•/›"
Text2.Text = "‹\•–¤Now TOSing-=>" + fn720() + "¤–•/›"
Text3.Text = "‹\•–¤METHOD: Guide Pager*¤–•/›"
Text4.Text = "‹\•–¤Weapon: phrase-=]>1¤–•/›"
End If
If  Label1.Caption = "3" Then
Text1.Text = "•\\'u•–[ Wu-WaRe 97 ToSeR by: OsB"
Text2.Text = "•\\'u•–[ Tos victim: " + fn720()
Text3.Text = "•\\'u•–[ Method: [GuidePager] Phrase: [ Gen ]"
Text4.Text = "•\\'u•–[ Reason: [no reason]"
End If
If  Label1.Caption = "4" Then
Text1.Text = "•›‹v^¥[ (\/)å§§ Ðë§†rúç†ïðñ †ö§êr v¹·º ]¥^v›‹•"
Text2.Text = "•›‹v^¥[ In all life comes death so ]¥^v›‹•"
Text3.Text = "•›‹v^¥[ *" + fn720() + "* - Kids Only ]¥^v›‹•"
Text4.Text = "•›‹v^¥[ I guess ill see ya in Hell ]¥^v›‹•"
End If
If  Label1.Caption = "5" Then
Text1.Text = "–v^§•[CRaZe v¹·º]•§^v– NoW TOSing " + fn720()
Text2.Text = "–v^§•[CRaZe v¹·º]•§^v– TOS MeTHoD [Guide Pager]"
Text3.Text = "–v^§•[CRaZe v¹·º]•§^v– PHRaSe [1]"
Text4.Text = "–v^§•[CRaZe v¹·º]•§^v– Go FuCK YouRSeLF"
End If
If  Label1.Caption = "6" Then
Text1.Text = "°º•¥•º° FaT FroG 1.0 TOSer ßy FuDD °º•¥•º°"
Text2.Text = "°º•¥•º° Now TOSing *" + fn720() + "* °º•¥•º°"
Text3.Text = "°º•¥•º° Method: KO Help - Phrase: #26 °º•¥•º°"
Text4.Text = "°º•¥•º° " + fn720() + ", Yer Ass Is Grass °º•¥•º°"
End If
If  Label1.Caption = "7" Then
Text1.Text = "‹–^vº•[¤Corruption TOS v4.0 ßy Pizza¤]•ºv^–›"
Text2.Text = "‹–^vº•[¤(" + fn720() + ")-(GuidePager)-(34)¤]•ºv^–›"
Text3.Text = "‹–^vº•[¤Corruption v4.0 TOS Complete!¤]•ºv^–›"
Text4.Text = "‹–^vº•[¤AOHell wont save ya this time¤]•ºv^–›"
End If
If  Label1.Caption = "8" Then
Text1.Text = "^v^-^-- -› The EXoRCiST TOSing: " + fn720() + " ‹- --^-^v^"
Text2.Text = "^v^-^-- -› TOSing VIA •KO HELP• ‹- --^-^v^"
Text3.Text = "^v^-^-- -› Catch Ya On The Flip Side ‹- --^-^v^"
Text4.Text = "^v^-^-- -› TOS Complete {S Term} ‹- --^-^v^"
End If
End Sub
