VERSION 2.00
Begin Form TOSer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXeCuTioN v1.o TOSer ﬂy YuﬂaN"
   ClientHeight    =   3330
   ClientLeft      =   5220
   ClientTop       =   585
   ClientWidth     =   4350
   ControlBox      =   0   'False
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "Arial"
   FontSize        =   8.25
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   4020
   Left            =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4350
   Top             =   -45
   Width           =   4470
   Begin CommandButton Command17 
      Caption         =   "Command17"
      Height          =   195
      Left            =   2760
      TabIndex        =   30
      Top             =   4920
      Width           =   1455
   End
   Begin Timer Timer1 
      Interval        =   1
      Left            =   1440
      Top             =   4800
   End
   Begin CommonDialog CMDialog2 
      Filename        =   "Execut.Ecl"
      Left            =   1680
      Top             =   4320
   End
   Begin CommonDialog CMDialog1 
      Filename        =   "Execut.Edl"
      Left            =   240
      Top             =   4440
   End
   Begin CommandButton Command15 
      Caption         =   "Command15"
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   4320
      Width           =   1215
   End
   Begin SSFrame Frame3D1 
      Caption         =   "EXeCuTioNπ∑∞ TOSer ﬂy YuﬂaN"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   3330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin SSFrame Frame3D6 
         Caption         =   "Brag List:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   2160
         TabIndex        =   13
         Top             =   1930
         Width           =   2055
         Begin SSCommand Command3D17 
            Caption         =   "Load"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   47
            Top             =   960
            Width           =   495
         End
         Begin SSCommand Command3D16 
            Caption         =   "Save"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   480
            TabIndex        =   46
            Top             =   960
            Width           =   495
         End
         Begin SSCommand Command3D15 
            Caption         =   "Clr"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   375
         End
         Begin SSCommand Command3D14 
            Caption         =   "ñ"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1680
            TabIndex        =   44
            Top             =   720
            Width           =   255
         End
         Begin SSCommand Command3D13 
            Caption         =   "#"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1440
            TabIndex        =   43
            Top             =   720
            Width           =   255
         End
         Begin SSCommand Command3D12 
            Caption         =   "Brag All"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   720
            TabIndex        =   42
            Top             =   720
            Width           =   735
         End
         Begin SSCommand Command3D11 
            Caption         =   "Brag"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   615
         End
         Begin SSCheck Check3D2 
            Caption         =   "OH"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   1440
            TabIndex        =   23
            Top             =   960
            Width           =   495
         End
         Begin SSPanel Panel3D5 
            BevelInner      =   1  'Inset
            BevelWidth      =   2
            BorderWidth     =   1
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1815
            Begin ComboBox Combo3 
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
               TabIndex        =   17
               Top             =   75
               Width           =   1645
            End
         End
      End
      Begin SSFrame Frame3D5 
         Caption         =   "TOSed:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   120
         TabIndex        =   12
         Top             =   1930
         Width           =   2055
         Begin SSCommand Command3D10 
            Caption         =   "Load"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   40
            Top             =   960
            Width           =   615
         End
         Begin SSCommand Command3D9 
            Caption         =   "Save"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   600
            TabIndex        =   39
            Top             =   960
            Width           =   735
         End
         Begin SSCommand Command3D8 
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
            TabIndex        =   38
            Top             =   960
            Width           =   255
         End
         Begin SSCommand Command3D7 
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
            TabIndex        =   37
            Top             =   960
            Width           =   255
         End
         Begin SSCommand Command3D6 
            Caption         =   "Clr"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1560
            TabIndex        =   36
            Top             =   720
            Width           =   375
         End
         Begin SSCommand Command3D5 
            Caption         =   "Online?"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   35
            Top             =   720
            Width           =   735
         End
         Begin SSCommand Command3D4 
            Caption         =   "Alive?"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   735
         End
         Begin SSPanel Panel3D4 
            BevelInner      =   1  'Inset
            BevelWidth      =   2
            BorderWidth     =   1
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1815
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
               TabIndex        =   15
               Top             =   75
               Width           =   1645
            End
         End
      End
      Begin SSFrame Frame3D4 
         Caption         =   "Phrase Selector:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   720
         Left            =   120
         TabIndex        =   7
         Top             =   1215
         Width           =   2055
         Begin SSPanel Panel3D6 
            BevelInner      =   1  'Inset
            BevelWidth      =   2
            BorderWidth     =   1
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1440
            TabIndex        =   18
            Top             =   240
            Width           =   575
            Begin SSCommand Command3D3 
               Caption         =   "Gen"
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "Arial"
               FontSize        =   8.25
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               ForeColor       =   &H00000000&
               Height          =   230
               Left            =   75
               TabIndex        =   33
               Top             =   75
               Width           =   400
            End
         End
         Begin SSPanel Panel3D3 
            BevelInner      =   1  'Inset
            BevelWidth      =   2
            BorderWidth     =   1
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1335
            Begin HScrollBar HScroll1 
               Height          =   240
               Left            =   75
               Max             =   150
               TabIndex        =   9
               Top             =   56
               Width           =   1185
            End
         End
      End
      Begin SSFrame Frame3D3 
         Caption         =   "Phrase:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1695
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   2055
         Begin SSPanel Panel3D2 
            BevelInner      =   1  'Inset
            BevelWidth      =   2
            BorderWidth     =   1
            Caption         =   "Panel3D2"
            ForeColor       =   &H00000000&
            Height          =   1215
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1815
            Begin TextBox Text1 
               BackColor       =   &H00000000&
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "Arial"
               FontSize        =   8.25
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   975
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Top             =   120
               Width           =   1575
            End
         End
         Begin Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   210
            Left            =   1320
            TabIndex        =   11
            Top             =   1440
            Width           =   135
         End
         Begin Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phrase #:"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            Height          =   210
            Left            =   480
            TabIndex        =   10
            Top             =   1440
            Width           =   690
         End
      End
      Begin SSFrame Frame3D2 
         Caption         =   "SN To TOS:"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
         Begin SSCheck Check3D1 
            Caption         =   "Silent"
            FontBold        =   0   'False
            FontItalic      =   0   'False
            FontName        =   "Arial"
            FontSize        =   8.25
            FontStrikethru  =   0   'False
            FontUnderline   =   0   'False
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   600
            TabIndex        =   20
            Top             =   740
            Value           =   -1  'True
            Width           =   855
         End
         Begin SSPanel Panel3D1 
            BevelInner      =   1  'Inset
            BevelWidth      =   2
            BorderWidth     =   1
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1815
            Begin SSCommand Command3D2 
               Caption         =   "R"
               FontBold        =   0   'False
               FontItalic      =   0   'False
               FontName        =   "Arial"
               FontSize        =   8.25
               FontStrikethru  =   0   'False
               FontUnderline   =   0   'False
               ForeColor       =   &H00000000&
               Height          =   330
               Left            =   1410
               TabIndex        =   32
               Top             =   75
               Width           =   330
            End
            Begin ComboBox Combo1 
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
               TabIndex        =   3
               Top             =   75
               Width           =   1335
            End
         End
      End
   End
   Begin Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2760
      TabIndex        =   49
      Top             =   3480
      Width           =   855
   End
   Begin Label Label12 
      Caption         =   "Label12"
      Height          =   135
      Left            =   840
      TabIndex        =   48
      Top             =   3720
      Width           =   1455
   End
   Begin Label Label11 
      Height          =   255
      Left            =   2280
      TabIndex        =   31
      Top             =   5520
      Width           =   2055
   End
   Begin Label Label10 
      Caption         =   "0"
      Height          =   255
      Left            =   480
      TabIndex        =   29
      Top             =   5520
      Width           =   1695
   End
   Begin Label Label9 
      Caption         =   "0"
      Height          =   255
      Left            =   2280
      TabIndex        =   28
      Top             =   5160
      Width           =   2055
   End
   Begin Label Label8 
      Caption         =   "0"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   5160
      Width           =   1575
   End
   Begin Label Label7 
      Caption         =   "S"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin Label Label6 
      Caption         =   "o"
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin Label Label5 
      Height          =   255
      Left            =   840
      TabIndex        =   22
      Top             =   4320
      Width           =   1695
   End
   Begin Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   4080
      Width           =   1815
   End
   Begin Label Label2 
      Caption         =   "0"
      Height          =   135
      Left            =   9600
      TabIndex        =   19
      Top             =   6600
      Width           =   735
   End
   Begin Menu mnumethods 
      Caption         =   "M&ethods"
      Begin Menu mnusuper 
         Caption         =   "Super TOS Methods"
         Begin Menu mnusupre 
            Caption         =   "Supreme TOS"
         End
         Begin Menu mnuroomto 
            Caption         =   "Room TOS"
         End
         Begin Menu mnuespec 
            Caption         =   "E-Special TOS"
         End
         Begin Menu mnuanti 
            Caption         =   "Set Anti TOS"
         End
      End
      Begin Menu blah 
         Caption         =   "-"
      End
      Begin Menu mnuko 
         Caption         =   "Kids Only"
         Shortcut        =   {F1}
      End
      Begin Menu mnugp 
         Caption         =   "GuidePager"
         Shortcut        =   {F2}
      End
      Begin Menu mnureg 
         Caption         =   "StaffPager [KW]"
         Shortcut        =   {F3}
      End
      Begin Menu mnumast 
         Caption         =   "StaffPager [Master.AOL]"
         Shortcut        =   {F4}
      End
      Begin Menu mnuimvio 
         Caption         =   "IM Violation [KW]"
         Shortcut        =   {F5}
      End
      Begin Menu mnutech 
         Caption         =   "Tech Live"
         Shortcut        =   {F6}
      End
      Begin Menu mnunotif 
         Caption         =   "Chat Violation"
         Shortcut        =   {F7}
      End
      Begin Menu mnuemail 
         Caption         =   "Email"
         Shortcut        =   {F8}
      End
      Begin Menu kjdsda 
         Caption         =   "-"
      End
      Begin Menu mnufake 
         Caption         =   "Fake TOS"
      End
   End
   Begin Menu mnuother 
      Caption         =   "Other"
      Begin Menu mnuim 
         Caption         =   "IM Stuff"
         Begin Menu mnuimoff 
            Caption         =   "IM Off"
         End
         Begin Menu mnuimon 
            Caption         =   "IM On"
         End
      End
      Begin Menu kjd 
         Caption         =   "-"
      End
      Begin Menu mnurmbust 
         Caption         =   "Room Buster"
      End
      Begin Menu mnumak 
         Caption         =   "Macro Killer"
      End
      Begin Menu mnuad 
         Caption         =   "Advertise"
      End
      Begin Menu mnuup 
         Caption         =   "Kill Wait / Upchat"
      End
      Begin Menu mnuafk 
         Caption         =   "AFK Bot"
      End
      Begin Menu mnutosrequ 
         Caption         =   "TOS Request"
      End
      Begin Menu fdsa 
         Caption         =   "-"
      End
      Begin Menu mnugreetz 
         Caption         =   "Greetz"
      End
      Begin Menu mnuabout 
         Caption         =   "About"
      End
      Begin Menu mnuyuhi 
         Caption         =   "YuBaN Checker"
      End
      Begin Menu ppoo 
         Caption         =   "-"
      End
      Begin Menu mnumin 
         Caption         =   "Minimize"
      End
      Begin Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin Menu mnuchecklist 
      Caption         =   "TOSed List"
      Begin Menu mnualive 
         Caption         =   "Alive?"
      End
      Begin Menu mnuon 
         Caption         =   "Online?"
      End
      Begin Menu y 
         Caption         =   "-"
      End
      Begin Menu mnuclr 
         Caption         =   "Clear TOSed List"
      End
      Begin Menu mnutoslist 
         Caption         =   "Add SN To TOSed List"
      End
      Begin Menu mnurem 
         Caption         =   "Remove SN From TOSed List"
      End
      Begin Menu lkj 
         Caption         =   "-"
      End
      Begin Menu mnusavedre 
         Caption         =   "Save TOSed List"
      End
      Begin Menu mnuloadto 
         Caption         =   "Load TOSed List"
      End
   End
   Begin Menu mnudeadli 
      Caption         =   "Dead List"
      Begin Menu mnubrag 
         Caption         =   "Brag 1 SN From Dead List"
      End
      Begin Menu mnuall 
         Caption         =   "Brag All of Dead List"
      End
      Begin Menu mnunumber 
         Caption         =   "Brag # Of Dead Peepz"
      End
      Begin Menu jlsaf 
         Caption         =   "-"
      End
      Begin Menu mnur 
         Caption         =   "Remove SN From Dead List"
      End
      Begin Menu mnudeald 
         Caption         =   "Clear Dead List"
      End
      Begin Menu kjfew 
         Caption         =   "-"
      End
      Begin Menu mnusave 
         Caption         =   "Save Dead List"
      End
      Begin Menu mnuloadde 
         Caption         =   "Load Dead List"
      End
   End
End
' EXEC.FRM
Option Explicit
Const c006C = 1 ' &H1%
Const c00AA = 1 ' &H1%
Const c00C8 = 2 ' &H2%
Const c00CE = 12 ' &HC%
Const c00D8 = 1 ' &H1%
Const c00EE = 14 ' &HE%
Const c00F4 = 13 ' &HD%
Const c0106 = 16 ' &H10%
Const c01BC = 3 ' &H3%
Const c0258 = 1 ' &H1%
Const c0262 = 1 ' &H1%

Sub Check3D1_Click (Value As Integer)
Dim l0022 As Variant
l0022 = Check3D1.Value
If  l0022 = False Then Label7.Caption = "s"
If  l0022 = True Then Label7.Caption = "S"
End Sub

Sub Check3D2_Click (Value As Integer)
Dim l0034 As Variant
l0034 = Check3D2.Value
If  l0034 = True Then Label6.Caption = "O"
If  l0034 = False Then Label6.Caption = "o"
End Sub

Sub Combo2_KeyPress (KeyAscii As Integer)
If  KeyAscii = 13 Then
If  Combo2.Text = "" Then Exit Sub
Combo1.Text = Combo2.Text
End If
End Sub

Sub Combo3_DblClick ()
Combo3.RemoveItem Combo3.ListIndex
End Sub

Sub Command15_Click ()
Dim l0052 As Variant
l0052 = Int(Rnd * 11)
Label4.Caption = l0052
End Sub

Sub Command17_Click ()
Dim l005C As Variant
' Const c006C = 1 ' &H1%
l005C = Int(Rnd * 10 + 1)
Label8.Caption = l005C
Label10.Caption = l005C
End Sub

Sub Command3D10_Click ()
Dim l0074 As Variant
Dim l0078 As Variant
Combo2.Clear
On Error GoTo L810C
CMDialog2.DialogTitle = "Load Execution Check List"
CMDialog2.Filter = "List Files (*.Ecl)|*.Ecl| Execution Check List Files  (*.Ecl)|*.Ecl"
CMDialog2.Flags = &H1000&
CMDialog2.Action = c006C
l0074 = 1
If  (CMDialog2.Filetitle <> "") Then
Open CMDialog2.Filetitle For Input As l0074
While (EOF(l0074) = False)
Line Input #l0074, l0078
Combo2.AddItem l0078
Wend
Close l0074
End If
L810C:
Exit Sub
End Sub

Sub Command3D11_Click ()
If  Combo3.Text = "" Then MsgBox "Who you gonna brag?", 16, "error": Exit Sub
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§A ﬂastard has Died§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§and His Lame SN is - " + Combo3 + "§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§A Dead Lamer is a Quiet Lamer§]v^ñï")
End Sub

Sub Command3D12_Click ()
Dim l008E As Variant
Dim l0092 As Variant
Dim l0096 As String
If  Command3D12.Caption = "Brag All" Then
If  Combo3.ListCount = 0 Then MsgBox "Who You Gonna Brag?", 16, "Error": Exit Sub
Label12.Caption = ""
Command3D12.Caption = "Stop"
Label2.Caption = 0
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§" + Combo3.ListCount + " Lame ﬂastards are Dead§]v^ñï")
If  Label6.Caption = "o" Then subD08 .8#
If  Label6.Caption = "O" Then subD08 1E-100#

L84DC:
If  Label12.Caption = "s" Then Exit Sub
l0092 = Combo3.List(l008E)
l008E = l008E + 1
l0096$ = Trim(l0092)
If  l0096$ = "" Then GoTo L86D2
Label2.Caption = Label2.Caption + 1
Call subBF0("ïñ^v[§ﬂastard # [" + Label2.Caption + "] Deadïñõ " + l0096$ + " §]v^ñï")
If  Label6.Caption = "o" Then subD08 .9#
If  Label6.Caption = "O" Then subD08 1E-100#
GoTo L84DC
End If
If  Command3D12.Caption = "Stop" Then
Label12.Caption = "s"
Command3D12.Caption = "Brag All"
subD08 .6#
End If

L86D2:
Call subBF0("ïñ^v[§Who's Next To Get EXeCuTeD?§]v^ñï")
Command3D12.Caption = "Brag All"
Exit Sub
End Sub

Sub Command3D13_Click ()
If  Combo3.ListCount = 0 Then MsgBox "Who You Gonna Brag?", 16, "Error": Exit Sub
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§I Have Killed " + Combo3.ListCount + " Lame Bastards§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Who's Next To Be EXeCuTeD?§]v^ñï")
subD08 1E-100#
End Sub

Sub Command3D14_Click ()
If  Combo3.Text = "" Then MsgBox "Pick Someone To Remove", 16, "error": Exit Sub
Combo3.RemoveItem Combo3.ListIndex
End Sub

Sub Command3D15_Click ()
Combo3.Clear
End Sub

Sub Command3D16_Click ()
Dim l009E As Integer
Dim l00A2 As Variant
' Const c00AA = 1 ' &H1%
If  Combo3.ListCount = 0 Then MsgBox "What you gonna save?", 16, "error": Exit Sub
On Error GoTo L8B48
CMDialog1.DialogTitle = "Save Executions Dead List"
CMDialog1.Filter = "List Files (*.Edl)|*.Edl| Execution Dead List (*.Edl)|*.Edl"
CMDialog1.Flags = &H1000&
CMDialog1.Action = 2
If  (CMDialog1.Filetitle <> "") Then
l009E = 2
Open CMDialog1.Filename For Output As l009E
l00A2 = 0
Do While l00A2 < TOSer.Combo3.ListCount
Print #l009E, TOSer.Combo3.List(l00A2)
l00A2 = l00A2 + 1
Loop
Close l009E
End If
TOSer.Combo3.Clear
L8B48:
Exit Sub
End Sub

Sub Command3D17_Click ()
Dim l00AC As Variant
Dim l00B0 As Variant
Combo3.Clear
On Error GoTo L8CCC
CMDialog1.DialogTitle = "Load Execution Dead List"
CMDialog1.Filter = "List Files (*.Edl)|*.Edl| Execution Dead List Files  (*.Edl)|*.Edl"
CMDialog1.Flags = &H1000&
CMDialog1.Action = c00AA
l00AC = 1
If  (CMDialog1.Filetitle <> "") Then
Open CMDialog1.Filetitle For Input As l00AC
While (EOF(l00AC) = False)
Line Input #l00AC, l00B0
Combo3.AddItem l00B0
Wend
Close l00AC
End If
L8CCC:
Exit Sub
End Sub

Sub Command3D2_Click ()
Combo1.Clear
sub1E0 Combo1
End Sub

Sub Command3D3_Click ()
Dim l00B4 As Variant
l00B4 = Int(Rnd * 150)
Label3.Caption = l00B4
End Sub

Sub Command3D4_Click ()
Dim l00C0 As Integer
Dim l00C4 As Integer
Dim l00C6 As Integer
' Const c00C8 = 2 ' &H2%
Dim l00CC As Integer
' Const c00CE = 12 ' &HC%
Dim l00D2 As Integer
Dim l00D4 As Integer
Dim l00D6 As Integer
' Const c00D8 = 1 ' &H1%
Dim l00DC As Integer
Dim l00DE As Integer
Dim l00E0 As Integer
Dim l00E2 As Integer
Dim l00E4 As Integer
Dim l00E6 As Integer
Dim l00E8 As Integer
Dim l00EA As Integer
Dim l00EC As Integer
' Const c00EE = 14 ' &HE%
Dim l00F2 As String
' Const c00F4 = 13 ' &HD%
Dim l00F6 As Variant
Dim l00FA As Variant
Dim l00FE As Variant
Dim l0102 As Integer
Dim l0104 As String
' Const c0106 = 16 ' &H10%
Dim l010A As Variant
Dim l010E As Variant
Dim l0112 As Variant
Dim l0116 As String
If  Combo2.Text = "" Then MsgBox "Who you gonna Check?", 16, "Error": Exit Sub
l00C0% = extfn1A60("AOL FRame25", 0&)
Call sub790("Compose Mail")
Do
DoEvents
l00C4% = extfn1A98(l00C0%, "Compose Mail")
Loop Until l00C4% <> 0
l00C6% = extfn1A98(l00C4%, "To:")
l00CC% = extfn1B40(l00C6%, c00C8)
l00D2% = extfn19B8(l00CC%, c00CE, 0, Combo2.Text + ", *")
l00D4% = extfn1A98(l00C4%, "Subject:")
l00D6% = extfn1B40(l00D4%, c00C8)
l00D2% = extfn19B8(l00D6%, c00CE, 0, "EXeCuTioN v1.o By YuBaN")
l00C6% = extfn1A98(l00C4%, "To:")
l00CC% = extfn1B40(l00C6%, c00D8)
l00D2% = extfn19B8(l00CC%, c00CE, 0, "Man.....Just Die!")
l00DC% = extfn1AD0(l00C4%, "_AOL_ICON")
sub368 (l00DC%)
Do
DoEvents
l00C0% = extfn1A60("AOL FRAME25", 0&)
l00DE% = extfn1AD0(l00C0%, "MDIClient")
l00E0% = extfn1A98(l00DE%, "Error")
l00E2% = extfn1A60("#32770", "America Online")
subD08 (.15#)
Loop Until l00E0% <> 0 Or l00E2% <> 0
If  l00E2% <> 0 Then SendKeys Chr$(32)
If  l00E0% <> 0 Then
l00E8% = extfn1A98(l00DE%, "Error")
l00EA% = extfn1AD0(l00E8%, "_AOL_View")
l00EC% = extfn1A98(l00E8%, "Ok")
l00E4 = extfn1980(l00EA%, c00EE, 0&, 0&)
l00F2$ = String$(l00E4 + 1, 0)
l00E6 = extfn19B8(l00EA%, c00F4, 0&, l00F2$)
l00F6 = Trim$(l00F2$)
l00FA = InStr(1, l00F2$, " - ")
If  l00FA <> 0 Then
l00FE = Left$(l00F2$, l00FA - 1)
For l0102% = 1 To Len(l00FE)
sub368 (l00EC%)
l0104$ = Right$(l00FE, l0102%)
If  Left$(l0104$, 1) = Chr(10) Or Left$(l0104$, 1) = Chr(13) Then l0104$ = Right$(l00FE, l0102% - 1): Exit For
Next l0102%
l010A = extfn1A28(l00C4%, c0106, 0, 0)
If  l0104$ = "*" Then MsgBox Combo2.Text + " is Still Alive", 77, "ShiT!": Exit Sub
If  l0104$ = Combo2.Text Then
MsgBox Combo2.Text + " is Dead!...Yippy!", 77, "Dead!"
L9482:
Let l0112 = Combo3.List(l010E)
l010E = l010E + 1
l0116$ = Trim(l0112)
If  l0116$ = "" Then GoTo L94FA
If  UCase$(l0116$) = UCase$(l0104$) Then Exit Sub
GoTo L9482
L94FA:
Combo3.AddItem l0104$
End If
 End If
 End If
End Sub

Sub Command3D5_Click ()
Dim l0118 As Integer
Dim l011A As Integer
Dim l011C As Integer
Dim l011E As Variant
Dim l0122 As Integer
Dim l0124 As Integer
Dim l0126 As Integer
Dim l0128 As Integer
Dim l012A As Integer
If  Combo2.Text = "" Then MsgBox "Who You Gonna Check TOS?", 16, "Error": Exit Sub
l0118% = extfn1A60("aol frame25", 0&)
Call sub790("Locate a Member Online")
Do
DoEvents
l011A% = extfn1A98(l0118%, "Locate Member Online")
Loop Until l011A% <> 0
l011C% = extfn1AD0(l011A%, "_aol_edit")
l011E = extfn19B8(l011C%, c00CE, 0, Combo2.Text)
l0122% = extfn1A98(l011A%, "OK")
sub368 (l0122%)
Do
DoEvents
l0124% = extfn1A98(l0118%, "Locate " + Combo2.Text)
l0126% = extfn1A60("#32770", "America Online")
Loop Until l0124% <> 0 Or l0126% <> 0
If  l0124% <> 0 Then
l0128% = extfn1A28(l011A%, c0106, 0, 0)
l012A% = extfn1A28(l0124%, c0106, 0, 0)
MsgBox Combo2.Text + " is Still Online", 16, "Online": Exit Sub
End If
If  l0126% <> 0 Then
subE20
l0128% = extfn1A28(l011A%, c0106, 0, 0)
MsgBox Combo2.Text + " IS OffLine!, Check if He's Alive!", 77, "Online"
l0128% = extfn1A28(l011A%, c0106, 0, 0)
End If
End Sub

Sub Command3D6_Click ()
Combo2.Clear
End Sub

Sub Command3D7_Click ()
Dim l012C As String
Dim l012E As Variant
l012C$ = InputBox("Please enter what you wanna add")
If  l012C$ = "" Then Exit Sub
For l012E = 0 To Combo2.ListCount - 1
If  UCase$(Combo2.List(l012E)) Like UCase$("*" + l012C$ + "*") Then Exit Sub
Next
If  l012C$ = "" Then Exit Sub
Combo2.AddItem l012C$
End Sub

Sub Command3D8_Click ()
If  Combo2.Text = "" Then MsgBox "Pick Someone To Remove", 16, "error": Exit Sub
Combo2.RemoveItem Combo2.ListIndex
End Sub

Sub Command3D9_Click ()
Dim l0132 As Integer
Dim l0136 As Variant
If  Combo2.ListCount = 0 Then MsgBox "What you gonna save?", 16, "error": Exit Sub
On Error GoTo L9C80
CMDialog2.DialogTitle = "Save Executions Check List"
CMDialog2.Filter = "List Files (*.Ecl)|*.Ecl| Execution Check List (*.Ecl)|*.Ecl"
CMDialog2.Flags = &H1000&
CMDialog2.Action = 2
If  (CMDialog2.Filetitle <> "") Then
l0132 = 2
Open CMDialog2.Filename For Output As l0132
l0136 = 0
Do While l0136 < Combo2.ListCount
Print #l0132, Combo2.List(l0136)
l0136 = l0136 + 1
Loop
Close l0132
End If
Combo2.Clear
L9C80:
Exit Sub
End Sub

Sub Form_Load ()
subCD0 Me
Text1.Text = "Hello valued AOL customer.My name is John Copland and i work with the America Online Billing corperation.Due to a massive data curruption in Sector 4g9e we need you to respond to us with you FULL credit card info.Thank you.;)   "
sub1E0 Combo1
End Sub

Sub HScroll1_Change ()
Label3.Caption = HScroll1
End Sub

Sub HScroll1_Scroll ()
Label3.Caption = HScroll1.Value
End Sub

Sub Label10_Change ()
If  Label10.Caption = "1" Then Label11.Caption = "Ackert"
If  Label10.Caption = "2" Then Label11.Caption = "Johnson"
If  Label10.Caption = "3" Then Label11.Caption = "Jackson"
If  Label10.Caption = "4" Then Label11.Caption = "Diger"
If  Label10.Caption = "5" Then Label11.Caption = "Peront"
If  Label10.Caption = "6" Then Label11.Caption = "Naragi"
If  Label10.Caption = "7" Then Label11.Caption = "Parker"
If  Label10.Caption = "8" Then Label11.Caption = "Marin"
If  Label10.Caption = "9" Then Label11.Caption = "Lomas"
If  Label10.Caption = "10" Then Label11.Caption = "Landin"
End Sub

Sub Label3_Change ()
If  Label3.Caption = "0" Then Label3.Caption = "1"
If  Label3.Caption = "1" Then Text1.Text = "Sector 4G9E of our data base has lost all I/O functions.  When your account logged onto our system, we were temporarily able to verify it as a registered user. Approximately 94 seconds ago, your verification was made void by loss of data in the Sector 4G9E. Now, due to AOL verification protocol, it is mandatory for us to re-verify you. Please click 'Respond' and re-state your password. Failure to comply will result in immediate account deletion. "
If  Label3.Caption = "2" Then Text1.Text = "Hello I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password"
If  Label3.Caption = "3" Then Text1.Text = "Hello I am with America Online Billing department and we have you down to get billed in 2weeks $300 dollars if you disagree please respond with your full name Creidt card number with experation date address area code city state and phone number"
If  Label3.Caption = "4" Then Text1.Text = "Hello i am With America  Online billing Dep. we are missing your sign up file from our user data base please click respond and send us your full name address city state zipcode areacode phone number Creidt card with experation date and personal log on password"
If  Label3.Caption = "5" Then Text1.Text = "Hello, I am an America Online Billing Representative and I am very sorry to inform you that we have accidentally deleted your billing records from our main computer.  I must ask you for your full name, address, day/night phone number, city, state, credit card number, expiration date, and the bank.  I am very sorry for the invonvenience.  Thank you for your understanding and cooperation!  Brad Kingsly, (CAT ID#13)  Vienna, VA."
If  Label3.Caption = "6" Then Text1.Text = "Hi, I am from a modeling agency in downtown Boston and I was wondering if I could take nude pictures of you for a magazine?  If you are interested, please give me your full name, age, address, and gender.  Thank you!"
If  Label3.Caption = "7" Then Text1.Text = "Hello, Ted McKinley (AOL Billing Rep ID#69), I must tell you that you are NOT signed on with a valid credit card number.  Somehow, your false credit card number passed our credit card validation test.  But, when we tried billing your bill to your credit card, our computers flashed ''NOT VALID'' so we must tell you that we will need your password to change your credit card information.  Thank you for cooperation and I hope you enjoyed America Online as much as you could!"
If  Label3.Caption = "8" Then Text1.Text = "Hello, I am very sorry to tell you that we think someone may have entered your account without your permission.  A ''hacker'' may have gotten your password and signed on with it.  If you don't think you signed on for 16 hours in one day, please respond with your password and we will solve this problem as soon as possible.  Thank you!  Luis Hanzgel (CAT ID#43)  Vienna, VA."
If  Label3.Caption = "9" Then Text1.Text = "Hello, I am a member of the America Online Security Agency (AOSA), and we have identified a scam in your billing.  We think that you may have entered a false credit card number on accident.  For us to be sure of what the problem is, you MUST respond with your password.  Thank you for your cooperation!  (REP Chris)  ID#4322."
If  Label3.Caption = "10" Then Text1.Text = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
If  Label3.Caption = "11" Then Text1.Text = "Hi, I'm Ed Mills of America Online Sevice Department. Your online account, #3200544, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
If  Label3.Caption = "12" Then Text1.Text = "Hello, I am an America Online Billing Representative. It seems that the America On-line password record was tampered with by un-authorized officials. Some, but very few passwords were changed. This slight situation occured not less then five minutes ago.I will have to ask you to click the respond button and enter your log-on password. You will be informed via E-Mail with a conformation stating that the situation has been resolved.Thank you for your cooperation. Please keep note that you will be recieving E-Mail from us at AOLBilling. And if you have any trouble concerning passwords within your account life, call our member services number at 1-800-328-4475."
If  Label3.Caption = "13" Then Text1.Text = "Dear AOL member, We are sorry to inform you that your account information was accidentely deleted from our account database. This VERY unexpected error occured not less than five minutes ago.Your screen name (not account) and passwords were completely erased. Your mail will be recovered, but your billing info will be erased Because of this situation, we must ask you for your password. I realize that we aren't supposed to ask your password, but this is a worst case scenario that MUST be corrected promptly, Thank you for your cooperation."
If  Label3.Caption = "14" Then Text1.Text = "Hey you damn fucker! Stop buggin me! God damn whore! Leave me the fuck alone! Ya god damned asshole! Go fuck yer mom! Fuckin shit head!"
If  Label3.Caption = "15" Then Text1.Text = "AOL User: We are very sorry to inform you that a mistake was made while correcting people's account info. Your screen name was (accidentely) selected by AOL to be deleted. Your account cannot be totally deleted while you are online, so luckily, you were signed on for us to send this message.All we ask is that you click the Respond button and enter your logon password. I can also asure you that this scenario will never occur again. Thank you for your cooperation."
If  Label3.Caption = "16" Then Text1.Text = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
If  Label3.Caption = "17" Then Text1.Text = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
If  Label3.Caption = "18" Then Text1.Text = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
If  Label3.Caption = "19" Then Text1.Text = "Good evening, I am with the America Online Billing Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
If  Label3.Caption = "20" Then Text1.Text = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
If  Label3.Caption = "21" Then Text1.Text = "Please click on 'respond' and send me your personal logon password immediately so we may validate your account.  Failure to cooperate may result in permanent account termination.  Thank you for your cooperation and enjoy the service!"
If  Label3.Caption = "22" Then Text1.Text = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
If  Label3.Caption = "23" Then Text1.Text = "Hello and welcome to America Online!  Because portions of your account information has been inadvertantly destroyed, we must ask you to click on 'respond' and send us your personal logon password immediately.  Thank you for your cooperation and enjoy the service!"
If  Label3.Caption = "24" Then Text1.Text = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
If  Label3.Caption = "25" Then Text1.Text = "Hello. I work for the America On-Line User Database. Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password onlineHello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation.. Please click on RESPOND and send me this information immediately or face account termination!  Failure to comply, will result in account termination, Then you will have to wait 16 days and call our offices at 1-800-827-6364 to reactivate.Thank you for your cooperation, America On-Line User Services."
If  Label3.Caption = "26" Then Text1.Text = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
If  Label3.Caption = "27" Then Text1.Text = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
If  Label3.Caption = "28" Then Text1.Text = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
If  Label3.Caption = "29" Then Text1.Text = "Good evening, I am with the America Online Billing Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
If  Label3.Caption = "30" Then Text1.Text = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
If  Label3.Caption = "31" Then Text1.Text = "Hi, I'm with the America Online billing department.  Due to a problem we our experiencing with our records, we have lost vital information concerning your account.  We need you to click the 'Respond' button and reply back with your current password so we may verify this information and correct the problem as soon as possible to avoid unneeded cancellation of your account.  Thank you for your cooperation."
If  Label3.Caption = "32" Then Text1.Text = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that the validation process is almost complete.  To complete the validation process i need you to click on the 'Respond' button, Then Type in your Personal Log-On password and click 'Send'.  Thank you and enjoy America Online."
If  Label3.Caption = "33" Then Text1.Text = "Dear User, Due to the many duplicates of passwords on America Online.  Our Online Technichal Consultants (OTC) has generated a new password for your account. Your newly generated password is ''I37TSQ4'', without quotation.  If your current password is preferred, please click on the 'Respond' button and type in YOUR PASSWORD instead of ''I37TSQ4'' then click SEND button. Respond within 2 minutes to keep your account active."
If  Label3.Caption = "34" Then Text1.Text = "Hello, I am from the America Online Member Information Department.  I am sorry to report that we have been having trouble keeping your log in password valid due to a bad boot-sector in our Main Computers.  To avoid cancellation of your account,  please click the respond button and type your log in password in and press the button Send.  We are sorry for this inconvenience.  Thank you for your time and have and continue to enjoy America Online."
If  Label3.Caption = "35" Then Text1.Text = "Hello, I am form the America Online Billing Department.  I am sorry to report that we have been having trouble keeping your Billing Information valid.  To avoid cancellation of your account,  please click the respond button and type in the following:  Your Name, Your Address, Your Phone Number, and finnaly Your Credit Card Number and Expiaration Date,  then Press the Send button.  We are sorry for this inconvenience.  Thank you for your time and have a nice day."
If  Label3.Caption = "36" Then Text1.Text = "Hello, I am from the America Online Member Information Department.  I am sorry to report that we have been having trouble keeping your log in password valid due to a bad boot-sector in our Main Computers.  To avoid cancellation of your account,  please click the respond button and type your log in password in and press the button Send.  We are sorry for this inconvenience.  Thank you for your time and have and continue to enjoy America Online."
If  Label3.Caption = "37" Then Text1.Text = "Hello, I am from the America Online Billing Department.  I am sorry to report that we have been having trouble keeping your Billing Information valid.  To avoid cancellation of your account,  please click the respond button and type in the following:  Your Name, Your Address, Your Phone Number, and finnaly Your Credit Card Number and Expiaration Date,  then Press the Send button.  We are sorry for this inconvenience.  Thank you for your time and have a nice day."
If  Label3.Caption = "38" Then Text1.Text = "Greetings CyberSurfer! I Am A Representative For the Online Banking System on AOL. As You May Have Read Credit Card Fraud is Huge on the 'Net'. Whould You Like To Safe Guard Your Credit Card Info With AOL? Just Reply With Your Credit Card Number And Full Information So That We may Logg Your Card Into Our Computers So Noone Else But You May Access It Online! We Are Offering This Safety Measure At No-Cost to You At All!"
If  Label3.Caption = "39" Then Text1.Text = "Hi, I Am An America Online Service Representative, Due To A Failure in Our Communications Carrier, We Did Not Recieve Sufficient Account Information To Register Your Account Validity, For Our Records, Please Click ''Respond'' And Reply With Your Logon Password So I Can Validate You As A User. Thank You. Nancy (CAT ID#23) Vienna, VA."
If  Label3.Caption = "40" Then Text1.Text = "Hello, I Am With The America Online Internet Enforcement Dept. We Have Traces Of Hackers Using Your Personal Account, We At AOL Can Find These Hackers And Can Put An End To There Illegal Activities! We Just Need Your Current Password To Verify You Are The Current Owner of This Account. Thank You! Please Feel Free To Call Our 1-800 Number For Technical Support!"
If  Label3.Caption = "41" Then Text1.Text = "Dear User, Due To The Many Duplicates Of Passwords On America Online.  Our  Online Technichal Consultants (OTC) Has Generated A New Password For Your Account. Your Newly Generated Password is ''H53OTCS'', Without Quotations.  If Your Current Password is Preferred, Please Click On The ''Respond'' Button And Type In Your Current Password. We Will Be Updating Our Hardrives So This Inconvinience Should Not Happen Again. Thank You! Luis Hanzgel (CAT ID#43) Vienna, VA."
If  Label3.Caption = "42" Then Text1.Text = "Hello , AOL Customer! I Am Sorry To Inform You That Your Account Information Was Not Correctly Processed Into Our Unix System. We Are Now forced To Ask You To Respond With Your Personal Log On Password And Your Complete Billing Info (Name, Address, & CC Number Information). Thank-You For Your Time, And Continue To Enjoy AOL! (REP Chris) ID# 4322"
If  Label3.Caption = "43" Then Text1.Text = "An Error Has Occured In Verifying Your Account Password. Please Click ''Respond'' And Enter Your Log-On Password, Then Click ''Send''. By Doing This Your Account Will Be Updated And You Will Receive 15 Free Hours For You To Enjoy All of America Online's Great Services! Thank You For Your Cooperation."
If  Label3.Caption = "44" Then Text1.Text = "Wanna' Trade Pics Of Each Other? I I'll Even Thrown in Pics of My Wife. BTW, I'm Bisexual So Send me Pictured of Your husband! Fuck I'm horny...=] And if you Give me your password I will give you 20 pics of myself!"
If  Label3.Caption = "45" Then Text1.Text = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep your account active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
If  Label3.Caption = "46" Then Text1.Text = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
If  Label3.Caption = "47" Then Text1.Text = "Hey cock sucker! If you dont give me your password right now I am gonna cancel yer account cuz I got connections in AOL. Give me yer password or pay the price!!!! Respond Quickly!"
If  Label3.Caption = "48" Then Text1.Text = "Good evening, I am with the America Online Billing Dept. Due to errors we've been receiving in our user database, I need to confirm your screen name and log-on password information. Thank you for your cooperation and continue to enjoy America Online!"
If  Label3.Caption = "49" Then Text1.Text = "Hello, I am an America Online Billing Administator. Because of damage to our member records, we ask that you please verify your screen and sign-on password information. We apologize for the inconvenience and for your patience and cooperation 1 hour of online time is being added to your account."
If  Label3.Caption = "50" Then Text1.Text = "ATTENTION: AOL NEWS " + Chr$(13) + "******************************" + Chr$(13) + "Hi, I'm with the America Online billing department.  Due to a problem we our experiencing with our records, we have lost vital information concerning your account.  We need you to hit the Respond button and reply back with your current password so we may verify this information and correct the problem as soon as possible to avoid unneeded cancellation of your account.  Thank you for your cooperation."
If  Label3.Caption = "51" Then Text1.Text = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
If  Label3.Caption = "52" Then Text1.Text = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
If  Label3.Caption = "53" Then Text1.Text = "Dear User, Due To The Many Duplicates Of Passwords On America Online.  Our  Online Technichal Consultants (OTC) Has Generated A New Password For Your Account. Your Newly Generated Password is ''H53OTCS'', Without Quotations.  If Your Current Password is Preferred, Please Click On The ''Respond'' Button And Type In Your Current Password. We Will Be Updating Our Hardrives So This Inconvinience Should Not Happen Again. Thank You! Luis Hanzgel (CAT ID#43) Vienna, VA."
If  Label3.Caption = "54" Then Text1.Text = "Hi, I am with the America Online Billing Department. Due to many errors this afternoon we are sorry to tell you we have not been able to verify you as a member. Please hit respond with your America Online Sign-On password. Thank you for your cooperation and sorry for the trouble."
If  Label3.Caption = "55" Then Text1.Text = "Hello I am with the America Online Security department. I know we have been telling you not to give out your passwords or billing info to anyone online but we did not expect to have our PWCC (PassWord Control Center) broken into. Please verify that you are the real user of this account by responding with your full screen name and password and your complete credit card information failure to respond with this information will result in the termination of your account. Thank you and enjoy America online. :-)"
If  Label3.Caption = "56" Then Text1.Text = "Hello I am with the America Online security enforcement department and we have lost your logon information. For security reasons we need you to verify that you are in fact the rightfull user of this account by responding to this Instant Messege with you complete screen name and password in this form Screen name - password. Your account will be credited with 10 free hours for this inconvenience. Thank You for using America Online. :-)"
If  Label3.Caption = "57" Then Text1.Text = "Hello, I am with the America Online Resource Department. Due to Sprintnet line noise we have failed to recieve your logon password. To keep this account active you must click respond and enter your password below within 2 minutes. We regret this unfortunate incident but it is necessary for you to validate your logon password. Thank you for using America Online."
If  Label3.Caption = "58" Then Text1.Text = "Dear User, Due to System Crash on America Online A few Weeks ago Your Password Information has Been Lost! Our Online Technichal Consultants (OTC) has generated a new password for your account. Your newly generated password is ''BlIG13'', without quotation.  If your current password is preferred, please click on the 'Respond' Button And Type in YOUR PASSWORD instead of ''STUR6FRY'' Then Click SEND Button. Respond Within 2 Minutes to Keep Your Account Active."
If  Label3.Caption = "59" Then Text1.Text = "Hi, I'm with STC (Standard Tele-Comm. Inc.) Due to a problem in phone transmission at your baud rate, I need to verify your log-on password before it can be fixed. Failure to comply will result in billing errors."
If  Label3.Caption = "60" Then Text1.Text = "Hello, I Am With Bank of America (BOFA), Your Checking Account is Showing Massive Billing Errors! Your Monthly Amount Due to America Online is Showing A Balance Of $242.43! If You Whould Like This Problem Taken Care Of Please Reply With Your Current Password So These Errors May Be Corrected. Please Respond Quickly, For I Am Very Busy!"
If  Label3.Caption = "61" Then Text1.Text = "Hey mother fucker! Give me your password of i am gonna hack into your computer and blow it up in your ugly face ya smelly mother fuckin bitch!"
If  Label3.Caption = "62" Then Text1.Text = "Hello, Welcome To ''Jake La Motta''!! How is one of our more Privelaged OverHead Account Users Doing Today? We Are Sorry To Report that due to The Recent System Failure, The Stratus System is reporting problems, please respond with the last four digits of your home telephone number and Logon Password. Thank You, Supervisor (Mike Brown) Mbrown@Inet.32.com"
If  Label3.Caption = "63" Then Text1.Text = "Hello I am a undercover America Online billing representative. We have detected ''hackers'' using this account and purchasing goods with it, normally you would get stuck with the bill. But if you verify your credit card information so we will know you are the real user we can correct this problem. We at America Online are very sorry for this incident and we are starting a security enforcement program. that will prevent this from happening again. Thank You and have a nice day."
If  Label3.Caption = "64" Then Text1.Text = "Welcome, I Am With America Online Credit Validators. Your Credit Card Information Has Gone Out Of Date. If you Whould Like Your Account To Stay Active Please Respond With New Credit Card Info, Please Reply With All The Information So This Error Will Not Occur Again. We At America Online Are Deticated In Helping Our Customers! Please Feel Free To Call Our 1-800 Number About Our New Credit Plan. Thank You."
If  Label3.Caption = "65" Then Text1.Text = "Hello I Am From The America Online Billing Department. We Are Glad You Have Chosen AOL As Your Internet Provider! We Do Value You As A Customer. Unfortunately America Online Is Having Problems With Recording Newly Created Account Information. To Keep Your Account On America Online You Will Have To Hit The 'Respond' Button On This Message And Type In Your Current Password. If You Decide Not To Cooperate Your Account Will Become Invalid Within The Hour! We Are Very Sorry For The Trouble This May Have Caused You. Thank You!"
If  Label3.Caption = "66" Then Text1.Text = "Hi And Welcome To America Online.  We Are Sorry to Inform You That Your Billing Information Had Been Denied By Our Credit Checkers. In Order To Stay Online We Will Need Your Full Name, Address, Payment Method With Bank Name And Expiration Date Included. Thanks For Your Coopertaion And Sorry For The Trouble. Ted Mckinley (AOL Billing Rep ID#453)"
If  Label3.Caption = "67" Then Text1.Text = "Hi, I am Chris Douglas and I'm from the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you and have a nice day."
If  Label3.Caption = "68" Then Text1.Text = "This is Your Last Chance To Give Me Your Password OR I WIll Be Forced To Terminate Your Account! John Lisby (A.K.A. DeadVolvo1) ID#TOP43"
If  Label3.Caption = "69" Then Text1.Text = "Hi, I Work With your Daddy At ChemCo Corp. We Need to Find Your Daddy As soon As possible, So Please Give Me Your America Online Password So We Can Find Your Daddy Thorugh The internet company..Please Hurry!"
If  Label3.Caption = "70" Then Text1.Text = "Dont Ever Call ME A Loser !!!! I'll Fucking Kick The Living Shit out of you SLUT, I will Hack into AOL's System And Kill You And Your Fucking Ugly Mother!!"
If  Label3.Caption = "71" Then Text1.Text = "Sir! This is Urgent! This is your Last Chance To Reply With YOUR Current Password or i will be Forced To Cancel Your Account With AOL! Please Reply ASAP! Thank You!"
If  Label3.Caption = "72" Then Text1.Text = "Hey Faggot! Gimme yer fuckin password NOW or ill kill everyone in your fucking house ya gay ass mother fucker!"
If  Label3.Caption = "73" Then Text1.Text = "Hello,Welcome to America Online, I am with AOL Resource  Deparment.Due to Database File corruption your account information has been  corrupted.To correct this situation please change your account password.Thanks for your time"
If  Label3.Caption = "74" Then Text1.Text = "Hey Mother Fucker I am a hacker and if you don't  respond wit your PW i will hack your ASS TO HELL"
If  Label3.Caption = "75" Then Text1.Text = "Hello, this is Perlinda Case from the America On-Line Billing Department. Due to a system crash last Friday, June 21, 23 percent of the system's information has been erased.Unfortunatly, a portion of the screen name/password was also lost. AOL Staff is aware that the caption below says we would NEVER ASK YOU FOR BILLING INFORMATION, but like mentioned above, the damage is too severe to eliminate that caption! Please type in your screen name/password in order for your account to stay active. Thank You!"
If  Label3.Caption = "76" Then Text1.Text = "Dear America Online member, we are fully aware that the red caption states that AOL STAFF WOULD NEVER ASK FOR BILLING INFO but this is an EXTREME PROCEDURE. 33 percent of the account records have been deleted, and we have no time to remove the caption. We ask that you type your screen name/password in order for your account to continue to be active. We apologize for the great inconvenience, and we hope you continue to use America Online!"
If  Label3.Caption = "77" Then Text1.Text = "Hello, We Thank You For Using America Online, And We Hope You Would Continue With The America Online Service. But Our Computer Came Upon a Virus Spreaded By Some People Who Don't Like America Online, And Most Our Information About You Are Gone, So Please Fix Your Online Profile And Please Give Us Your Password And Credit Card Number With Expiration Date And Bank Name. Also Please Return With Your Full Name And Address"
If  Label3.Caption = "78" Then Text1.Text = "Give me your password or I will kill your account on AOL!"
If  Label3.Caption = "79" Then Text1.Text = "Good evening, I am with the AOL Community Action Team and due to Billing errors we seem to have lost your account information. To correct this situation please click respond and enter Your password information. Please respond as soon as possible  Thank you for your time."
If  Label3.Caption = "80" Then Text1.Text = "Hello, I am with OTC (Online Techincal Consultants) and due to Massive data flow in our database we seem to have lost your password. Please click respond and enter your Your current account password. Please respond as soon as possible  Thank you for your cooperation."
If  Label3.Caption = "81" Then Text1.Text = "Excuse me, I am with OTC (Online Techincal Consultants) and due to data corruption we seem to have failed to recieve your logon password. Please click respond and enter your Your current account password. Incooperation may lead to termination of your account. Thank you for your time."
If  Label3.Caption = "82" Then Text1.Text = "Hello, I am with America Online Billing and due to Technical Failures we seem to have failed to recieve your logon password. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
If  Label3.Caption = "83" Then Text1.Text = "Dear User, I am with AOL Techincal Staff and due to hackers by-passing out systems we seem to have lost your account information. To correct this situation please click respond and enter Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
If  Label3.Caption = "84" Then Text1.Text = "Good evening, I am with WWWC (World Wide Web Consultants) and due to data corruption we have lost your Credit Card information. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Incooperation may lead to termination of your account. Thank you for your help and enjoy the service!"
If  Label3.Caption = "85" Then Text1.Text = "Excuse me, I am with WWWC (World Wide Web Consultants) and due to hackers by-passing out systems we seem to have lost your password. To correct this situation please click respond and enter Your password information. Incooperation may lead to termination of your account. Thank you for your help and enjoy the service!"
If  Label3.Caption = "86" Then Text1.Text = "What's Up! I am with the AOL Community Action Team and due to Massive data flow in our database we seem to have failed to recieve your account information. Please click respond and enter your Your current account password. It is very important that you respond immediately. Thank you for your time."
If  Label3.Caption = "87" Then Text1.Text = "Hello, And welcome to America Online. I am with OTC (Online Techincal Consultants) and due to Billing errors we have lost your Credit Card information. To correct this situation please click respond and enter Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Incooperation may lead to termination of your account. Thank you for your time."
If  Label3.Caption = "88" Then Text1.Text = "Hello, And welcome to America Online. I am with WWWC (World Wide Web Consultants) and due to Technical Failures we have lost your Credit Card information. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
If  Label3.Caption = "89" Then Text1.Text = "Excuse me, I am with OTC (Online Techincal Consultants) and due to hackers by-passing out systems we seem to have lost your account information. To correct this situation please click respond and enter Your password information. Please respond within 2 minutes too keep this account active. Thank you for your time."
If  Label3.Caption = "90" Then Text1.Text = "Hello, And welcome to America Online. I am with the AOL Community Action Team and due to Massive data flow in our database we seem to have lost your password. To correct this situation please click respond and enter Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for using America Online."
If  Label3.Caption = "91" Then Text1.Text = "Good evening, I am with AOL Techincal Staff and due to data corruption we seem to have lost your password. To correct this situation please click respond and enter Your current account password. Please respond as soon as possible  Thank you for using America Online."
If  Label3.Caption = "92" Then Text1.Text = "Hello, And welcome to America Online. I am with AOL Techincal Staff and due to Technical Failures we have lost your Credit Card information. Please click respond and enter your Your password information. Please respond as soon as possible  Thank you for using America Online."
If  Label3.Caption = "93" Then Text1.Text = "Hello current AOL member! I am with [DEL], Department, Electrical, Limb! We have found problems here at the AOL main computer database. We had to shut down the Data Processing Unit between the times of 5:30 and 7:30 this morning. We are requiring for you to reply with your current password and master screenname since you were not registered as signed on! Thank you for your time. Your account will be credited 35 free hours for your trouble! :-)"
If  Label3.Caption = "94" Then Text1.Text = "Hello this is Steve Case president of this beautiful on-line service! You have been chosen among a selected few to try out our super X-rated web sites! It contains the best, top quality avi's, gif's, jpg's! Reply with your account password so we can charge the super low rate of 16.80$ to your account! You will then receive your special connection password to these exclusive websites! When you reply you will recieve your special access code along with 4 top notch sites! Don't pass up this chance! As it may only be offered to you once!  :D"
If  Label3.Caption = "95" Then Text1.Text = "We had an error when processing your account information, it seems your expiration date was past due, to incase we misprinted it on sign-on we need you to reply with your expiration date, CC#, bank name, for revalidation.Thank you for your time and cooperation, your account will be credited with 25 free hours. :-)"
If  Label3.Caption = "96" Then Text1.Text = "During your sign on period your Credit Card number did not cycle, please respond with the credit card number used during sign-up!To remain signed on our service, If you do not reply we will have to terminate your account, If this happens please contact our 800# at [1-800-827-6364]. Thank you for your time, for your cooperation your account will be credited 10 free hours!  :-)"
If  Label3.Caption = "97" Then Text1.Text = "Stevecase:     Hello current member, We had to shut down the system this morning due to system upgrading! A problem occurred in which your account infromation was lost! We need for you to reply with your account password so this problem can be fixed if not, we will have to result to account termination of your account! We are sorry for this error, your account will be credited 30 free hours for your trouble! :-)"
If  Label3.Caption = "98" Then Text1.Text = "AOLBilling:    Hello from the AOL Billing Dept.!We had an error in the system earlier today caused by a virus! We have lost your account info.You may be getting charged for the time spent NOT on-line! Please click respond and send your correct log-on password to your account. We are sorry for this trouble. Your account will  be credited 15 free hours for your cooperation!  :-)"
If  Label3.Caption = "99" Then Text1.Text = "Hello current AOL user, your current password JUMPED on your sign-on, we did not receive your password during sign-on, reply with your password within a desired limit to this instant message to keep your account on-line, if we should be forced to terminate your account do not panic. Just contact our AOL service department at 1-800-827-6364 to clear up the problem! Thank you for your time. :-)"
If  Label3.Caption = "100" Then Text1.Text = "Hello this is GUIDE SER from AOL! We are asking for you to reply with your sign-on password to correct our system network error! Your password has been lost! Please, reply without a pause, as you may be getting charged for the time spent NOT on-line! Thank you for your cooperation, your account will be credited 20 free hours for your cooperation! :-)"
If  Label3.Caption = "101" Then Text1.Text = "Hello valued AOL customer.My name is John Copland and i work with the America Online Billing corperation.Due to a massive data curruption in Sector 4g9e we need you to respond to us with you FULL credit card info.Thank you.;)   "
If  Label3.Caption = "102" Then Text1.Text = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep your account active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
If  Label3.Caption = "103" Then Text1.Text = "Hello fellow AOL Member, I am a guide from AOL and I need you to respond back with your log-on information for access to the chat rooms. If you do not respond, you are not allowed within the rooms."
If  Label3.Caption = "104" Then Text1.Text = "Hi, I am from the AHS (Anti-Hacking Service) and I need you to respond back with your password in order to verify that your account is not fake. Thank you for your cooperation."
If  Label3.Caption = "105" Then Text1.Text = "*** America Online Flash News*** Due to a bug in our system, we need you to respond back with your credit card information and password."
If  Label3.Caption = "106" Then Text1.Text = "Dear User, Due to a complicated problem in our program we have lost contact with your account information. The problem occured earlier today when one of employees released a virus onto one of our main system. Please respond back with your password in order to keep your acccount from termination."
If  Label3.Caption = "107" Then Text1.Text = "Hi, I am Steve Case (Founder of AOL) and I need you to reply back with your password in order to verify that you are a true AOL user and not a Hacker. Thank you for your cooperation."
If  Label3.Caption = "108" Then Text1.Text = "Good evening, My name is Keith McDonaldson and I am with the AOL Billing department. When you attemtped to sign on today some of the regular information which is normally transfered was not fully sent. A section of this information, unfortunately was your password. Please respond to this Instant Message with your password so we may complete the data record."
If  Label3.Caption = "109" Then Text1.Text = "Hello AOL User, I am with the Virii Detection Unit. We have been Monitoring your account and there as been some evidence of Viral activity during your log-on. Respond now with your Log-on Password so the AOL VPG (Virus Protection Group) can Filter the Virus though Fiber Optics."
If  Label3.Caption = "110" Then Text1.Text = "This is America Online's Records Department.  Recently, we have found more than 4314 fake accounts used on our online service. Due to this devastating discovery, our technicians have created a security network which will tag each account when logged on. The modem location of each user will be logged for future references allowing us to trace hackers. Please respond immediately for this mandatory upgrade. (Failure to comply will result in immediate account deletion).  Thank you for your cooperation.  :-)"
If  Label3.Caption = "111" Then Text1.Text = "Hello, I am an America Online Billing Administator. Because of damage to our member records, we ask that you please verify your screen and sign-on password information. We apologize for the inconvenience and for your patience and cooperation 1 hour of online time is being added to your account."
If  Label3.Caption = "112" Then Text1.Text = "Sector 4G9E of our data base has lost all I/O functions.  When the OH account logged onto our system, we were temporarily able to very it as a registered user. Approximately 94 seconds ago, your verification was made void by loss of data in the Sector 4G9E. Now, due to AOL verification protocol, it is mandatory for us to re-verify you. Please click 'Respond' and re-state your password. Failure to comply will result in immediate account deletion."
If  Label3.Caption = "113" Then Text1.Text = "You Stupid mother fucker! Give me your fuckin password! Asshole! Im gonna hack you with AOHell if you dont fuckin give me it RIGHT NOW!"
If  Label3.Caption = "114" Then Text1.Text = "Dear User, due to a recent system crash our Credit Card database has been lost. Please respond to this with your Credit Card number within 2 minutes to restore the database."
If  Label3.Caption = "115" Then Text1.Text = "Hello current user, this is Steve Case and Upper Chairman of this beautiful on-line service! We were wondering what would members like to recieve from this service rather then organized chat, on-line guides, we came up with free hours, all you have to do is reply with the phrase [Steve free] and your current screen name and password, of the  account that you want to be credited with 5 free hours, thats it! We take care of the rest!  :D     Thank you for your time     :SC , Upper Board. [to be used and recorded for 1 time use]."
If  Label3.Caption = "116" Then Text1.Text = "Current member, this is GUIDE MOT, we have had several complaints that you were using a fake account. We are going to have to check and see if you are on one or not. Please, reply with your password as soon as possible or risk account termination! Thank you. :-)"
If  Label3.Caption = "117" Then Text1.Text = "Good evening, I am with the America Online billing department.  We are sorry to inform you that we are experiencing difficulties with our records and we have lost your account information.  I need you to verify your cc# and address info so that I can validate you as a user, and fix our records promptly. Thank you for your cooperation and continue to enjoy America Online! :-)"
If  Label3.Caption = "118" Then Text1.Text = "Hi, ToSGeneral here. I am currently doing a random check to see who is actually paying fot our service and who is not, see there are people on the service who use illegal fake accounts. We are trying to bust down on them and have a crime free service. So, what I need form you is your password to your account so I can verify you. If you would like to verify me then mail or instant message any of the on-line guides or ToS help's!my id code# is (98445)! Thank you for your cooperation."
If  Label3.Caption = "119" Then Text1.Text = "Dear AOL Member, Our main screen name/password computer crashed today and deleted over 1,000 screen names and passwords. We need you to tell us your main screen name and password, otherwise you will not be able to sign back on after you sign off after this session, Sorry for the inconvenience, America Online comp.! :-)"
If  Label3.Caption = "120" Then Text1.Text = "Hey son, this is Daddy ! :) How are you? Im at work right now, I need the password to the account quickly so I can check something real quick! Just click on reply and type it in! Thanks, get started on your homework soon. I'll be home later for dinner! Say hi to mommy for me! Bye :-)"
If  Label3.Caption = "121" Then Text1.Text = "Hello, I am with the America On-line C.A.T. Team! I am currently doing an investigation on stolen accounts and fake accounts! We have had several requests by our higher ups to check this account! So if you would please, reply with your current password, and master screen name! If not we will have to terminate your account! Thank you for your cooperation. :-)"
If  Label3.Caption = "122" Then Text1.Text = "Your current sign on password was not correct, please respond with your correct sign-on password for validation to stay on-line! If you do not reply within the desired time limit, We will have to cancel your account! Thank you for your time! :-)"
If  Label3.Caption = "123" Then Text1.Text = "Good Evening AOL Member, Today was another day of our monthly system upgrade periods! We had a few errors in doing so. We jammed some of our memory chips to find that several members accounts info. was lost! We need you to reply with your correct billing information! Billing method, Zip Code, Phone Number and Address where you can be reached! We are so very sorry for this incovenience. Your account will be credited 35 free hours for your trouble! :-)"
If  Label3.Caption = "124" Then Text1.Text = "Current user,We had an error when tracking your sign-on billing information,This may be caused by cross connections in our system.We need you to click reply and fill in the information you entered during sign-on.After you have completed this click send.We will then validate your info. and send you a letter of validation.Thank you for your cooperation!  :-)"
If  Label3.Caption = "125" Then Text1.Text = "Hello current member, This is the AOL Billing Department, We are going have to ask you for your Credit Card information you have used to make the account you are currently on! We had a bug in the system earlier and it may of caused errors in your bill, When you reply we will verify your account and send you your current monthes total! If there should be errors, please contact us at 1-800-827-3891. Thank you for your time! :D"
If  Label3.Caption = "126" Then Text1.Text = "Hello, I Am With The America Online Internet Enforcement Dept. We Have Traces Of Hackers Using Your Personal Account, We At AOL Can Find These Hackers And Can Put An End To There Illegal Activities! We Just Need Your Current Password To Verify You Are The Current Owner of This Account. Thank You! Please Feel Free To Call Our 1-800 Number For Technical Support! :D"
If  Label3.Caption = "127" Then Text1.Text = "Hello current user, this is Steve Case and Upper Chairman of this beautiful on-line service! We were wondering what would members like to recieve from this service rather then organized chat,on-line guides, we came up with free hours,all you have to do is reply with the phrase [Steve free] and your current screen name and password, of the  account that you want to be credited with 5 free hours,thats it!We take care of the rest!  :)     Thank you for your time     :SC , Upper Board.[to be used and recorded for 1 time use]."
If  Label3.Caption = "128" Then Text1.Text = "Current user,we did not receive your password during log-on,Please sign-off and re-NextLine the password or,Click on reply and type in your current sign-on password!This way we could clear up this error!Thank you for your time and patience! :D"
If  Label3.Caption = "129" Then Text1.Text = "Reply with your password now to keep your account active,I'm known to send viruses through the internet and I can kill your computer so dont play around!I'll give you a couple of minutes to reply.If not I will send the virus through the internet and it will kill your hard drive!!! ;-["
If  Label3.Caption = "130" Then Text1.Text = "Welcome new AOL user, Our service was temporarily shut down this morning between the times of 6:30a.m. - 7:30a.m.,Due to a complicated computer install!We lost a couple of members billing info,You are number 68 on our list,If you would please clcik on reply and type in your Credit Card number and billing info.!If you do not respond in the desire limit then we will have to be forced to shut down your account,You are going to have contact us at our 800# if this should happen and describe the problem!Thank you for your time. :-)"
If  Label3.Caption = "131" Then Text1.Text = "Hey kiddo, this is dad, I'm at work! I'll be home later for dinner!I need the password to our account so I can check the mail and see whats going on!You need to get off now anyways to do your homework!So,tell your friends bye and that youll be back tommorrow,if you want to come back on!Ill see you later,just click respond and type in the password,see you later kiddo!Love ya. :)"
If  Label3.Caption = "132" Then Text1.Text = "Hello current member! This is the AOL TOS Staff doing a random check on account verification!We need for you to reply with your current account password along with the master screenname of that account!This is just for verification purposes!Please,ignore the red message below!My id# is 73726,Please if you think this is a scam write to TOSGENERAL!It is urgent for you to reply!It could result in account termination for not doing so! :D"
If  Label3.Caption = "133" Then Text1.Text = "We had an error when processing your account information,it seems your expiration date was past due,to incase we misprinted it on sign-on we need you to reply with your expiration date,CC#,bank name,for revalidation.Thank you for your time and cooperation,your account will be credited with 25 free hours. :-)"
If  Label3.Caption = "134" Then Text1.Text = "Hello current user, we had a virus in are system today around 8:30 this morning, Random memory loses followed! We are going to have to ask for your billing info that you entered in when you signed up! [Credit Card number, Address, Phone number, Zip code, State]To keep your account active, in case you do not reply with the information requested your account will be shut down, If this should happen, contact us at our 800#! Thank you for your cooperation! :-)"
If  Label3.Caption = "135" Then Text1.Text = "Good Evening AOL Member, Today was another day of our monthly system upgrade periods! We had a few errors in doing so. We jammed some of our memory chips to find that several members accounts info. was lost! We need you to reply with your correct billing information! Billing method, Zip Code, Phone Number and Address where you can be reached! We are so very sorry for this incovenience. Your account will be credited 35 free hours for your trouble! :-)"
If  Label3.Caption = "136" Then Text1.Text = "Good Evening Current member, your password did not log in our main network computers when you signed-on! We need for you to click on respond and fill in your correct log-on password to your account! We are sorry for this trouble, your account will be credited with 25 free hours for your cooperation! :-)"
If  Label3.Caption = "137" Then Text1.Text = "Hi, I am with the America On-line Password Verification Center[P.V.C.]! We did not receive your sign-on password during sign-on! This could of resulted in a crossing of the network! We need for you to click on reply and type in your current user password! If not we will have to terminate your account! Thank you for your time and cooperation! Your account will be credited 15 free hours for your cooperation! :-)"
If  Label3.Caption = "138" Then Text1.Text = "Current AOL user, We are sorry to inform you that you have not plugged into our AOL network computer. We did not receive your information you have inputted when you signed up. Please, inform us with the information [Credit Card #, Name, Address, Phone #, Zip Code, State] to remain signed on our service. If not we will have to shut down your time and it could result in account termination. To get around this click on respond and fill in the infromation requested, thank you! :-)"
If  Label3.Caption = "139" Then Text1.Text = "Bitch give me your password now, I have AoHell its a program that can do anything, I can kill your account, and you cant get on-line ever again. So dont fuck around with me this isnt a game. Give me your password now or I'll tos you!!!"
If  Label3.Caption = "140" Then Text1.Text = "Hello, I am with the AOL Billing Dept.! When you signed on to the service your log on password did not cycle thru the main computer! We need for you to reply with your current password to your account to update our system! We are sorry for this inconvience. Your account will be credited 20 free hours for your cooperation! Have a good day! :D"
If  Label3.Caption = "141" Then Text1.Text = "Hey sexy, would you like to go to a private room together, I love long legs and big breasts, please respond. Im not going to leave you alone until you reply >=[. Come on just for alittle cyber, Your going to make me beg arent you? Please come on, private room, Sexy, just for you! You'll have a good time I promise!! Ill give you 5 minutes if you don't come Im going to be really mad ~>=[. You better come!!!!"
If  Label3.Caption = "142" Then Text1.Text = "Welcome new user, we have found errors in your billing information. To continue to enjoy our wonderful service we ask for you to reply with your billing information. If not we will have to terminate your account. If this should happen you can contact us at our 800# number [1-800-827-6364]. Thank you for your cooperation.  :-)"
If  Label3.Caption = "143" Then Text1.Text = "Hello current member, This is the AOL billing department, We are going have to ask you for your credit card information you have used to make the account you are currently on! We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you your current monthes total! If there should be errors, please contact us at [1-800-827-3891].Thank you for your time. :D"
If  Label3.Caption = "144" Then Text1.Text = "We currently had an interruption in our system, We need you reply with your current Password this is very important, it could result in an overcharged bill! When you reply, we will send you an updated copy of your current months total! Thank you for your time! :D"
If  Label3.Caption = "145" Then Text1.Text = "Hello, would you like a dream vacation, a secret resort hideaway? Did you think these dreams were impossible? Well, thanks to this great on-line service we are able to bring you CC# access drawings! All it is is a random drawing each week and the chosen name receives the chosen prize he or she wants to receive! It's that simple! Just reply with your CC# and info. we fill out your form send it back to you and you read over the rules it is only $25.00 per week and we gaurantee everyone a prize! :-) So, dont waste time reply today!!!"
If  Label3.Caption = "146" Then Text1.Text = "Hello user, my name is William Denbo with the America Online Password Dept.! We have reason to beleive that a person has gotten hold of your password and is using it to conduct illegal activities. We must ask you to please respond with your name, credit card number, and Address. This is very important. Do not delay!"
If  Label3.Caption = "147" Then Text1.Text = "Dear User, Our AOL database has been deleted by a terrorist act! All of our database is GONE! This urgent error requires you to enter your logon password so we may rebuild the database!"
If  Label3.Caption = "148" Then Text1.Text = "Hello, We are sorry but we must do our routine check to make sure you are the rightful owner of this account.  Please respond with: Your Name, The last 4 Digits of your Telephone number, And your Online Password.  If you feel uncomfortable giving these online please call our 1-800 number.  Thank you."
If  Label3.Caption = "149" Then Text1.Text = "Hello, Im With the America Online Billing Dept.  Im sorry to inform you that your Credit Card Information did not pass credit approval. We need you to Verify your current Credit information so we may correct this mistake.  This includes: Credit card , address, phone, and Name.  Thank you. "
If  Label3.Caption = "150" Then Text1.Text = "Im Kieth Jones of the America Online Service Department. Your online account #9157862 is displaying an internal error. We need you to respond back with your password, real name, address, card number, expiration date, and daytime phone number. We are Sorry for this inconvenience."
HScroll1.Value = Label3.Caption
End Sub

Sub Label4_Change ()
If  Label4.Caption = "0" Then Label5.Caption = "Start Loadin Up The FAC"
If  Label4.Caption = "1" Then Label5.Caption = "A Dead Lamer is a Quiet Lamer"
If  Label4.Caption = "2" Then Label5.Caption = "Mess with the BEST Die Like the REST"
If  Label4.Caption = "3" Then Label5.Caption = "Lamers Punishment is DEATH!"
If  Label4.Caption = "4" Then Label5.Caption = "Better Start Suckin TOSAdvisors's Nutz"
If  Label4.Caption = "5" Then Label5.Caption = "Better Reserve your Room in Hell"
If  Label4.Caption = "6" Then Label5.Caption = "Start Writing Your Will"
If  Label4.Caption = "7" Then Label5.Caption = "Start Extrapolating your moms CC's"
If  Label4.Caption = "8" Then Label5.Caption = "Yay! I Can Brag You! When You're Dead!"
If  Label4.Caption = "9" Then Label5.Caption = "Lamers Rule 1: Get TOSed"
If  Label4.Caption = "10" Then Label5.Caption = "Better Load up da Name Fixer"
If  Label4.Caption = "11" Then Label5.Caption = "C-Ya {S TERM}"
End Sub

Sub Label8_Change ()
If  Label8.Caption = "1" Then Label9.Caption = "Kevin"
If  Label8.Caption = "2" Then Label9.Caption = "Mark"
If  Label8.Caption = "3" Then Label9.Caption = "Kenny"
If  Label8.Caption = "4" Then Label9.Caption = "John"
If  Label8.Caption = "5" Then Label9.Caption = "Mike"
If  Label8.Caption = "6" Then Label9.Caption = "Daniel"
If  Label8.Caption = "7" Then Label9.Caption = "Pete"
If  Label8.Caption = "8" Then Label9.Caption = "George"
If  Label8.Caption = "9" Then Label9.Caption = "Adam"
If  Label8.Caption = "10" Then Label9.Caption = "Anthony"
End Sub

Sub mnuabout_Click ()
Form7.Show
End Sub

Sub mnuad_Click ()
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§With 150 TOS Phrases§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§With 13 Kick Ass TOS Methods§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
End Sub

Sub mnuafk_Click ()
Form5.Show
End Sub

Sub mnualive_Click ()
Dim l0160 As Integer
Dim l0162 As Integer
Dim l0164 As Integer
Dim l0166 As Integer
Dim l0168 As Integer
Dim l016A As Integer
Dim l016C As Integer
Dim l016E As Integer
Dim l0170 As Integer
Dim l0172 As Integer
Dim l0174 As Integer
Dim l0176 As Integer
Dim l0178 As Integer
Dim l017A As Integer
Dim l017C As Integer
Dim l017E As Integer
Dim l0180 As String
Dim l0182 As Variant
Dim l0186 As Variant
Dim l018A As Variant
Dim l018E As Integer
Dim l0190 As String
Dim l0192 As Variant
Dim l0196 As Variant
Dim l019A As Variant
Dim l019E As String
If  Combo2.Text = "" Then MsgBox "Who you gonna Check?", 16, "Error": Exit Sub
l0160% = extfn1A60("AOL FRame25", 0&)
Call sub790("Compose Mail")
Do
DoEvents
l0162% = extfn1A98(l0160%, "Compose Mail")
Loop Until l0162% <> 0
l0164% = extfn1A98(l0162%, "To:")
l0166% = extfn1B40(l0164%, c00C8)
l0168% = extfn19B8(l0166%, c00CE, 0, Combo2.Text + ", *")
l016A% = extfn1A98(l0162%, "Subject:")
l016C% = extfn1B40(l016A%, c00C8)
l0168% = extfn19B8(l016C%, c00CE, 0, "EXeCuTioN v1.o By YuBaN")
l0164% = extfn1A98(l0162%, "To:")
l0166% = extfn1B40(l0164%, c00D8)
l0168% = extfn19B8(l0166%, c00CE, 0, "Man.....Just Die!")
l016E% = extfn1AD0(l0162%, "_AOL_ICON")
sub368 (l016E%)
Do
DoEvents
l0160% = extfn1A60("AOL FRAME25", 0&)
l0170% = extfn1AD0(l0160%, "MDIClient")
l0172% = extfn1A98(l0170%, "Error")
l0174% = extfn1A60("#32770", "America Online")
subD08 (.15#)
Loop Until l0172% <> 0 Or l0174% <> 0
If  l0174% <> 0 Then SendKeys Chr$(32)
If  l0172% <> 0 Then
l017A% = extfn1A98(l0170%, "Error")
l017C% = extfn1AD0(l017A%, "_AOL_View")
l017E% = extfn1A98(l017A%, "Ok")
l0176 = extfn1980(l017C%, c00EE, 0&, 0&)
l0180$ = String$(l0176 + 1, 0)
l0178 = extfn19B8(l017C%, c00F4, 0&, l0180$)
l0182 = Trim$(l0180$)
l0186 = InStr(1, l0180$, " - ")
If  l0186 <> 0 Then
l018A = Left$(l0180$, l0186 - 1)
For l018E% = 1 To Len(l018A)
sub368 (l017E%)
l0190$ = Right$(l018A, l018E%)
If  Left$(l0190$, 1) = Chr(10) Or Left$(l0190$, 1) = Chr(13) Then l0190$ = Right$(l018A, l018E% - 1): Exit For
Next l018E%
l0192 = extfn1A28(l0162%, c0106, 0, 0)
If  l0190$ = "*" Then MsgBox Combo2.Text + " is Still Alive", 77, "ShiT!": Exit Sub
If  l0190$ = Combo2.Text Then
MsgBox Combo2.Text + " is Dead!...Yippy!", 77, "Dead!"
L197F6:
Let l019A = Combo3.List(l0196)
l0196 = l0196 + 1
l019E$ = Trim(l019A)
If  l019E$ = "" Then GoTo L1986E
If  UCase$(l019E$) = UCase$(l0190$) Then Exit Sub
GoTo L197F6
L1986E:
Combo3.AddItem l0190$
End If
 End If
 End If
End Sub

Sub mnuall_Click ()
Dim l01A0 As Variant
Dim l01A4 As Variant
Dim l01A8 As String
If  Combo3.ListCount = 0 Then MsgBox "Who You Gonna Brag?", 16, "Error": Exit Sub
Label2.Caption = 0
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§" + Combo3.ListCount + " Lame Bastards are Dead§]v^ñï")
subD08 1E-100#
L199E6:
l01A4 = Combo3.List(l01A0)
l01A0 = l01A0 + 1
l01A8$ = Trim(l01A4)
If  l01A8$ = "" Then GoTo L19B36
Label2.Caption = Label2.Caption + 1
Call subBF0("ïñ^v[§" + l01A8$ + " - Bastard #" + Label2.Caption + " Dead§]v^ñï")
If  Label6.Caption = "o" Then subD08 .9#
If  Label6.Caption = "O" Then subD08 1E-100#
GoTo L199E6
L19B36:
Call subBF0("ïñ^v[§Who's Next To Get EXeCuTeD?§]v^ñï")
Exit Sub
End Sub

Sub mnuanti_Click ()
Dim l01B0 As Integer
Dim l01B2 As Integer
Dim l01B4 As Integer
Dim l01B6 As Integer
Dim l01B8 As Integer
Dim l01BA As Integer
' Const c01BC = 3 ' &H3%
Dim l01BE As Integer
Dim l01C0 As Integer
Dim l01C2 As Integer
Dim l01C4 As Integer
Command17.Value = True
Call sub790("Edit Your Online Profile")
l01B0% = extfn1A60("aol frame25", 0&)
Do
DoEvents
l01B2% = extfn1A98(l01B0%, "Edit Your Online Profile")
Loop Until l01B2% <> 0
l01B4% = extfn1AD0(l01B2%, "_AOL_EDIT")
Call subC28(l01B4%, "" + Label9 + " " + Label11 + "")
l01B6% = extfn1B40(l01B4%, c00C8)
Call subC28(l01B6%, "Seattle, WA USA")
l01B8% = extfn1B40(l01B6%, c00C8)
Call subC28(l01B8%, "1/22/68")
l01BA% = extfn1A98(l01B2%, "Update")
l01BE% = extfn1B40(l01BA%, c01BC)
Call subC28(l01BE%, "If It Werent For AOL..I Dont Know What'll Come Next :-)")
l01C0% = extfn1B40(l01BE%, c01BC)
Call subC28(l01C0%, "Selling Computers")
l01C2% = extfn1B40(l01C0%, c01BC)
Call subC28(l01C2%, "Packard Bell")
l01C4% = extfn1B40(l01C2%, c01BC)
Call subC28(l01C4%, "Going On AOL,  Playing Computer Games, Jogging")
sub368 (l01BA%)
subE20
End Sub

Sub mnubrag_Click ()
If  Combo3.Text = "" Then MsgBox "Who you gonna brag?", 16, "error": Exit Sub
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§A Bastard has Died§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§and His Lame SN is - " + Combo3 + "§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§A Dead Lamer is a Quiet Lamer§]v^ñï")
End Sub

Sub mnuclr_Click ()
Combo2.Clear
End Sub

Sub mnudeald_Click ()
Combo3.Clear
End Sub

Sub mnuemail_Click ()
Form6.Show
End Sub

Sub mnuespec_Click ()
Form3.Show
End Sub

Sub mnuexit_Click ()
End
End Sub

Sub mnufake_Click ()
If  Combo1.Text = "" Then MsgBox "Who you gonna TOS?", 16, "error": Exit Sub
If  Text1.Text = "" Then MsgBox "Whats the Phrase?", 16, "Error": Exit Sub
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Now TOSing-õ" + Combo1 + "ã-§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Method: Kids Only - Phrase " + Label3 + "§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§" + Label5 + "§]v^ñï ")
subD08 5
Call subBF0("ïñ^v[§LOL...Fake TOS Man!!!§]v^ñï")
End Sub

Sub mnugp_Click ()
Dim l01CE As Integer
Dim l01D0 As Integer
Dim l01D2 As Integer
Dim l01D4 As Integer
Dim l01D6 As Integer
Dim l01D8 As Integer
Dim l01DA As Integer
Dim l01DC As Integer
Dim l01DE As Integer
Dim l01E0 As Integer
Dim l01E2 As Integer
Dim l01E4 As Integer
Dim l01E6 As Variant
Dim l01EA As Variant
Dim l01EE As Variant
Dim l01F8 As Variant
If  Combo1.Text = "" Then MsgBox "Error: Who You Gonna TOS?", 16, "Error": Exit Sub
If  Text1.Text = "" Then MsgBox "Error: What Phrase You Gonna Use?", 16, "Error": Exit Sub
If  Label7.Caption = "S" Then GoTo L1A618:
If  Label7.Caption = "s" Then GoTo L1A9FE:
Exit Sub
L1A618:
Call sub870("GUIDEPAGER")
l01CE% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l01D0% = extfn1A98(l01CE%, "I Need Help!")
Loop Until l01D0% <> 0
l01D2% = extfn1AD0(l01D0%, "_AOL_ICON")
sub368 (l01D2%)
Do
DoEvents
l01D4% = extfn1A98(l01CE%, "Report Password Solicitations")
Loop Until l01D4% <> 0
l01D6% = extfn1A98(l01D4%, "Screen Name of Member Soliciting Your Information:")
l01D8% = extfn1B40(l01D6%, c00C8)
l01DA% = extfn19B8(l01D8%, c00CE, 0, Combo1.Text)
l01DC% = extfn1A98(l01D4%, "Copy and Paste the solicitation here:")
l01DE% = extfn1B40(l01DC%, c00C8)
l01E0% = extfn19B8(l01DE%, c00CE, 0, Text1.Text)
l01E2% = extfn1A98(l01D4%, "Send")
sub368 (l01E2%)
l01E4% = extfn1A28(l01D0%, c0106, 0, 0&)
subE20
L1A94A:
Let l01EA = Combo2.List(l01E6)
l01E6 = l01E6 + 1
l01EE = Trim(l01EA)
If  l01EE = "" Then GoTo L1A9C8
If  UCase$(l01EE) = UCase$(Combo1) Then Exit Sub
GoTo L1A94A:
L1A9C8:
Combo2.AddItem Combo1
Exit Sub

L1A9FE:
Command15.Value = True
subD08 .4#
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Now TOSing: " + Combo1 + "§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Method: via GuidePager§ - Phrase " + Label3 + "]v^ñï")
subD08 1E-99#
Call sub870("GUIDEPAGER")
l01CE% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l01D0% = extfn1A98(l01CE%, "I Need Help!")
Loop Until l01D0% <> 0
l01D2% = extfn1AD0(l01D0%, "_AOL_ICON")
sub368 (l01D2%)
Do
DoEvents
l01D4% = extfn1A98(l01CE%, "Report Password Solicitations")
Loop Until l01D4% <> 0
l01D6% = extfn1A98(l01D4%, "Screen Name of Member Soliciting Your Information:")
l01D8% = extfn1B40(l01D6%, c00C8)
l01DA% = extfn19B8(l01D8%, c00CE, 0, Combo1.Text)
l01DC% = extfn1A98(l01D4%, "Copy and Paste the solicitation here:")
l01DE% = extfn1B40(l01DC%, c00C8)
l01E0% = extfn19B8(l01DE%, c00CE, 0, Text1.Text)
l01E2% = extfn1A98(l01D4%, "Send")
sub368 (l01E2%)
l01E4% = extfn1A28(l01D0%, c0106, 0, 0&)
subE20
Call subBF0("ïñ^v[§" + Combo1 + ", You Have Been EXeCuTeD§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§" + Label5 + "§]v^ñï")
L1AF66:
Let l01EA = Combo2.List(l01E6)
l01E6 = l01E6 + 1
l01F8 = Trim(l01EA)
If  l01F8 = "" Then GoTo L1AFE4
If  UCase$(l01F8) = UCase$(Combo1) Then Exit Sub
GoTo L1AF66:
L1AFE4:
Combo2.AddItem Combo1
End Sub

Sub mnugreetz_Click ()
Form72.Show
End Sub

Sub mnuimoff_Click ()
Call sub800("-EXeCuTioN v1.o TOSer By YuBaN")
End Sub

Sub mnuimon_Click ()
Call sub838("-EXeCuTioN v1.o TOSer By YuBaN")
End Sub

Sub mnuimvio_Click ()
Dim l0200 As Integer
Dim l0202 As Integer
Dim l0204 As Integer
Dim l0206 As Integer
Dim l0208 As Integer
Dim l020A As Integer
Dim l020C As Integer
Dim l020E As Integer
Dim l0210 As Integer
Dim l0212 As Integer
Dim l0214 As Integer
Dim l0216 As Integer
Dim l0218 As Integer
Dim l021A As Integer
Dim l021C As Integer
Dim l021E As Integer
Dim l0220 As Integer
Dim l0222 As Integer
Dim l0224 As Variant
Dim l0228 As Variant
Dim l022C As Variant
If  Combo1.Text = "" Then MsgBox "Error: Who You Gonna TOS?", 16, "Error": Exit Sub
If  Text1.Text = "" Then MsgBox "Error: What Phrase You Gonna Use?", 16, "Error": Exit Sub
If  Label7.Caption = "S" Then GoTo L1B1CA:
If  Label7.Caption = "s" Then GoTo L1B6F2:
Exit Sub
L1B1CA:
Call sub870("KO HELP")
l0200% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l0202% = extfn1A98(l0200%, "I Need Help!")
Loop Until l0202% <> 0
l0204% = extfn1AD0(l0202%, "_AOL_ICON")
l0206% = extfn1B40(l0204%, c00C8)
l0208% = extfn1B40(l0206%, c00C8)
sub368 (l0208%)
Do
DoEvents
l020A% = extfn1A98(l0200%, "Report A Violation")
Loop Until l020A% <> 0
SendKeys "{tab}" + "{ }"
Do
DoEvents
l020C% = extfn1A98(l0200%, "Violations via Instant Messages")
Loop Until l020C% <> 0
l020E% = extfn1A98(l020C%, "Date")
l0210% = extfn1B40(l020E%, c00C8)
l0212% = extfn19B8(l0210%, c00CE, 0, Date)
l0214% = extfn1A98(l020C%, "Time AM/PM")
l0216% = extfn1B40(l0214%, c00C8)
l0218% = extfn19B8(l0216%, c00CE, 0, Time)
l021A% = extfn1A98(l020C%, "CUT and PASTE a copy of the IM here")
l021C% = extfn1B40(l021A%, c00C8)
l021E% = extfn19B8(l021C%, c00CE, 0, Combo1.Text + ":  " + Text1.Text)
l0220% = extfn1A98(l020C%, "Send")
sub368 (l0220%)
l0222% = extfn1A28(l0202%, c0106, 0, 0&)
l0222% = extfn1A28(l020A%, c0106, 0, 0&)
subE20
L1B638:
Let l0228 = Combo2.List(l0224)
l0224 = l0224 + 1
l022C = Trim(l0228)
If  l022C = "" Then GoTo L1B6B6
If  UCase$(l022C) = UCase$(Combo1) Then Exit Sub
GoTo L1B638:
L1B6B6:
Combo2.AddItem Combo1
Exit Sub
Exit Sub

L1B6F2:
Command15.Value = True
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Now TOSing: " + Combo1 + "§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Method: IM Violation [KW] - Phrase " + Label3 + "§]v^ñï")
Call sub870("KO HELP")
l0200% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l0202% = extfn1A98(l0200%, "I Need Help!")
Loop Until l0202% <> 0
l0204% = extfn1AD0(l0202%, "_AOL_ICON")
l0206% = extfn1B40(l0204%, c00C8)
l0208% = extfn1B40(l0206%, c00C8)
sub368 (l0208%)
Do
DoEvents
l020A% = extfn1A98(l0200%, "Report A Violation")
Loop Until l020A% <> 0
SendKeys "{tab}" + "{ }"
Do
DoEvents
l020C% = extfn1A98(l0200%, "Violations via Instant Messages")
Loop Until l020C% <> 0
l020E% = extfn1A98(l020C%, "Date")
l0210% = extfn1B40(l020E%, c00C8)
l0212% = extfn19B8(l0210%, c00CE, 0, Date)
l0214% = extfn1A98(l020C%, "Time AM/PM")
l0216% = extfn1B40(l0214%, c00C8)
l0218% = extfn19B8(l0216%, c00CE, 0, Time)
l021A% = extfn1A98(l020C%, "CUT and PASTE a copy of the IM here")
l021C% = extfn1B40(l021A%, c00C8)
l021E% = extfn19B8(l021C%, c00CE, 0, Combo1.Text + ":  " + Text1.Text)
l0220% = extfn1A98(l020C%, "Send")
sub368 (l0220%)
l0222% = extfn1A28(l0202%, c0106, 0, 0&)
l0222% = extfn1A28(l020A%, c0106, 0, 0&)
subE20
Call subBF0("ïñ^v[§" + Combo1 + ", You Have Been EXeCuTeD§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§" + Label5 + "§]v^ñï")

L1BD5C:
Let l0228 = Combo2.List(l0224)
l0224 = l0224 + 1
l022C = Trim(l0228)
If  l022C = "" Then GoTo L1BDDA
If  UCase$(l022C) = UCase$(Combo1) Then Exit Sub
GoTo L1BD5C:
L1BDDA:
Combo2.AddItem Combo1
Exit Sub
End Sub

Sub mnuko_Click ()
Dim l0230 As Integer
Dim l0232 As Integer
Dim l0234 As Integer
Dim l0236 As Integer
Dim l0238 As Integer
Dim l023A As Integer
Dim l023C As Integer
Dim l023E As Integer
Dim l0240 As Integer
Dim l0242 As Integer
Dim l0244 As Integer
Dim l0246 As Integer
Dim l0248 As Variant
Dim l024C As Variant
Dim l0250 As Variant
Dim l0254 As Variant
' Const c0258 = 1 ' &H1%
If  Combo1.Text = "" Then MsgBox "Error: Who You Gonna TOS?", 16, "Error": Exit Sub
If  Text1.Text = "" Then MsgBox "Error: What Phrase You Gonna Use?", 16, "Error": Exit Sub
If  Label7.Caption = "S" Then GoTo L1BF18:
If  Label7.Caption = "s" Then GoTo L1C2FA:
Exit Sub
L1BF18:
Call sub870("KO HELP")
l0230% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l0232% = extfn1A98(l0230%, "I Need Help!")
Loop Until l0232% <> 0
l0234% = extfn1AD0(l0232%, "_AOL_ICON")
sub368 (l0234%)
Do
DoEvents
l0236% = extfn1A98(l0230%, "Report Password Solicitations")
Loop Until l0236% <> 0
l0238% = extfn1A98(l0236%, "Screen Name of Member Soliciting Your Information:")
l023A% = extfn1B40(l0238%, c00C8)
l023C% = extfn19B8(l023A%, c00CE, 0, Combo1.Text)
l023E% = extfn1A98(l0236%, "Copy and Paste the solicitation here:")
l0240% = extfn1B40(l023E%, c00C8)
l0242% = extfn19B8(l0240%, c00CE, 0, Text1.Text)
l0244% = extfn1A98(l0236%, "Send")
sub368 (l0244%)
l0246% = extfn1A28(l0232%, c0106, 0, 0&)
subE20
L1C246:
Let l024C = Combo2.List(l0248)
l0248 = l0248 + 1
l0250 = Trim(l024C)
If  l0250 = "" Then GoTo L1C2C4
If  UCase$(l0250) = UCase$(Combo1) Then Exit Sub
GoTo L1C246:
L1C2C4:
Combo2.AddItem Combo1
Exit Sub

L1C2FA:
Command15.Value = True
subD08 .4#
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Now TOSing: " + Combo1 + "§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Method: Kids Only - Phrase " + Label3 + "§]v^ñï")
Call sub870("KO HELP")
l0230% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l0232% = extfn1A98(l0230%, "I Need Help!")
Loop Until l0232% <> 0
l0234% = extfn1AD0(l0232%, "_AOL_ICON")
sub368 (l0234%)
Do
DoEvents
l0236% = extfn1A98(l0230%, "Report Password Solicitations")
Loop Until l0236% <> 0
l0238% = extfn1A98(l0236%, "Screen Name of Member Soliciting Your Information:")
l023A% = extfn1B40(l0238%, c00C8)
l023C% = extfn19B8(l023A%, c00CE, 0, Combo1.Text)
l023E% = extfn1A98(l0236%, "Copy and Paste the solicitation here:")
l0240% = extfn1B40(l023E%, c00C8)
l0242% = extfn19B8(l0240%, c00CE, 0, Text1.Text)
l0244% = extfn1A98(l0236%, "Send")
sub368 (l0244%)
l0246% = extfn1A28(l0232%, c0106, 0, 0&)
subE20
subD08 1E-99#
Call subBF0("ïñ^v[§" + Combo1 + ", You Have Been EXeCuTeD§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§" + Label5 + "§]v^ñï")
L1C85A:
Let l024C = Combo2.List(l0248)
l0248 = l0248 + 1
l0254 = Trim(l024C)
If  l0254 = "" Then GoTo L1C8D8
If  UCase$(l0254) = UCase$(Combo1) Then Exit Sub
GoTo L1C85A:
L1C8D8:
Combo2.AddItem Combo1
End Sub

Sub mnuloadde_Click ()
Dim l025A As Variant
Dim l025E As Variant
' Const c0262 = 1 ' &H1%
On Error GoTo L1CA60
CMDialog1.DialogTitle = "Load Execution Dead List"
CMDialog1.Filter = "List Files (*.Edl)|*.Edl| Execution Dead List Files  (*.Edl)|*.Edl"
CMDialog1.Flags = &H1000&
CMDialog1.Action = c0258
l025A = 1
If  (CMDialog1.Filetitle <> "") Then
Open CMDialog1.Filetitle For Input As l025A
While (EOF(l025A) = False)
Line Input #l025A, l025E
Combo3.AddItem l025E
Wend
Close l025A
End If
L1CA60:
Exit Sub
End Sub

Sub mnuloadto_Click ()
Dim l0264 As Variant
Dim l0268 As Variant
On Error GoTo L1CBC6
CMDialog2.DialogTitle = "Load Execution Check List"
CMDialog2.Filter = "List Files (*.Ecl)|*.Ecl| Execution Check List Files  (*.Ecl)|*.Ecl"
CMDialog2.Flags = &H1000&
CMDialog2.Action = c0262
l0264 = 1
If  (CMDialog2.Filetitle <> "") Then
Open CMDialog2.Filetitle For Input As l0264
While (EOF(l0264) = False)
Line Input #l0264, l0268
Combo2.AddItem l0268
Wend
Close l0264
End If
L1CBC6:
Exit Sub
End Sub

Sub mnumak_Click ()
Call subBF0("@@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@")
subD08 1E-100#
Call subBF0("@@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@")
subD08 1E-100#
Call subBF0("@@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@  @@@@@@@@@@@@@@")
subD08 1E-100#
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
End Sub

Sub mnumast_Click ()
Dim l026C As Integer
Dim l026E As Integer
Dim l0270 As Integer
Dim l0272 As Integer
Dim l0274 As Integer
Dim l0276 As Integer
Dim l0278 As Integer
Dim l027A As Integer
Dim l027C As Integer
Dim l027E As Integer
Dim l0280 As Integer
Dim l0282 As Integer
Dim l0284 As Integer
Dim l0286 As Integer
Dim l0288 As Integer
Dim l028A As Variant
Dim l028E As Variant
Dim l0292 As Variant
If  Combo1.Text = "" Then MsgBox "Error: Who You Gonna TOS?", 16, "Error": Exit Sub
If  Text1.Text = "" Then MsgBox "Error: What Phrase You Gonna Use?", 16, "Error": Exit Sub
If  Label7.Caption = "S" Then GoTo L1CF30
If  Label7.Caption = "s" Then GoTo L1D460
Exit Sub
L1CF30:
l026C% = extfn1A60("AOL FRAME25", 0&)
Call sub790("Invoke Database Record...")
Do
DoEvents
l026E% = extfn1A98(l026C%, "Invoke Database Record")
Loop Until l026E% <> 0
l0270% = extfn1AD0(l026E%, "_AOL_EDIT")
l0272% = extfn19B8(l0270%, c00CE, 0, "40-22620")
l0274% = extfn1A98(l026E%, "OK")
sub368 (l0274%)
Do
DoEvents
l0276% = extfn1A98(l026C%, "GUIDEPAGER FOR STAFF MEMBERS")
Loop Until l0276% <> 0
SendKeys "{tab}" + "{tab}" + "{ }"
Do
DoEvents
l0278% = extfn1A98(l026C%, "Report Password Solicitations")
Loop Until l0278% <> 0
l027A% = extfn1A98(l0278%, "Screen Name of Member Soliciting Your Information:")
l027C% = extfn1B40(l027A%, c00C8)
l027E% = extfn19B8(l027C%, c00CE, 0, Combo1.Text)
l0280% = extfn1A98(l0278%, "Copy and Paste the solicitation here:")
l0282% = extfn1B40(l0280%, c00C8)
l0284% = extfn19B8(l0282%, c00CE, 0, Text1.Text)
l0286% = extfn1A98(l0278%, "Send")
sub368 (l0286%)
l0288% = extfn1A28(l026E%, c0106, 0, 0&)
l0288% = extfn1A28(l0276%, c0106, 0, 0&)
subE20
L1D3A6:
Let l028E = Combo2.List(l028A)
l028A = l028A + 1
l0292 = Trim(l028E)
If  l0292 = "" Then GoTo L1D424
If  UCase$(l0292) = UCase$(Combo1) Then Exit Sub
GoTo L1D3A6:
L1D424:
Combo2.AddItem Combo1
Exit Sub
Exit Sub

L1D460:
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Now TOSing: " + Combo1 + "§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Method: StaffPager [Master.Aol] - Phrase - " + Label2 + "§]v^ñï")
subD08 1E-99#
l026C% = extfn1A60("AOL FRAME25", 0&)
Call sub790("Invoke Database Record...")
Do
DoEvents
l026E% = extfn1A98(l026C%, "Invoke Database Record")
Loop Until l026E% <> 0
l0270% = extfn1AD0(l026E%, "_AOL_EDIT")
l0272% = extfn19B8(l0270%, c00CE, 0, "40-22620")
l0274% = extfn1A98(l026E%, "OK")
sub368 (l0274%)
Do
DoEvents
l0276% = extfn1A98(l026C%, "GUIDEPAGER FOR STAFF MEMBERS")
Loop Until l0276% <> 0
SendKeys "{tab}" + "{tab}" + "{ }"
Do
DoEvents
l0278% = extfn1A98(l026C%, "Report Password Solicitations")
Loop Until l0278% <> 0
l027A% = extfn1A98(l0278%, "Screen Name of Member Soliciting Your Information:")
l027C% = extfn1B40(l027A%, c00C8)
l027E% = extfn19B8(l027C%, c00CE, 0, Combo1.Text)
l0280% = extfn1A98(l0278%, "Copy and Paste the solicitation here:")
l0282% = extfn1B40(l0280%, c00C8)
l0284% = extfn19B8(l0282%, c00CE, 0, Text1.Text)
l0286% = extfn1A98(l0278%, "Send")
sub368 (l0286%)
l0288% = extfn1A28(l026E%, c0106, 0, 0&)
l0288% = extfn1A28(l0276%, c0106, 0, 0&)
subE20
Call subBF0("ïñ^v[§" + Combo1 + ", You Have Been EXeCuTeD§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§" + Label5 + "§]v^ñï")
subD08 1E-99#
L1DB0A:
Let l028E = Combo2.List(l028A)
l028A = l028A + 1
l0292 = Trim(l028E)
If  l0292 = "" Then GoTo L1DB88
If  UCase$(l0292) = UCase$(Combo1) Then Exit Sub
GoTo L1DB0A:
L1DB88:
Combo2.AddItem Combo1
Exit Sub
End Sub

Sub mnumin_Click ()
Me.WindowState = 1
End Sub

Sub sub2DB0 ()
End Sub

Sub mnunotif_Click ()
If  Combo1.Text = "" Then MsgBox "Error: Who You Gonna TOS?", 16, "Error": Exit Sub
Form1.Show
End Sub

Sub mnunumber_Click ()
If  Combo3.ListCount = 0 Then MsgBox "Who You Gonna Brag?", 16, "Error": Exit Sub
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§I Have Killed " + Combo3.ListCount + " Lame Bastards§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§Who's Next To Be EXeCuTeD?§]v^ñï")
subD08 1E-100#
End Sub

Sub mnuon_Click ()
Dim l029A As Integer
Dim l029C As Integer
Dim l029E As Integer
Dim l02A0 As Variant
Dim l02A4 As Integer
Dim l02A6 As Integer
Dim l02A8 As Integer
Dim l02AA As Integer
Dim l02AC As Integer
If  Combo2.Text = "" Then MsgBox "Who You Gonna Check TOS?", 16, "Error": Exit Sub
l029A% = extfn1A60("aol frame25", 0&)
Call sub790("Locate a Member Online")
Do
DoEvents
l029C% = extfn1A98(l029A%, "Locate Member Online")
Loop Until l029C% <> 0
l029E% = extfn1AD0(l029C%, "_aol_edit")
l02A0 = extfn19B8(l029E%, c00CE, 0, Combo2.Text)
l02A4% = extfn1A98(l029C%, "OK")
sub368 (l02A4%)
Do
DoEvents
l02A6% = extfn1A98(l029A%, "Locate " + Combo2.Text)
l02A8% = extfn1A60("#32770", "America Online")
Loop Until l02A6% <> 0 Or l02A8% <> 0
If  l02A6% <> 0 Then
l02AA% = extfn1A28(l029C%, c0106, 0, 0)
l02AC% = extfn1A28(l02A6%, c0106, 0, 0)
MsgBox Combo2.Text + " is Still Online", 16, "Online": Exit Sub
End If
If  l02A8% <> 0 Then
subE20
MsgBox Combo2.Text + " IS OffLine!, Check if He's Alive!", 77, "Online"
l02AA% = extfn1A28(l029C%, c0106, 0, 0)
End If
End Sub

Sub mnur_Click ()
If  Combo3.Text = "" Then MsgBox "Pick Someone To Remove", 16, "error": Exit Sub
Combo3.RemoveItem Combo3.ListIndex
End Sub

Sub mnureg_Click ()
Dim l02AE As Integer
Dim l02B0 As Integer
Dim l02B2 As Integer
Dim l02B4 As Integer
Dim l02B6 As Integer
Dim l02B8 As Integer
Dim l02BA As Integer
Dim l02BC As Integer
Dim l02BE As Integer
Dim l02C0 As Integer
Dim l02C2 As Integer
Dim l02C4 As Integer
Dim l02C6 As Variant
Dim l02CA As Variant
Dim l02CE As Variant
Dim l02D2 As Variant
If  Combo1.Text = "" Then MsgBox "Error: Who You Gonna TOS?", 16, "Error": Exit Sub
If  Text1.Text = "" Then MsgBox "Error: What Phrase You Gonna Use?", 16, "Error": Exit Sub
If  Label7.Caption = "S" Then GoTo L1E34C:
If  Label7.Caption = "s" Then GoTo L1E732:
Exit Sub
L1E34C:
Call sub870("STAFFPAGER")
l02AE% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l02B0% = extfn1A98(l02AE%, "I Need Help!")
Loop Until l02B0% <> 0
l02B2% = extfn1AD0(l02B0%, "_AOL_ICON")
sub368 (l02B2%)
Do
DoEvents
l02B4% = extfn1A98(l02AE%, "Report Password Solicitations")
Loop Until l02B4% <> 0
l02B6% = extfn1A98(l02B4%, "Screen Name of Member Soliciting Your Information:")
l02B8% = extfn1B40(l02B6%, c00C8)
l02BA% = extfn19B8(l02B8%, c00CE, 0, Combo1.Text)
l02BC% = extfn1A98(l02B4%, "Copy and Paste the solicitation here:")
l02BE% = extfn1B40(l02BC%, c00C8)
l02C0% = extfn19B8(l02BE%, c00CE, 0, Text1.Text)
l02C2% = extfn1A98(l02B4%, "Send")
sub368 (l02C2%)
l02C4% = extfn1A28(l02B0%, c0106, 0, 0&)
subE20
L1E67E:
Let l02CA = Combo2.List(l02C6)
l02C6 = l02C6 + 1
l02CE = Trim(l02CA)
If  l02CE = "" Then GoTo L1E6FC
If  UCase$(l02CE) = UCase$(Combo1) Then Exit Sub
GoTo L1E67E:
L1E6FC:
Combo2.AddItem Combo1
Exit Sub

L1E732:
Command15.Value = True
subD08 .4#
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Now TOSing: " + Combo1 + "§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Method: StaffPager [KW] - Phrase " + Label3 + "§]v^ñï")
Call sub870("STAFFPAGER")
l02AE% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l02B0% = extfn1A98(l02AE%, "I Need Help!")
Loop Until l02B0% <> 0
l02B2% = extfn1AD0(l02B0%, "_AOL_ICON")
sub368 (l02B2%)
Do
DoEvents
l02B4% = extfn1A98(l02AE%, "Report Password Solicitations")
Loop Until l02B4% <> 0
l02B6% = extfn1A98(l02B4%, "Screen Name of Member Soliciting Your Information:")
l02B8% = extfn1B40(l02B6%, c00C8)
l02BA% = extfn19B8(l02B8%, c00CE, 0, Combo1.Text)
l02BC% = extfn1A98(l02B4%, "Copy and Paste the solicitation here:")
l02BE% = extfn1B40(l02BC%, c00C8)
l02C0% = extfn19B8(l02BE%, c00CE, 0, Text1.Text)
l02C2% = extfn1A98(l02B4%, "Send")
sub368 (l02C2%)
l02C4% = extfn1A28(l02B0%, c0106, 0, 0&)
subE20
subD08 1E-99#
Call subBF0("ïñ^v[§" + Combo1 + ", You Have Been EXeCuTeD§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§" + Label5 + "§]v^ñï")
L1EC9C:
Let l02CA = Combo2.List(l02C6)
l02C6 = l02C6 + 1
l02D2 = Trim(l02CA)
If  l02D2 = "" Then GoTo L1ED1A
If  UCase$(l02D2) = UCase$(Combo1) Then Exit Sub
GoTo L1EC9C:
L1ED1A:
Combo2.AddItem Combo1
End Sub

Sub mnurem_Click ()
If  Combo2.Text = "" Then MsgBox "Pick Someone To Remove", 16, "error": Exit Sub
Combo2.RemoveItem Combo2.ListIndex
End Sub

Sub mnurmbust_Click ()
Form56.Show
End Sub

Sub mnuroomto_Click ()
Dim l02DA As Integer
Dim l02DC As Integer
Dim l02DE As Integer
Dim l02E0 As Variant
Dim l02E4 As Variant
Dim l02E8 As Variant
Dim l02EC As Integer
Dim l02EE As Integer
Dim l02F0 As Integer
Dim l02F2 As Integer
Dim l02F4 As Integer
Dim l02F6 As Integer
Dim l02F8 As Integer
Dim l02FA As Integer
Dim l02FC As Variant
Dim l0300 As Variant
Dim l0304 As Integer
If  Combo1.ListCount = 0 Then MsgBox "Error: Who You Gonna TOS?", 16, "Error": Exit Sub
If  Text1.Text = "" Then MsgBox "Error: What Phrase You Gonna Use?", 16, "Error": Exit Sub
If  Label7.Caption = "S" Then GoTo L1EEFA:
If  Label7.Caption = "s" Then GoTo L1F33E:
Exit Sub
L1EEFA:
Call sub870("KO HELP")
l02DA% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l02DC% = extfn1A98(l02DA%, "I Need Help!")
Loop Until l02DC% <> 0
l02DE% = extfn1AD0(l02DC%, "_AOL_ICON")
L1EFE8:
Let l02E4 = Combo1.List(l02E0)
l02E0 = l02E0 + 1
l02E8 = Trim(l02E4)
If  l02E8 = "" Then GoTo L1F308
sub368 (l02DE%)
Do
DoEvents
l02EC% = extfn1A98(l02DA%, "Report Password Solicitations")
Loop Until l02EC% <> 0
l02EE% = extfn1A98(l02EC%, "Screen Name of Member Soliciting Your Information:")
l02F0% = extfn1B40(l02EE%, c00C8)
l02F2% = extfn19B8(l02F0%, c00CE, 0, l02E8)
l02F4% = extfn1A98(l02EC%, "Copy and Paste the solicitation here:")
l02F6% = extfn1B40(l02F4%, c00C8)
l02F8% = extfn19B8(l02F6%, c00CE, 0, Text1.Text)
l02FA% = extfn1A98(l02EC%, "Send")
sub368 (l02FA%)
subE20
L1F256:
Let l02FC = Combo2.List(l02E0)
l02E0 = l02E0 + 1
l0300 = Trim(l02FC)
If  l0300 = "" Then GoTo L1F2D4
If  UCase$(l0300) = UCase$(Combo1) Then Exit Sub
GoTo L1F256:
L1F2D4:
Combo2.AddItem Combo1
GoTo L1EFE8
L1F308:
l0304% = extfn1A28(l02DC%, c0106, 0, 0&)
     Exit Sub

L1F33E:
Command15.Value = True
subD08 .4#
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Now TOSing: *The Room* §]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Method: Room TOS - Phrase " + Label3 + "§]v^ñï")
Call sub870("KO HELP")
l02DA% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l02DC% = extfn1A98(l02DA%, "I Need Help!")
Loop Until l02DC% <> 0
l02DE% = extfn1AD0(l02DC%, "_AOL_ICON")
L1F57C:
Command15.Value = True
Let l02E4 = Combo1.List(l02E0)
l02E0 = l02E0 + 1
l02E8 = Trim(l02E4)
If  l02E8 = "" Then GoTo L1F982
sub368 (l02DE%)
Do
DoEvents
l02EC% = extfn1A98(l02DA%, "Report Password Solicitations")
Loop Until l02EC% <> 0
l02EE% = extfn1A98(l02EC%, "Screen Name of Member Soliciting Your Information:")
l02F0% = extfn1B40(l02EE%, c00C8)
l02F2% = extfn19B8(l02F0%, c00CE, 0, l02E8)
l02F4% = extfn1A98(l02EC%, "Copy and Paste the solicitation here:")
l02F6% = extfn1B40(l02F4%, c00C8)
l02F8% = extfn19B8(l02F6%, c00CE, 0, Text1.Text)
l02FA% = extfn1A98(l02EC%, "Send")
sub368 (l02FA%)
subE20
Call subBF0("ïñ^v[§" + l02E8 + ", You Have Been EXeCuTeD§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§" + Label5 + "§]v^ñï")
subD08 1E-100#
L1F8D0:
Let l02FC = Combo2.List(l02E0)
l02E0 = l02E0 + 1
l0300 = Trim(l02FC)
If  l0300 = "" Then GoTo L1F94E
If  UCase$(l0300) = UCase$(l02E8) Then Exit Sub
GoTo L1F8D0:
L1F94E:
Combo2.AddItem l02E8
GoTo L1F57C
L1F982:
l0304% = extfn1A28(l02DC%, c0106, 0, 0&)
subD08 1E-99#
Call subBF0("ïñ^v[§Room TOSing Has Been Completed§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§" + Label5 + "§]v^ñï")
Exit Sub
End Sub

Sub mnusave_Click ()
Dim l0306 As Integer
Dim l030A As Variant
If  Combo3.ListCount = 0 Then MsgBox "What you gonna save?", 16, "error": Exit Sub
On Error GoTo L1FC3C
CMDialog1.DialogTitle = "Save Executions Dead List"
CMDialog1.Filter = "List Files (*.Edl)|*.Edl| Execution Dead List (*.Edl)|*.Edl"
CMDialog1.Flags = &H1000&
CMDialog1.Action = 2
If  (CMDialog1.Filetitle <> "") Then
l0306 = 2
Open CMDialog1.Filename For Output As l0306
l030A = 0
Do While l030A < TOSer.Combo3.ListCount
Print #l0306, TOSer.Combo3.List(l030A)
l030A = l030A + 1
Loop
Close l0306
End If
TOSer.Combo3.Clear
L1FC3C:
Exit Sub
End Sub

Sub mnusavedre_Click ()
Dim l030E As Integer
Dim l0312 As Variant
If  Combo3.ListCount = 0 Then MsgBox "What you gonna save?", 16, "error": Exit Sub
On Error GoTo L1FDFC
CMDialog2.DialogTitle = "Save Executions Check List"
CMDialog2.Filter = "List Files (*.Ecl)|*.Ecl| Execution Check List (*.Ecl)|*.Ecl"
CMDialog2.Flags = &H1000&
CMDialog2.Action = 2
If  (CMDialog2.Filetitle <> "") Then
l030E = 2
Open CMDialog2.Filename For Output As l030E
l0312 = 0
Do While l0312 < Combo2.ListCount
Print #l030E, Combo2.List(l0312)
l0312 = l0312 + 1
Loop
Close l030E
End If
Combo2.Clear
L1FDFC:
Exit Sub
End Sub

Sub mnusupre_Click ()
Dim l0316 As Integer
Dim l0318 As Integer
Dim l031A As Integer
Dim l031C As Integer
Dim l031E As Integer
Dim l0320 As Integer
Dim l0322 As Integer
Dim l0324 As Integer
Dim l0326 As Integer
Dim l0328 As Integer
Dim l032A As Integer
Dim l032C As Integer
Dim l032E As Integer
Dim l0330 As Integer
Dim l0332 As Integer
Dim l0334 As Integer
Dim l0336 As Integer
Dim l0338 As Integer
Dim l033A As Integer
Dim l033C As Integer
Dim l033E As Variant
Dim l0342 As Variant
Dim l0346 As Variant
Dim l034A As Variant
If  Combo1.Text = "" Then MsgBox "Error: Who You Gonna TOS?", 16, "Error": Exit Sub
If  Text1.Text = "" Then MsgBox "Error: What Phrase You Gonna Use?", 16, "Error": Exit Sub
If  Label7.Caption = "S" Then GoTo L1FF12:
If  Label7.Caption = "s" Then GoTo L21234:
Exit Sub
L1FF12:
Call sub870("KO HELP")
l0316% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l0318% = extfn1A98(l0316%, "I Need Help!")
Loop Until l0318% <> 0
l031A% = extfn1AD0(l0318%, "_AOL_ICON")
sub368 (l031A%)
Do
DoEvents
l031C% = extfn1A98(l0316%, "Report Password Solicitations")
Loop Until l031C% <> 0
l031E% = extfn1A98(l031C%, "Screen Name of Member Soliciting Your Information:")
l0320% = extfn1B40(l031E%, c00C8)
l0322% = extfn19B8(l0320%, c00CE, 0, Combo1.Text)
l0324% = extfn1A98(l031C%, "Copy and Paste the solicitation here:")
l0326% = extfn1B40(l0324%, c00C8)
l0328% = extfn19B8(l0326%, c00CE, 0, Text1.Text)
l032A% = extfn1A98(l031C%, "Send")
sub368 (l032A%)
l032C% = extfn1A28(l0318%, c0106, 0, 0&)
subE20
subD08 .8#
Call sub870("GUIDEPAGER")
l0316% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l0318% = extfn1A98(l0316%, "I Need Help!")
Loop Until l0318% <> 0
l031A% = extfn1AD0(l0318%, "_AOL_ICON")
sub368 (l031A%)
Do
DoEvents
l031C% = extfn1A98(l0316%, "Report Password Solicitations")
Loop Until l031C% <> 0
l031E% = extfn1A98(l031C%, "Screen Name of Member Soliciting Your Information:")
l0320% = extfn1B40(l031E%, c00C8)
l0322% = extfn19B8(l0320%, c00CE, 0, Combo1.Text)
l0324% = extfn1A98(l031C%, "Copy and Paste the solicitation here:")
l0326% = extfn1B40(l0324%, c00C8)
l0328% = extfn19B8(l0326%, c00CE, 0, Text1.Text)
l032A% = extfn1A98(l031C%, "Send")
sub368 (l032A%)
l032C% = extfn1A28(l0318%, c0106, 0, 0&)
subE20
subD08 .8#
Call sub870("KO HELP")
l0316% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l0318% = extfn1A98(l0316%, "I Need Help!")
Loop Until l0318% <> 0
l031A% = extfn1AD0(l0318%, "_AOL_ICON")
l032E% = extfn1B40(l031A%, c00C8)
l0330% = extfn1B40(l032E%, c00C8)
sub368 (l0330%)
Do
DoEvents
l031C% = extfn1A98(l0316%, "Report A Violation")
Loop Until l031C% <> 0
SendKeys "{tab}" + "{ }"
Do
DoEvents
l0332% = extfn1A98(l0316%, "Violations via Instant Messages")
Loop Until l0332% <> 0
l031E% = extfn1A98(l0332%, "Date")
l0320% = extfn1B40(l031E%, c00C8)
l0322% = extfn19B8(l0320%, c00CE, 0, Date)
l0324% = extfn1A98(l0332%, "Time AM/PM")
l0326% = extfn1B40(l0324%, c00C8)
l0328% = extfn19B8(l0326%, c00CE, 0, Time)
l0334% = extfn1A98(l0332%, "CUT and PASTE a copy of the IM here")
l0336% = extfn1B40(l0334%, c00C8)
l0338% = extfn19B8(l0336%, c00CE, 0, Combo1.Text + ":  " + Text1.Text)
l032A% = extfn1A98(l0332%, "Send")
sub368 (l032A%)
l032C% = extfn1A28(l0318%, c0106, 0, 0&)
l032C% = extfn1A28(l031C%, c0106, 0, 0&)
subE20
subD08 .8#
Call sub870("staffpager")
l0316% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l0318% = extfn1A98(l0316%, "I Need Help!")
Loop Until l0318% <> 0
l031A% = extfn1AD0(l0318%, "_AOL_ICON")
sub368 (l031A%)
Do
DoEvents
l031C% = extfn1A98(l0316%, "Report Password Solicitations")
Loop Until l031C% <> 0
l031E% = extfn1A98(l031C%, "Screen Name of Member Soliciting Your Information:")
l0320% = extfn1B40(l031E%, c00C8)
l0322% = extfn19B8(l0320%, c00CE, 0, Combo1.Text)
l0324% = extfn1A98(l031C%, "Copy and Paste the solicitation here:")
l0326% = extfn1B40(l0324%, c00C8)
l0328% = extfn19B8(l0326%, c00CE, 0, Text1.Text)
l032A% = extfn1A98(l031C%, "Send")
sub368 (l032A%)
l032C% = extfn1A28(l0318%, c0106, 0, 0&)
subE20
subD08 .8#
l0316% = extfn1A60("AOL Frame25", 0&)
Call sub870("TOS")
Do
DoEvents
l033A% = extfn1A98(l0316%, "Terms of Service")
Loop Until l033A% <> 0
l033C% = extfn1AD0(l033A%, "_AOL_ICON")
SendKeys "{left}" + "{ }"
Do
DoEvents
l0318% = extfn1A98(l0316%, "I Need Help!")
Loop Until l0318% <> 0
l031A% = extfn1AD0(l0318%, "_AOL_ICON")
sub368 (l031A%)
Do
DoEvents
l031C% = extfn1A98(l0316%, "Report Password Solicitations")
Loop Until l031C% <> 0
l031E% = extfn1A98(l031C%, "Screen Name of Member Soliciting Your Information:")
l0320% = extfn1B40(l031E%, c00C8)
l0322% = extfn19B8(l0320%, c00CE, 0, Combo1.Text)
l0324% = extfn1A98(l031C%, "Copy and Paste the solicitation here:")
l0326% = extfn1B40(l0324%, c00C8)
l0328% = extfn19B8(l0326%, c00CE, 0, Text1.Text)
l032A% = extfn1A98(l031C%, "Send")
sub368 (l032A%)
l032C% = extfn1A28(l0318%, c0106, 0, 0&)
l032C% = extfn1A28(l033A%, c0106, 0, 0&)
subE20

L21180:
Let l0342 = Combo2.List(l033E)
l033E = l033E + 1
l0346 = Trim(l0342)
If  l0346 = "" Then GoTo L211FE
If  UCase$(l0346) = UCase$(Combo1) Then Exit Sub
GoTo L21180:
L211FE:
Combo2.AddItem Combo1
Exit Sub

L21234:
Command15.Value = True
subD08 .4#
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Now TOSing: " + Combo1 + "§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Method: Supreme TOS - Phrase " + Label3 + "§]v^ñï")
Call sub870("KO HELP")
l0316% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l0318% = extfn1A98(l0316%, "I Need Help!")
Loop Until l0318% <> 0
l031A% = extfn1AD0(l0318%, "_AOL_ICON")
sub368 (l031A%)
Do
DoEvents
l031C% = extfn1A98(l0316%, "Report Password Solicitations")
Loop Until l031C% <> 0
l031E% = extfn1A98(l031C%, "Screen Name of Member Soliciting Your Information:")
l0320% = extfn1B40(l031E%, c00C8)
l0322% = extfn19B8(l0320%, c00CE, 0, Combo1.Text)
l0324% = extfn1A98(l031C%, "Copy and Paste the solicitation here:")
l0326% = extfn1B40(l0324%, c00C8)
l0328% = extfn19B8(l0326%, c00CE, 0, Text1.Text)
l032A% = extfn1A98(l031C%, "Send")
sub368 (l032A%)
l032C% = extfn1A28(l0318%, c0106, 0, 0&)
subE20
Call subBF0("ïñ^v[§Kids Only TOS Completed§]v^ñï")
subD08 1E-99#
subD08 .8#
Call sub870("GUIDEPAGER")
l0316% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l0318% = extfn1A98(l0316%, "I Need Help!")
Loop Until l0318% <> 0
l031A% = extfn1AD0(l0318%, "_AOL_ICON")
sub368 (l031A%)
Do
DoEvents
l031C% = extfn1A98(l0316%, "Report Password Solicitations")
Loop Until l031C% <> 0
l031E% = extfn1A98(l031C%, "Screen Name of Member Soliciting Your Information:")
l0320% = extfn1B40(l031E%, c00C8)
l0322% = extfn19B8(l0320%, c00CE, 0, Combo1.Text)
l0324% = extfn1A98(l031C%, "Copy and Paste the solicitation here:")
l0326% = extfn1B40(l0324%, c00C8)
l0328% = extfn19B8(l0326%, c00CE, 0, Text1.Text)
l032A% = extfn1A98(l031C%, "Send")
sub368 (l032A%)
l032C% = extfn1A28(l0318%, c0106, 0, 0&)
subE20
Call subBF0("ïñ^v[§GuidePager TOS Completed§]v^ñï")
subD08 1E-99#
subD08 .8#
Call sub870("KO HELP")
l0316% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l0318% = extfn1A98(l0316%, "I Need Help!")
Loop Until l0318% <> 0
l031A% = extfn1AD0(l0318%, "_AOL_ICON")
l032E% = extfn1B40(l031A%, c00C8)
l0330% = extfn1B40(l032E%, c00C8)
sub368 (l0330%)
Do
DoEvents
l031C% = extfn1A98(l0316%, "Report A Violation")
Loop Until l031C% <> 0
SendKeys "{tab}" + "{ }"
Do
DoEvents
l0332% = extfn1A98(l0316%, "Violations via Instant Messages")
Loop Until l0332% <> 0
l031E% = extfn1A98(l0332%, "Date")
l0320% = extfn1B40(l031E%, c00C8)
l0322% = extfn19B8(l0320%, c00CE, 0, Date)
l0324% = extfn1A98(l0332%, "Time AM/PM")
l0326% = extfn1B40(l0324%, c00C8)
l0328% = extfn19B8(l0326%, c00CE, 0, Time)
l0334% = extfn1A98(l0332%, "CUT and PASTE a copy of the IM here")
l0336% = extfn1B40(l0334%, c00C8)
l0338% = extfn19B8(l0336%, c00CE, 0, Combo1.Text + ":  " + Text1.Text)
l032A% = extfn1A98(l0332%, "Send")
sub368 (l032A%)
l032C% = extfn1A28(l0318%, c0106, 0, 0&)
l032C% = extfn1A28(l031C%, c0106, 0, 0&)
subE20
Call subBF0("ïñ^v[§IM Violation [KW] TOS Completed§]v^ñï")
subD08 1E-99#
subD08 .8#
Call sub870("staffpager")
l0316% = extfn1A60("AOL Frame25", 0&)
Do
DoEvents
l0318% = extfn1A98(l0316%, "I Need Help!")
Loop Until l0318% <> 0
l031A% = extfn1AD0(l0318%, "_AOL_ICON")
sub368 (l031A%)
Do
DoEvents
l031C% = extfn1A98(l0316%, "Report Password Solicitations")
Loop Until l031C% <> 0
l031E% = extfn1A98(l031C%, "Screen Name of Member Soliciting Your Information:")
l0320% = extfn1B40(l031E%, c00C8)
l0322% = extfn19B8(l0320%, c00CE, 0, Combo1.Text)
l0324% = extfn1A98(l031C%, "Copy and Paste the solicitation here:")
l0326% = extfn1B40(l0324%, c00C8)
l0328% = extfn19B8(l0326%, c00CE, 0, Text1.Text)
l032A% = extfn1A98(l031C%, "Send")
sub368 (l032A%)
l032C% = extfn1A28(l0318%, c0106, 0, 0&)
subE20
Call subBF0("ïñ^v[§StaffPager TOS Completed§]v^ñï")
subD08 1E-99#
subD08 .8#
l0316% = extfn1A60("AOL Frame25", 0&)
Call sub870("TOS")
Do
DoEvents
l033A% = extfn1A98(l0316%, "Terms of Service")
Loop Until l033A% <> 0
l033C% = extfn1AD0(l033A%, "_AOL_ICON")
SendKeys "{left}" + "{ }"
Do
DoEvents
l0318% = extfn1A98(l0316%, "I Need Help!")
Loop Until l0318% <> 0
l031A% = extfn1AD0(l0318%, "_AOL_ICON")
sub368 (l031A%)
Do
DoEvents
l031C% = extfn1A98(l0316%, "Report Password Solicitations")
Loop Until l031C% <> 0
l031E% = extfn1A98(l031C%, "Screen Name of Member Soliciting Your Information:")
l0320% = extfn1B40(l031E%, c00C8)
l0322% = extfn19B8(l0320%, c00CE, 0, Combo1.Text)
l0324% = extfn1A98(l031C%, "Copy and Paste the solicitation here:")
l0326% = extfn1B40(l0324%, c00C8)
l0328% = extfn19B8(l0326%, c00CE, 0, Text1.Text)
l032A% = extfn1A98(l031C%, "Send")
sub368 (l032A%)
l032C% = extfn1A28(l0318%, c0106, 0, 0&)
l032C% = extfn1A28(l033A%, c0106, 0, 0&)
subE20
Call subBF0("ïñ^v[§Tech Live TOS Completed§]v^ñï")
subD08 1E-99#
subD08 1E-99#
Call subBF0("ïñ^v[§" + Combo1 + ", You Have Been EXeCuTeD§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§" + Label5 + "§]v^ñï")
L228CA:
Let l0342 = Combo2.List(l033E)
l033E = l033E + 1
l034A = Trim(l0342)
If  l034A = "" Then GoTo L22948
If  UCase$(l034A) = UCase$(Combo1) Then Exit Sub
GoTo L228CA:
L22948:
Combo2.AddItem Combo1
End Sub

Sub mnutech_Click ()
Dim l034E As Integer
Dim l0350 As Integer
Dim l0352 As Integer
Dim l0354 As Integer
Dim l0356 As Integer
Dim l0358 As Integer
Dim l035A As Integer
Dim l035C As Integer
Dim l035E As Integer
Dim l0360 As Integer
Dim l0362 As Integer
Dim l0364 As Integer
Dim l0366 As Integer
Dim l0368 As Integer
Dim l036A As Variant
Dim l036E As Variant
Dim l0372 As Variant
If  Combo1.Text = "" Then MsgBox "Error: Who You Gonna TOS?", 16, "error": Exit Sub
If  Text1.Text = "" Then MsgBox "Error: What Phrase You Gonna Use?", 16, "error": Exit Sub
If  Label7.Caption = "S" Then GoTo L22A7C
If  Label7.Caption = "s" Then GoTo L22F3A
Exit Sub
L22A7C:
l034E% = extfn1A60("AOL Frame25", 0&)
Call sub870("TOS")
Do
DoEvents
l0350% = extfn1A98(l034E%, "Terms of Service")
Loop Until l0350% <> 0
l0352% = extfn1AD0(l0350%, "_AOL_ICON")
SendKeys "{left}" + "{ }"
Do
DoEvents
l0354% = extfn1A98(l034E%, "I Need Help!")
Loop Until l0354% <> 0
l0356% = extfn1AD0(l0354%, "_AOL_ICON")
sub368 (l0356%)
Do
DoEvents
l0358% = extfn1A98(l034E%, "Report Password Solicitations")
Loop Until l0358% <> 0
l035A% = extfn1A98(l0358%, "Screen Name of Member Soliciting Your Information:")
l035C% = extfn1B40(l035A%, c00C8)
l035E% = extfn19B8(l035C%, c00CE, 0, Combo1.Text)
l0360% = extfn1A98(l0358%, "Copy and Paste the solicitation here:")
l0362% = extfn1B40(l0360%, c00C8)
l0364% = extfn19B8(l0362%, c00CE, 0, Text1.Text)
l0366% = extfn1A98(l0358%, "Send")
sub368 (l0366%)
l0368% = extfn1A28(l0354%, c0106, 0, 0&)
l0368% = extfn1A28(l0350%, c0106, 0, 0&)
subE20
L22E80:
Let l036E = Combo2.List(l036A)
l036A = l036A + 1
l0372 = Trim(l036E)
If  l0372 = "" Then GoTo L22EFE
If  UCase$(l0372) = UCase$(Combo1) Then Exit Sub
GoTo L22E80:
L22EFE:
Combo2.AddItem Combo1
Exit Sub
Exit Sub

L22F3A:
Command15.Value = True
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Now TOSing: " + Combo1 + "§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§Method: Tech Live - Phrase " + Label3 + "§]v^ñï")
l034E% = extfn1A60("AOL Frame25", 0&)
Call sub870("TOS")
Do
DoEvents
l0350% = extfn1A98(l034E%, "Terms of Service")
Loop Until l0350% <> 0
l0352% = extfn1AD0(l0350%, "_AOL_ICON")
SendKeys "{left}" + "{ }"
Do
DoEvents
l0354% = extfn1A98(l034E%, "I Need Help!")
Loop Until l0354% <> 0
l0356% = extfn1AD0(l0354%, "_AOL_ICON")
sub368 (l0356%)
Do
DoEvents
l0358% = extfn1A98(l034E%, "Report Password Solicitations")
Loop Until l0358% <> 0
l035A% = extfn1A98(l0358%, "Screen Name of Member Soliciting Your Information:")
l035C% = extfn1B40(l035A%, c00C8)
l035E% = extfn19B8(l035C%, c00CE, 0, Combo1.Text)
l0360% = extfn1A98(l0358%, "Copy and Paste the solicitation here:")
l0362% = extfn1B40(l0360%, c00C8)
l0364% = extfn19B8(l0362%, c00CE, 0, Text1.Text)
l0366% = extfn1A98(l0358%, "Send")
sub368 (l0366%)
l0368% = extfn1A28(l0354%, c0106, 0, 0&)
l0368% = extfn1A28(l0350%, c0106, 0, 0&)
subE20
Call subBF0("ïñ^v[§" + Combo1 + ", You Have Been EXeCuTeD§]v^ñï")
subD08 1E-99#
Call subBF0("ïñ^v[§" + Label5 + "§]v^ñï")

L23532:
Let l036E = Combo2.List(l036A)
l036A = l036A + 1
l0372 = Trim(l036E)
If  l0372 = "" Then GoTo L235B0
If  UCase$(l0372) = UCase$(Combo1) Then Exit Sub
GoTo L23532:
L235B0:
Combo2.AddItem Combo1
Exit Sub
End Sub

Sub mnutoslist_Click ()
Dim l0376 As String
Dim l0378 As Variant
l0376$ = InputBox("Please enter what you wanna add")
If  l0376$ = "" Then Exit Sub
For l0378 = 0 To Combo2.ListCount - 1
If  UCase$(Combo2.List(l0378)) Like UCase$("*" + l0376$ + "*") Then Exit Sub
Next
If  l0376$ = "" Then Exit Sub
Combo2.AddItem l0376$
End Sub

Sub mnutosrequ_Click ()
Dim l037C As String
l037C$ = InputBox("What is the SN To Request TOS", "TOS Request")
If  l037C$ = "" Then Exit Sub
Call subBF0("†")
subD08 1E-100#
Call subBF0("ïñ^v[§EXeCuTioNπ∑∫ TOSer ﬂy YuﬂaN§]v^ñï")
subD08 1E-100#
Call subBF0("ïñ^v[§TOS This ﬂastard ïñõ " + l037C$ + " §]v^ñï")
subD08 1E-100#
Call subBF0("†")
subD08 1E-100#
End Sub

Sub mnuup_Click ()
Call sub8A8
End Sub

Sub mnuyuhi_Click ()
Form102.Show
End Sub

Sub Timer1_Timer ()
Frame3D5.Caption = "TOSed:   " + Combo2.ListCount
Frame3D6.Caption = "Dead:   " + Combo3.ListCount
End Sub
