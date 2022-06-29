VERSION 2.00
Begin Form MMMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "FrostBite Mass Mailer v2.o By: DeCK"
   ClientHeight    =   2136
   ClientLeft      =   1596
   ClientTop       =   1752
   ClientWidth     =   6000
   ClipControls    =   0   'False
   FontBold        =   0   'False
   FontItalic      =   0   'False
   FontName        =   "Arial"
   FontSize        =   7.8
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   2472
   Icon            =   FEDMM.FRX:0000
   Left            =   1548
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2136
   ScaleWidth      =   6000
   Top             =   1464
   Width           =   6096
   Begin SSCommand Command3D12 
      Caption         =   "Room Buster"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   1320
      Width           =   1215
   End
   Begin SSCommand Command3D13 
      Caption         =   "Options"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   960
      Width           =   1215
   End
   Begin Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2280
      Top             =   120
   End
   Begin SSCommand Command3D3 
      Caption         =   "Exit MMer"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Width           =   1215
   End
   Begin SSCommand Command3D2 
      Caption         =   "MM Bot"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   1215
   End
   Begin SSCommand Command3D1 
      Caption         =   "Start MM"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   1215
   End
   Begin ListBox List5 
      Height          =   225
      Left            =   840
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   0
   End
   Begin CommonDialog CMDialog1 
      Left            =   480
      Top             =   -120
   End
   Begin ListBox List4 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   240
      Left            =   240
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin ListBox List3 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   240
      Left            =   2640
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin ListBox List2 
      Height          =   225
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin SSFrame Frame3D6 
      Alignment       =   2  'Center
      Caption         =   "Mail"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   1335
      Begin OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "New Mail"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Old Mail"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Flash Mail"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
   End
   Begin SSFrame Frame3D4 
      Alignment       =   2  'Center
      Caption         =   "Amount"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      Begin TextBox Text2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   990
         TabIndex        =   15
         Text            =   "300"
         Top             =   720
         Width           =   275
      End
      Begin OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MM Only"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "All Mail"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select Mails"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
   End
   Begin SSFrame Frame3D2 
      Alignment       =   2  'Center
      Caption         =   "Mass Mail Message"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
      Begin SSCommand Command3D6 
         Caption         =   "Load"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   600
         Width           =   855
      End
      Begin SSCommand Command3D5 
         Caption         =   "Save As"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   855
      End
      Begin SSCommand Command3D4 
         Caption         =   "Clear"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin TextBox Text1 
         BackColor       =   &H00000000&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "Mass Mail Message goes here."
         Top             =   240
         Width           =   2535
      End
   End
   Begin SSFrame Frame3D1 
      Alignment       =   2  'Center
      Caption         =   "People 0"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Arial"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin SSCommand Command3D11 
         Caption         =   "Save"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   1560
         Width           =   735
      End
      Begin SSCommand Command3D10 
         Caption         =   "Load"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   615
      End
      Begin SSCommand Command3D9 
         Caption         =   "Clear"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   1320
         Width           =   495
      End
      Begin SSCommand Command3D8 
         Caption         =   "Room"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1320
         Width           =   615
      End
      Begin SSCommand Command3D7 
         Caption         =   "+"
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   255
      End
      Begin ListBox List1 
         BackColor       =   &H00000000&
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "Arial"
         FontSize        =   7.8
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   1080
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
' FEDMM.FRM
Option Explicit
Dim m001A As Variant
Dim m002A As Variant
Dim m003A As String
Dim m003E As Variant
Const mc0066 = 16 ' &H10%
Const mc0072 =-1 ' &HFFFF%
Const mc0074 = 16 ' &H10%
Const mc0076 = 64 ' &H40%
Const mc00AA = 1025 ' &H401%
Const mc00CA = 1036 ' &H40C%
Const mc00D8 = 1031 ' &H407%
Const mc00E4 = 1034 ' &H40A%
Const mc015C = 12 ' &HC%
Const mc0160 = 2 ' &H2%
Const mc0166 = 14 ' &HE%
Const mc0168 = 13 ' &HD%

Sub Command3D1_Click ()
Dim l006A As Variant
Dim l0070 As Integer
' Const mc0072 =-1 ' &HFFFF%
' Const mc0074 = 16 ' &H10%
' Const mc0076 = 64 ' &H40%
Dim l0078 As Integer
Dim l007C As Integer
Dim l0080 As Integer
Dim l0084 As Integer
Dim l0086 As Integer
Dim l0088 As Integer
Dim l008A As Integer
Dim l008C As Integer
Dim l008E As Integer
Dim l0090 As Integer
Dim l0092 As Integer
Dim l0094 As Integer
Dim l0096 As Integer
Dim l0098 As Integer
Dim l009A As Integer
Dim l009E As Integer
Dim l00A0 As Integer
Dim l00A2 As Integer
Dim l00A4 As Integer
Dim l00A6 As Integer
Dim l00A8 As Integer
' Const mc00AA = 1025 ' &H401%
Dim l00AC As Integer
Dim l00B4 As Integer
Dim l00B6 As Integer
Dim l00B8 As Integer
Dim l00BA As Integer
Dim l00C8 As Integer
' Const mc00CA = 1036 ' &H40C%
Dim l00CC As Variant
Dim l00D0 As Variant
Dim l00D4 As Variant
' Const mc00D8 = 1031 ' &H407%
Dim l00DA As String
Dim l00DC As Variant
Dim l00E2 As Integer
' Const mc00E4 = 1034 ' &H40A%
Dim l00E8 As Integer
Dim l00F0 As Variant
Dim l00F4 As Integer
Dim l011C As Variant
Dim l0120 As Integer
Dim l0122 As String
Dim l0124 As Integer
Dim l0126 As Variant
Dim l012A As Variant
Dim l013A As Variant
Dim l013E As Variant
Dim l0142 As Integer
Dim l0144 As Variant
Dim l014A As Variant
Dim l014E As Variant
Dim l0152 As Variant
Dim l0158 As Integer
Dim l015A As String
' Const mc015C = 12 ' &HC%
Dim l015E As Integer
' Const mc0160 = 2 ' &H2%
Dim l0164 As Integer
' Const mc0166 = 14 ' &HE%
' Const mc0168 = 13 ' &HD%
Dim l016A As Variant
Dim l016E As Integer
Dim l0170 As String
Dim l0172 As Integer
Dim l0174 As Integer
Dim l0176 As Integer
Dim l0178 As Variant
Dim l017C As String
Dim l017E As String
Dim l0180 As Integer
Dim l0182 As Integer
Dim l0184 As Integer
Dim l0186 As Variant
Dim l018A As Integer
Dim l018C As String
Dim l0194 As Variant
If  List1.ListCount = 0 Then MsgBox "Nobody to massmail!", 16, "Error": Exit Sub
If  mmoptions.Check7.Value = 1 And mmoptions.Text1.Text = "" Then MsgBox "Please enter a prefix", 16, "Error: exit sub"
gv1210 = 0
MMstatus.Show
Me.Hide
MMstatus.Label10.Caption = "Status: Turning IM's Off"
If  gv1210 = 1 Then Exit Sub
Timer2.Enabled = False
Call sub0191("      -=FrostBite Mass Mailer v2.o By: DeCK=-")
If  mmoptions.Check11.Value = 1 Then
l006A = extfn0AB5(fn03A3(), mc0066, 0, 0)
End If
If  mmoptions.Check6.Value = 1 Then
l0070% = extfn0B02("mIRC", 0&)
extsub0E4F l0070%, mc0072, -1, -1, -1, -1, mc0074 Or mc0076
End If
MMstatus.Label10.Caption = "Status: Setting Preferences"
    Call sub045F("Preferences")
        m001A = 0
        Do
        DoEvents
            l0078% = extfn0B02("AOL Frame25", 0&)
            l007C% = extfn0B10(l0078%, "Preferences")
            l0080% = extfn0B24(l007C%, "_AOL_Icon")
        If  gv1210 = 1 Then Exit Sub
        Loop Until l007C% > 0
            l0084% = extfn0AA4(l0080%, 2)
            l0086% = extfn0AA4(l0084%, 2)
            l0088% = extfn0AA4(l0086%, 2)
            l008A% = extfn0AA4(l0088%, 2)
            l008C% = extfn0AA4(l008A%, 2)
            l008E% = extfn0AA4(l008C%, 2)
            l0090% = extfn0AA4(l008E%, 2)
            l0092% = extfn0AA4(l0090%, 2)
            l0094% = extfn0AA4(l0092%, 2)
            l0096% = extfn0AA4(l0094%, 2)
            l0098% = extfn0AA4(l0096%, 2)
            l009A% = extfn0AA4(l0098%, 2)
            If  fn0271() = 25 Then sub02B3 l0098%
            If  fn0271() = 30 Then sub02B3 l009A%
        Do
        DoEvents
            l009E% = extfn0B02("_AOL_Modal", "Mail Preferences")
        If  gv1210 = 1 Then Exit Sub
        Loop Until l009E% > 0
            l00A0% = extfn0B10(l009E%, "Confirm mail after it has been sent")
            l00A2% = extfn0B10(l009E%, "Close mail after it has been sent")
            l00A4% = extfn0B10(l009E%, "Retain all mail I send in my Personal Filing Cabinet")
            l00A6% = extfn0B10(l009E%, "Retain all mail I read in my Personal Filing Cabinet")
            l00A8% = extfn0B10(l009E%, "Use AOL style quoting in mail")
            l006A = extfn0AB5(l00A0%, mc00AA, 0, 0)
            l006A = extfn0AB5(l00A2%, mc00AA, 1, 0)
            l006A = extfn0AB5(l00A4%, mc00AA, 0, 0)
            l006A = extfn0AB5(l00A6%, mc00AA, 0, 0)
            l006A = extfn0AB5(l00A8%, mc00AA, 1, 0)
            l00AC% = extfn0B10(l009E%, "OK")
            sub02B3 l00AC%
        Do
        DoEvents
            l009E% = extfn0B02("_AOL_Modal", "Mail Preferences")
        If  gv1210 = 1 Then Exit Sub
        Loop Until l009E% = 0
            l0078% = extfn0B02("AOL Frame25", 0&)
            l007C% = extfn0B10(l0078%, "Preferences")
            l006A = extfn0AB5(l007C%, mc0066, 0, 0)
        Do
        DoEvents
            l0078% = extfn0B02("AOL Frame25", 0&)
            l007C% = extfn0B10(l0078%, "Preferences")
        If  gv1210 = 1 Then Exit Sub
        Loop Until l007C% = 0
MMstatus.Label10.Caption = "Status: Opening Mail"
If  Option1.Value = True Then
    l0078% = extfn0B02("AOL Frame25", 0&)
    l00B4% = extfn0B24(l0078%, "AOL Toolbar")
    l00B6% = extfn0B24(l00B4%, "_AOL_Icon")
    sub02B3 l00B6%
        Do
        DoEvents
            If  gv1210 = 1 Then Exit Sub
            l00B8% = extfn0B10(l0078%, "New Mail")
            l00BA% = extfn0B02("#32770", "America Online")
        Loop Until l00B8% > 0 Or l00BA% > 0
            If  l00BA% > 0 Then MsgBox "You have no new mail", 16, "Error": Exit Sub
ElseIf Option2.Value = True Then
    Call sub045F("Check Mail You've &Read")
        Do
        DoEvents
            If  gv1210 = 1 Then Exit Sub
            l00B8% = extfn0B10(l0078%, "Old Mail")
            l00BA% = extfn0B02("#32770", "America Online")
        Loop Until l00B8% > 0 Or l00BA% > 0
            If  l00BA% > 0 Then MsgBox "You have no old mail", 16, "Error": Exit Sub
ElseIf Option3.Value = True Then
    Call sub045F("Read &Incoming FlashMail")
        Do
        DoEvents
            If  gv1210 = 1 Then Exit Sub
            l00B8% = extfn0B10(l0078%, "Incoming FlashMail")
        Loop Until l00B8% > 0
End If
MMstatus.Label10.Caption = "Status: Waiting for mail to load..."
    l00C8% = extfn0B24(l00B8%, "_AOL_Tree")
    Do
    DoEvents
        If  gv1210 = 1 Then Exit Sub
        l00CC = extfn0AB5(l00C8%, mc00CA, 0, 0)
        sub060A 3
        l00D0 = extfn0AB5(l00C8%, mc00CA, 0, 0)
    Loop Until l00CC = l00D0
    If  mmoptions.Check3.Value = 1 Then
        For l00D4 = 0 To extfn0AB5(l00C8%, mc00CA, 0, 0) - 1
        l0078% = extfn0B02("AOL Frame25", 0&)
        l00C8% = extfn0B24(l0078%, "_AOL_Tree")
        l006A = extfn0AB5(l00C8%, mc00D8, l00D4, 0)
        l006A = extfn0AB5(l00C8%, mc00D8, l00D4, 0)
        l00DA$ = String(200, 0)
        l00C8% = extfn0B24(l0078%, "_AOL_Tree")
        l00E2% = extfn0AF3(l00B8%, l00DC, l00D4, l00DA$)
        l00E8% = extfn0AC9(l00C8%, mc00E4, l00D4, l00DA$)
        l00DA$ = Mid$(l00DA$, InStr(l00DA$, Chr$(9)) + 1)
        l00DA$ = Mid$(l00DA$, InStr(l00DA$, Chr$(9)) + 1)
        For l006A = 0 To List2.ListCount - 1
        l00F0 = InStr(l00DA$, List2.List(l006A))
        If  l00F0 Then
        l0078% = extfn0B02("AOL Frame25", 0&)
        l00F4% = extfn0B10(l0078%, "Delete")
        sub02B3 l00F4%
        GoTo LCCCC
        End If
        Next l006A
        List5.AddItem l00DA$
LCCCC:
        Next l00D4
    End If
    If  Option5.Value = True Then
            l00C8% = extfn0B24(l0078%, "_AOL_Tree")
                For l00D4 = 0 To extfn0AB5(l00C8%, mc00CA, 0, 0) - 1
                If  gv1210 = 1 Then Exit Sub
                l00DA$ = String(200, 0)
                l00C8% = extfn0B24(l0078%, "_AOL_Tree")
                l00E2% = extfn0AF3(l00B8%, l00DC, l00D4, l00DA$)
                l00E8% = extfn0AC9(l00C8%, mc00E4, l00D4, l00DA$)
                l00DA$ = Mid$(l00DA$, InStr(l00DA$, Chr$(9)) + 1)
                l00DA$ = Mid$(l00DA$, InStr(l00DA$, Chr$(9)) + 1)
                MMselect.List1.AddItem l00DA$
                Next l00D4
                MMselect.Show
               Do
               DoEvents
               Loop Until MMselect.Visible = False
        End If
If  Option5.Value = True Then
    l00CC = MMMain.List2.ListCount
End If
If  mmoptions.Check5.Value = 1 And Option5.Value = False Then
                For l00D4 = 0 To extfn0AB5(l00C8%, mc00CA, 0, 0) - 1
                l00DA$ = String(200, 0)
                l00C8% = extfn0B24(l0078%, "_AOL_Tree")
                l00E2% = extfn0AF3(l00B8%, l00DC, l00D4, l00DA$)
                l00E8% = extfn0AC9(l00C8%, mc00E4, l00D4, l00DA$)
                l00DA$ = Mid$(l00DA$, InStr(l00DA$, Chr$(9)) + 1)
                l00DA$ = Mid$(l00DA$, InStr(l00DA$, Chr$(9)) + 1)
                List3.AddItem l00DA$
                If  gv1210 = 1 Then Exit Sub
                Next l00D4
                sub060A .5#
                l00CC = MMMain.List3.ListCount
End If
If  Option6.Value = True Then
l00CC = Text2.Text
End If
MMstatus.Label7.Caption = l00CC
MMstatus.Label6.Caption = l00CC
For l011C = 0 To l00CC - 1
If  mmoptions.Check10.Value = 1 Then
l0120% = extfn0B02("_AOL_Palette", "America Online Timer")
l0080% = extfn0B24(l0120%, "_AOL_Button")
sub02B3 (l0080%)
End If
m002A = l00CC
If  gv1210 = 1 Then Exit Sub
l0122$ = ""
l0122$ = List3.List(l011C)
    l0078% = extfn0B02("AOL Frame25", 0&)
    If  Option1.Value = True Then l00B8% = extfn0B10(l0078%, "New Mail")
    If  Option2.Value = True Then l00B8% = extfn0B10(l0078%, "Old Mail")
    If  Option3.Value = True Then l00B8% = extfn0B10(l0078%, "Incoming FlashMail")
    l00C8% = extfn0B24(l00B8%, "_AOL_Tree")
        If  mmoptions.Check1.Value = 1 Then
            MMstatus.Label10.Caption = "Status: Keeping Mail as New"
            l0124% = extfn0B10(l00B8%, "Keep As New")
            sub02B3 l0124%
            sub060A .5#
            sub02B3 l0124%
        End If
    If  Option5.Value = False Then
        l006A = extfn0AB5(l00C8%, mc00D8, (l011C), (0))
        l006A = extfn0AB5(l00C8%, mc00D8, (l011C), (0))
    ElseIf Option5.Value = True Then
        l006A = extfn0AB5(l00C8%, mc00D8, List2.List(l011C), 0)
        l006A = extfn0AB5(l00C8%, mc00D8, List2.List(l011C), 0)
    End If
    If  mmoptions.Check5.Value = 1 Then
                If  Option1.Value = True Then l00B8% = extfn0B10(l0078%, "New Mail")
                If  Option2.Value = True Then l00B8% = extfn0B10(l0078%, "Old Mail")
                If  Option3.Value = True Then l00B8% = extfn0B10(l0078%, "Incoming FlashMail")
                l00C8% = extfn0B24(l00B8%, "_AOL_Tree")
                For l00D4 = 0 To extfn0AB5(l00C8%, mc00CA, 0, 0) - 1
                l00DA$ = String(200, 0)
                l00C8% = extfn0B24(l0078%, "_AOL_Tree")
                l00E2% = extfn0AF3(l00B8%, l00DC, l00D4, l00DA$)
                If  gv1210 = 1 Then Exit Sub
                l00E8% = extfn0AC9(l00C8%, mc00E4, l00D4, l00DA$)
                l00DA$ = Mid$(l00DA$, InStr(l00DA$, Chr$(9)) + 1)
                l00DA$ = Mid$(l00DA$, InStr(l00DA$, Chr$(9)) + 1)
                l0126 = InStr(l00DA$, List3.List(l011C))
                           If  l0126 Then
                                For l012A = 0 To List4.ListCount - 1
                                If  UCase$(l00D4) = UCase$(List4.List(l012A)) Then
                                GoTo LD894
                                End If
                                Next l012A
                                l006A = extfn0AB5(l00C8%, mc00D8, (l00D4), (0))
                                l006A = extfn0AB5(l00C8%, mc00D8, (l00D4), (0))
                                List4.AddItem l00D4
                                Exit For
                            End If
LD894:
                    Next l00D4
            End If
        If  mmoptions.Check9.Value = 1 Then
            If  Option4.Value = True Then
            l013A = l011C + 1
            l013E = l00CC - l013A
            ElseIf Option5.Value = True Then
            l013A = l011C + 1
            l013E = List2.List(List2.ListCount - l013A)
            ElseIf Option6.Value = True Then
            l00D0 = extfn0AB5(l00C8%, mc00CA, 0, 0)
            l013A = l011C + 1
            l013E = l00D0 - l013A
            End If
            l006A = extfn0AB5(l00C8%, mc00D8, (l013E), (0))
        l006A = extfn0AB5(l00C8%, mc00D8, (l013E), (0))
        End If
    l0142% = extfn0B10(l00B8%, "Read")
    If  Option3.Value = True Then
        l0142% = extfn0B10(l00B8%, "Open")
    End If
    MMstatus.Label10.Caption = "Status: Clicking Read"
    sub02B3 l0142%
    Do
    DoEvents
        MMstatus.Label10.Caption = "Status: Waiting for mail to open..."
        l0078% = extfn0B02("AOL Frame25", 0&)
        l0144 = extfn0B10(l0078%, "Reply")
        l014A = extfn0B5D(l0144)
        l014E = extfn0B10(l0078%, "Forward")
        l014A = extfn0B5D(l014E)
        If  gv1210 = 1 Then Exit Sub
        l0152 = extfn0B10(l0078%, "Reply to All")
        l014A = extfn0B5D(l0152)
        Call sub045F("Stop Incoming Text")
    Loop Until l014A > 0
    Call sub0477
    l0080% = extfn0B24(l014A, "_AOL_Icon")
    l0084% = extfn0AA4(l0080%, 2)
    l0086% = extfn0AA4(l0084%, 2)
    sub02B3 l0086%
    MMstatus.Label10.Caption = "Status: Clicking Forward"
    Do
    DoEvents
        If  gv1210 = 1 Then Exit Sub
        sub060A .15#
    Loop Until fn03B3() > 0

LDD7A:
MMstatus.Label10.Caption = "Status: Sending Mail"
    l0158% = extfn0B24(fn03B3(), "_AOL_Edit")
        l015A$ = ""
        If  mmoptions.Check8.Value = 1 Then
            For l00D4 = 0 To List1.ListCount - 1
                l015A$ = l015A$ + "(" + List1.List(l00D4) + "), "
            Next l00D4
        ElseIf mmoptions.Check8.Value <> 1 Then
            For l00D4 = 0 To List1.ListCount - 1
                l015A$ = l015A$ + List1.List(l00D4) + ", "
            Next l00D4
        End If
    l006A = extfn0AC9(l0158%, mc015C, 0, l015A$)
    l015E% = extfn0B10(fn03B3(), "Subject:")
    l0164% = extfn0B50(l015E%, mc0160)
    l0126 = extfn0AB5(l0164%, mc0166, 0, 0)
    l00DA$ = Space(l0126 + 1)
    l016A = extfn0AC9(l0164%, mc0168, l0126 + 1, l00DA$)
    l016E% = InStr(l00DA$, "Fwd: ")
    l0170$ = Mid$(l00DA$, 6, (Len(l00DA$) - l016E%))
        If  mmoptions.Check2.Value = 1 Then
            l016E% = InStr(l00DA$, "Fwd: ")
            If  l016E% Then
            l006A = extfn0AC9(l0164%, mc015C, 0, "")
            l006A = extfn0AC9(l0164%, mc015C, 0, l0170$)
            End If
        End If
        If  mmoptions.Check7.Value = 1 Then
            l015E% = extfn0B10(fn03B3(), "Subject:")
            l0164% = extfn0B50(l015E%, mc0160)
            l0126 = extfn0AB5(l0164%, mc0166, 0, 0)
            l00DA$ = Space(l0126 + 1)
            l016A = extfn0AC9(l0164%, mc0168, l0126 + 1, l00DA$)
            l006A = extfn0AC9(l0164%, mc015C, 0, "")
            l006A = extfn0AC9(l0164%, mc015C, 0, mmoptions.Text1.Text + " " + l00DA$)
        End If
l0172% = extfn0B24(fn03B3(), "_AOL_Edit")
l0172% = extfn0B50(l0172%, mc0160)
l0172% = extfn0B50(l0172%, mc0160)
l0172% = extfn0B50(l0172%, mc0160)
l0172% = extfn0B50(l0172%, mc0160)
l0174% = extfn0B24(fn03B3(), "RICHCNTL")
If  l0174% <> 0 Then
    l0176% = extfn0B24(fn03B3(), "RICHCNTL")
End If
If  l0174% = 0 Then
    l015E% = extfn0B10(fn03B3(), "Subject:")
    l0164% = extfn0B50(l015E%, mc0160)
    l0176% = extfn0B50(l0164%, mc0160)
    End If
    l013A = l011C + 1
    l014A = (l013A / l00CC) * 100
    l0178 = Int(l014A)
    l006A = extfn0AC9(l0176%, mc015C, 0, "")
    l017C$ = Chr(13) + Chr(10)
    l017E$ = ""
    l017E$ = l017E$ + l017C$
    l017E$ = l017E$ + "×–•[·FrostBite·]•–×   FrostBite Mass Mailer v2.o By DeCK" + l017C$
    l017E$ = l017E$ + "×–•[·FrostBite·]•–×   There Are " & MMMain.List1.ListCount & " People On This MM" + l017C$
    l017E$ = l017E$ + "×–•[·FrostBite·]•–×   This Mass Mail is " & l0178 & "% Complete" + l017C$
    l017E$ = l017E$ + "×–•[·FrostBite·]•–×   This Is Mail " & l013A & " Out Of " & l00CC & "" + l017C$
    l017E$ = l017E$ + l017C$
    l017E$ = l017E$ + MMMain.Text1.Text
    l006A = extfn0AC9(l0176%, mc015C, 0, l017E$)
    l0180% = extfn0B24(fn03B3(), "_AOL_Icon")
    sub02B3 l0180%
        Do
        DoEvents
            l0182% = extfn0B10(l0078%, "Error")
            If  gv1210 = 1 Then Exit Sub
        Loop Until fn03B3() = 0 Or l0182% > 0
            If  l0182% > 0 Then
                MMstatus.Label10.Caption = "Status: Removing Dead Boxes..."
                Do
                DoEvents
                l0184% = extfn0B24(l0182%, "_AOL_View")
                l0186 = extfn0AB5(l0184%, mc0166, 0, 0)
                l00DA$ = Space$(l0186)
                l018A% = extfn0AC9(l0184%, mc0168, l0186 + 1, l00DA$)
                l00DA$ = Trim$(l00DA$)
                sub060A 1.5#
                l0184% = extfn0B24(l0182%, "_AOL_View")
                l0186 = extfn0AB5(l0184%, mc0166, 0, 0)
                l018C$ = Space$(l0186)
                l018A% = extfn0AC9(l0184%, mc0168, l0186 + 1, l018C$)
                l018C$ = Trim$(l018C$)
                Loop Until l00DA$ = l018C$
LE900:
                    For l00D4 = 0 To MMMain.List1.ListCount - 1
                        l014A = InStr(l00DA$, List1.List(l00D4))
                            If  l014A Then
                            MMstatus.List1.AddItem List1.List(l00D4)
                            List1.RemoveItem l00D4
                            GoTo LE900
                            End If
                    Next l00D4
                    Frame3D1.Caption = "People " & List1.ListCount
                l006A = extfn0AB5(l0182%, mc0066, 0, 0)
                Do
                DoEvents
                l0078% = extfn0B02("AOL Frane25", 0&)
                l0182% = extfn0B10(l0078%, "Error")
                Loop Until l0182% = 0
                If  List1.ListCount = 0 Then MsgBox "No more people to MM", 16, "Error": Exit Sub
                GoTo LDD7A:
            End If
    l0078% = extfn0B02("AOL Frame25", 0&)
    l0144 = extfn0B10(l0078%, "Reply")
    l0194 = extfn0B5D(l0144)
    l014E = extfn0B10(l0078%, "Forward")
    l0194 = extfn0B5D(l014E)
    l0152 = extfn0B10(l0078%, "Reply to All")
    l0194 = extfn0B5D(l0152)
    l006A = extfn0AB5(l0194, mc0066, 0, 0)
        Do
        DoEvents
        l0152 = extfn0B10(l0078%, "Reply to All")
        l0194 = extfn0B5D(l0152)
        If  gv1210 = 1 Then Exit Sub
        Loop Until l0194 = 0
        MMstatus.Label9.Visible = True
        MMstatus.Label6.Caption = MMstatus.Label6.Caption - 1
        MMstatus.Label5.Caption = MMstatus.Label5.Caption + 1
        MMstatus.Label10.Caption = "Status: Going to next mail..."
        l014A = (l013A / l00CC) * 4215
        l0178 = Int(l014A)
        MMstatus.Label9.Width = l0178
        l014A = (l013A / l00CC) * 100
        l0178 = Int(l014A)
        MMstatus.Label11.Caption = l0178 & "%"
        If  gv1210 = 1 Then Exit Sub
Next l011C
    l0078% = extfn0B02("AOL Frame25", 0&)
    If  Option1.Value = True Then l00B8% = extfn0B10(l0078%, "New Mail")
    If  Option2.Value = True Then l00B8% = extfn0B10(l0078%, "Old Mail")
    If  Option3.Value = True Then l00B8% = extfn0B10(l0078%, "Incoming FlashMail")
    l00C8% = extfn0B24(l00B8%, "_AOL_Tree")
        If  mmoptions.Check1.Value = 1 Then
            MMstatus.Label10.Caption = "Status: Keeping Mail as New"
            l0124% = extfn0B10(l00B8%, "Keep As New")
            sub02B3 l0124%
            sub060A .5#
            sub02B3 l0124%
        End If
If  mmoptions.Check4.Value = 1 Then
    MMstatus.Label10.Caption = "Status: Signing Off"
    Call sub045F("E&xit")
End If
m001A = 1
m003E = MMMain.List1.ListCount
m003A$ = ""
For l00D4 = 0 To MMMain.List1.ListCount - 1
m003A$ = m003A$ + MMMain.List1.List(l00D4) + Chr(13) + Chr(10)
Next l00D4
Unload MMstatus
Me.Show
End Sub

Sub Command3D10_Click ()
loadd.Show
End Sub

Sub Command3D11_Click ()
save.Show
End Sub

Sub Command3D12_Click ()
bust.Show
Me.Hide
End Sub

Sub Command3D13_Click ()
mmoptions.Show
End Sub

Sub Command3D2_Click ()
Dim l01A8 As Variant
bot.List1.Clear
For l01A8 = 0 To List1.ListCount - 1
bot.List1.AddItem List1.List(l01A8)
Next l01A8
bot.Label4.Caption = "People: " + bot.List1.ListCount
bot.Show
Me.Hide
End Sub

Sub Command3D3_Click ()
End
End Sub

Sub Command3D4_Click ()
Text1 = ""
End Sub

Sub Command3D5_Click ()
CMDialog1.DialogTitle = "Save MM Message"
 CMDialog1.Filter = "Tag Files (*.tag)|*.tag|"
 CMDialog1.FilterIndex = 2
CMDialog1.Action = 2
If  CMDialog1.Filetitle = "" Then Exit Sub
Open CMDialog1.Filetitle For Output As #1
Print #1, Text1.Text
Close #1
End Sub

Sub Command3D6_Click ()
Dim l01B8 As Integer
Dim l01BA As String
On Error Resume Next
 CMDialog1.Filter = "Tag Files (*.tag)|*.tag|"
 CMDialog1.FilterIndex = 2
 CMDialog1.Action = 1
 CMDialog1.DialogTitle = "Load MM Message"
 If  CMDialog1.Filetitle = "" Then Exit Sub
    l01B8% = FreeFile
    Open CMDialog1.Filetitle For Input As l01B8%
    Text1 = ""
    Do: DoEvents
    Line Input #1, l01BA$
    If  l01BA$ <> "" Then Text1 = Text1 & l01BA$ + Chr(13) + Chr(10)
    Loop Until EOF(l01B8%)
    Close l01B8%
End Sub

Sub Command3D7_Click ()
Dim l01BC As Variant
Dim l01C0 As Variant
l01BC = InputBox("Enter the name to add:", "Add")
If  Len(l01BC) > 10 Or Len(l01BC) < 3 Then MsgBox "Invalid Username": Exit Sub
For l01C0 = 0 To List1.ListCount - 1
If  List1.List(l01C0) = l01BC Then MsgBox "Name is already added to list": Exit Sub
Next l01C0
List1.AddItem l01BC
Frame3D1.Caption = "People " + List1.ListCount
End Sub

Sub Command3D8_Click ()
Dim l01C4 As Integer
l01C4% = fn03A3()
Call sub01DD(l01C4%, List1)
Frame3D1.Caption = "People " + List1.ListCount
End Sub

Sub Command3D9_Click ()
List1.Clear
Frame3D1.Caption = "People " + List1.ListCount
End Sub

Sub Form_Load ()
sub05BC Me
sub02A9 Me
List2.Clear
End Sub

Sub Form_Unload (Cancel As Integer)
End
End Sub

Sub List1_DblClick ()
List1.RemoveItem List1.ListIndex
Frame3D1.Caption = "People " & List1.ListCount
End Sub

Sub Option1_Click ()
mmoptions.Check1.Enabled = True
mmoptions.Check3.Enabled = True
End Sub

Sub Option2_Click ()
mmoptions.Check1.Enabled = True
mmoptions.Check3.Enabled = True
End Sub

Sub Option3_Click ()
mmoptions.Check1.Value = 0
mmoptions.Check1.Enabled = False
mmoptions.Check3.Value = 0
mmoptions.Check3.Enabled = False
End Sub

Sub Timer1_Timer ()
Command3D1_Click
Timer1.Enabled = False
End Sub

Sub Timer2_Timer ()
Dim l01CE As Integer
Dim l01D0 As Integer
l01CE% = extfn0B02("_AOL_Palette", "America Online Timer")
l01D0% = extfn0B24(l01CE%, "_AOL_Button")
sub02B3 (l01D0%)
End Sub
