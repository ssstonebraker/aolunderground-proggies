VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form board 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Vegas Video Keno"
   ClientHeight    =   7185
   ClientLeft      =   105
   ClientTop       =   -180
   ClientWidth     =   9615
   Icon            =   "board.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7185
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar autobar 
      Height          =   255
      Left            =   720
      TabIndex        =   10
      ToolTipText     =   "AutoPlay Progress Meter"
      Top             =   240
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.CommandButton change_button 
      BackColor       =   &H00000000&
      Height          =   450
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Keno Options"
      Top             =   6615
      Width           =   1080
   End
   Begin VB.CommandButton play4_button 
      Enabled         =   0   'False
      Height          =   480
      Left            =   4815
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Play 4 Credits"
      Top             =   6570
      Width           =   1185
   End
   Begin VB.CommandButton betone 
      Height          =   435
      Left            =   6195
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Bet One Credit"
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton erase_button 
      Enabled         =   0   'False
      Height          =   480
      Left            =   2625
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Erase All Marks"
      Top             =   6555
      Width           =   1185
   End
   Begin VB.CommandButton start_button 
      Enabled         =   0   'False
      Height          =   495
      Left            =   7425
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deal Numbers"
      Top             =   6555
      UseMaskColor    =   -1  'True
      Width           =   1185
   End
   Begin VB.Label autoplaylabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AutoPlay Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3960
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label paid_label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   4350
      TabIndex        =   8
      Top             =   825
      Width           =   810
   End
   Begin VB.Image top_display 
      Height          =   285
      Left            =   3840
      Top             =   630
      Width           =   1935
   End
   Begin VB.Image bottom_display 
      Height          =   570
      Left            =   3165
      Top             =   5805
      Width           =   3120
   End
   Begin VB.Label coinsin_label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1965
      TabIndex        =   7
      ToolTipText     =   "Amount Bet"
      Top             =   6060
      Width           =   330
   End
   Begin VB.Label credits_label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   6870
      TabIndex        =   6
      ToolTipText     =   "Amount Of Credits Left"
      Top             =   6075
      Width           =   1275
   End
   Begin VB.Label hits_label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   7590
      TabIndex        =   5
      ToolTipText     =   "Spots Hit"
      Top             =   645
      Width           =   435
   End
   Begin VB.Label spots_marked 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Tag             =   "Spots Marked"
      Top             =   675
      Width           =   585
   End
   Begin VB.Image c80 
      Height          =   435
      Left            =   7560
      Top             =   5250
      Width           =   630
   End
   Begin VB.Image c78 
      Height          =   435
      Left            =   6180
      Top             =   5250
      Width           =   630
   End
   Begin VB.Image c74 
      Height          =   405
      Left            =   3420
      Top             =   5265
      Width           =   630
   End
   Begin VB.Image c79 
      Height          =   435
      Left            =   6885
      Top             =   5235
      Width           =   645
   End
   Begin VB.Image c75 
      Height          =   405
      Left            =   4095
      Top             =   5265
      Width           =   630
   End
   Begin VB.Image c73 
      Height          =   435
      Left            =   2700
      Top             =   5235
      Width           =   630
   End
   Begin VB.Image c77 
      Height          =   435
      Left            =   5505
      Top             =   5250
      Width           =   600
   End
   Begin VB.Image c76 
      Height          =   435
      Left            =   4800
      Top             =   5265
      Width           =   645
   End
   Begin VB.Image c72 
      Height          =   435
      Left            =   2025
      Top             =   5250
      Width           =   645
   End
   Begin VB.Image c71 
      Height          =   435
      Left            =   1350
      Top             =   5250
      Width           =   615
   End
   Begin VB.Image c70 
      Height          =   435
      Left            =   7530
      Top             =   4710
      Width           =   675
   End
   Begin VB.Image c68 
      Height          =   435
      Left            =   6165
      Top             =   4710
      Width           =   615
   End
   Begin VB.Image c64 
      Height          =   435
      Left            =   3405
      Top             =   4710
      Width           =   645
   End
   Begin VB.Image c69 
      Height          =   435
      Left            =   6840
      Top             =   4710
      Width           =   660
   End
   Begin VB.Image c65 
      Height          =   435
      Left            =   4080
      Top             =   4710
      Width           =   660
   End
   Begin VB.Image c63 
      Height          =   435
      Left            =   2715
      Top             =   4710
      Width           =   645
   End
   Begin VB.Image c67 
      Height          =   435
      Left            =   5475
      Top             =   4710
      Width           =   645
   End
   Begin VB.Image c66 
      Height          =   435
      Left            =   4770
      Top             =   4710
      Width           =   660
   End
   Begin VB.Image c62 
      Height          =   435
      Left            =   2010
      Top             =   4710
      Width           =   645
   End
   Begin VB.Image c61 
      Height          =   435
      Left            =   1350
      Top             =   4710
      Width           =   630
   End
   Begin VB.Image c60 
      Height          =   435
      Left            =   7545
      Top             =   4155
      Width           =   675
   End
   Begin VB.Image c58 
      Height          =   435
      Left            =   6180
      Top             =   4155
      Width           =   615
   End
   Begin VB.Image c54 
      Height          =   435
      Left            =   3435
      Top             =   4140
      Width           =   615
   End
   Begin VB.Image c59 
      Height          =   435
      Left            =   6855
      Top             =   4155
      Width           =   630
   End
   Begin VB.Image c55 
      Height          =   450
      Left            =   4095
      Top             =   4125
      Width           =   660
   End
   Begin VB.Image c53 
      Height          =   435
      Left            =   2730
      Top             =   4140
      Width           =   630
   End
   Begin VB.Image c57 
      Height          =   435
      Left            =   5475
      Top             =   4155
      Width           =   615
   End
   Begin VB.Image c56 
      Height          =   435
      Left            =   4785
      Top             =   4140
      Width           =   615
   End
   Begin VB.Image c52 
      Height          =   435
      Left            =   2025
      Top             =   4155
      Width           =   645
   End
   Begin VB.Image c51 
      Height          =   435
      Left            =   1365
      Top             =   4140
      Width           =   600
   End
   Begin VB.Image c50 
      Height          =   435
      Left            =   7560
      Top             =   3600
      Width           =   675
   End
   Begin VB.Image c48 
      Height          =   435
      Left            =   6150
      Top             =   3600
      Width           =   660
   End
   Begin VB.Image c44 
      Height          =   450
      Left            =   3420
      Top             =   3600
      Width           =   600
   End
   Begin VB.Image c49 
      Height          =   435
      Left            =   6855
      Top             =   3600
      Width           =   645
   End
   Begin VB.Image c45 
      Height          =   450
      Left            =   4095
      Top             =   3600
      Width           =   645
   End
   Begin VB.Image c43 
      Height          =   435
      Left            =   2715
      Top             =   3585
      Width           =   630
   End
   Begin VB.Image c47 
      Height          =   435
      Left            =   5490
      Top             =   3600
      Width           =   615
   End
   Begin VB.Image c46 
      Height          =   450
      Left            =   4785
      Top             =   3585
      Width           =   615
   End
   Begin VB.Image c42 
      Height          =   435
      Left            =   2055
      Top             =   3600
      Width           =   600
   End
   Begin VB.Image c41 
      Height          =   435
      Left            =   1350
      Top             =   3585
      Width           =   615
   End
   Begin VB.Image c40 
      Height          =   435
      Left            =   7545
      Top             =   2895
      Width           =   675
   End
   Begin VB.Image c38 
      Height          =   435
      Left            =   6180
      Top             =   2895
      Width           =   615
   End
   Begin VB.Image c34 
      Height          =   435
      Left            =   3450
      Top             =   2895
      Width           =   600
   End
   Begin VB.Image c39 
      Height          =   435
      Left            =   6840
      Top             =   2895
      Width           =   630
   End
   Begin VB.Image c35 
      Height          =   435
      Left            =   4125
      Top             =   2895
      Width           =   615
   End
   Begin VB.Image c33 
      Height          =   435
      Left            =   2745
      Top             =   2895
      Width           =   630
   End
   Begin VB.Image c37 
      Height          =   435
      Left            =   5475
      Top             =   2880
      Width           =   645
   End
   Begin VB.Image c36 
      Height          =   435
      Left            =   4800
      Top             =   2895
      Width           =   615
   End
   Begin VB.Image c32 
      Height          =   435
      Left            =   2055
      Top             =   2895
      Width           =   600
   End
   Begin VB.Image c31 
      Height          =   435
      Left            =   1380
      Top             =   2895
      Width           =   600
   End
   Begin VB.Image c30 
      Height          =   435
      Left            =   7530
      Top             =   2370
      Width           =   675
   End
   Begin VB.Image c28 
      Height          =   435
      Left            =   6165
      Top             =   2370
      Width           =   615
   End
   Begin VB.Image c24 
      Height          =   435
      Left            =   3435
      Top             =   2370
      Width           =   600
   End
   Begin VB.Image c29 
      Height          =   435
      Left            =   6840
      Top             =   2355
      Width           =   630
   End
   Begin VB.Image c25 
      Height          =   435
      Left            =   4110
      Top             =   2370
      Width           =   615
   End
   Begin VB.Image c23 
      Height          =   435
      Left            =   2745
      Top             =   2370
      Width           =   630
   End
   Begin VB.Image c27 
      Height          =   435
      Left            =   5475
      Top             =   2370
      Width           =   615
   End
   Begin VB.Image c26 
      Height          =   435
      Left            =   4815
      Top             =   2370
      Width           =   615
   End
   Begin VB.Image c22 
      Height          =   435
      Left            =   2055
      Top             =   2370
      Width           =   600
   End
   Begin VB.Image c21 
      Height          =   435
      Left            =   1380
      Top             =   2370
      Width           =   600
   End
   Begin VB.Image c20 
      Height          =   435
      Left            =   7530
      Top             =   1830
      Width           =   675
   End
   Begin VB.Image c18 
      Height          =   435
      Left            =   6165
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image c14 
      Height          =   435
      Left            =   3420
      Top             =   1830
      Width           =   600
   End
   Begin VB.Image c19 
      Height          =   435
      Left            =   6840
      Top             =   1815
      Width           =   630
   End
   Begin VB.Image c15 
      Height          =   435
      Left            =   4110
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image c13 
      Height          =   435
      Left            =   2745
      Top             =   1830
      Width           =   630
   End
   Begin VB.Image c17 
      Height          =   435
      Left            =   5475
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image c16 
      Height          =   435
      Left            =   4800
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image c12 
      Height          =   435
      Left            =   2055
      Top             =   1830
      Width           =   600
   End
   Begin VB.Image c11 
      Height          =   435
      Left            =   1380
      Top             =   1830
      Width           =   600
   End
   Begin VB.Image c10 
      Height          =   435
      Left            =   7530
      Top             =   1305
      Width           =   675
   End
   Begin VB.Image c8 
      Height          =   435
      Left            =   6165
      Top             =   1305
      Width           =   615
   End
   Begin VB.Image c4 
      Height          =   435
      Left            =   3435
      Top             =   1305
      Width           =   600
   End
   Begin VB.Image c9 
      Height          =   435
      Left            =   6825
      Top             =   1290
      Width           =   630
   End
   Begin VB.Image c5 
      Height          =   435
      Left            =   4110
      Top             =   1305
      Width           =   615
   End
   Begin VB.Image c3 
      Height          =   435
      Left            =   2745
      Top             =   1305
      Width           =   630
   End
   Begin VB.Image c7 
      Height          =   435
      Left            =   5475
      Top             =   1305
      Width           =   615
   End
   Begin VB.Image c6 
      Height          =   435
      Left            =   4800
      Top             =   1305
      Width           =   615
   End
   Begin VB.Image c2 
      Height          =   435
      Left            =   2055
      Top             =   1290
      Width           =   600
   End
   Begin VB.Image c1 
      Height          =   435
      Left            =   1380
      Top             =   1305
      Width           =   600
   End
   Begin VB.Image h80 
      Height          =   435
      Left            =   7575
      Top             =   5250
      Width           =   630
   End
   Begin VB.Image h78 
      Height          =   435
      Left            =   6195
      Top             =   5250
      Width           =   585
   End
   Begin VB.Image h74 
      Height          =   405
      Left            =   3435
      Top             =   5265
      Width           =   630
   End
   Begin VB.Image h79 
      Height          =   435
      Left            =   6855
      Top             =   5235
      Width           =   645
   End
   Begin VB.Image h75 
      Height          =   405
      Left            =   4110
      Top             =   5265
      Width           =   630
   End
   Begin VB.Image h73 
      Height          =   435
      Left            =   2730
      Top             =   5235
      Width           =   630
   End
   Begin VB.Image h77 
      Height          =   435
      Left            =   5520
      Top             =   5250
      Width           =   600
   End
   Begin VB.Image h76 
      Height          =   435
      Left            =   4815
      Top             =   5265
      Width           =   645
   End
   Begin VB.Image h72 
      Height          =   435
      Left            =   2040
      Top             =   5250
      Width           =   645
   End
   Begin VB.Image h71 
      Height          =   435
      Left            =   1365
      Top             =   5250
      Width           =   615
   End
   Begin VB.Image h70 
      Height          =   435
      Left            =   7545
      Top             =   4710
      Width           =   675
   End
   Begin VB.Image h68 
      Height          =   435
      Left            =   6180
      Top             =   4710
      Width           =   615
   End
   Begin VB.Image h64 
      Height          =   435
      Left            =   3420
      Top             =   4710
      Width           =   645
   End
   Begin VB.Image h69 
      Height          =   435
      Left            =   6855
      Top             =   4710
      Width           =   660
   End
   Begin VB.Image h65 
      Height          =   435
      Left            =   4095
      Top             =   4710
      Width           =   660
   End
   Begin VB.Image h63 
      Height          =   435
      Left            =   2730
      Top             =   4710
      Width           =   645
   End
   Begin VB.Image h67 
      Height          =   435
      Left            =   5490
      Top             =   4710
      Width           =   645
   End
   Begin VB.Image h66 
      Height          =   435
      Left            =   4785
      Top             =   4710
      Width           =   660
   End
   Begin VB.Image h62 
      Height          =   435
      Left            =   2040
      Top             =   4710
      Width           =   645
   End
   Begin VB.Image h61 
      Height          =   435
      Left            =   1380
      Top             =   4710
      Width           =   630
   End
   Begin VB.Image h60 
      Height          =   435
      Left            =   7560
      Top             =   4155
      Width           =   675
   End
   Begin VB.Image h58 
      Height          =   435
      Left            =   6195
      Top             =   4155
      Width           =   615
   End
   Begin VB.Image h54 
      Height          =   435
      Left            =   3450
      Top             =   4140
      Width           =   615
   End
   Begin VB.Image h59 
      Height          =   435
      Left            =   6870
      Top             =   4155
      Width           =   630
   End
   Begin VB.Image h55 
      Height          =   435
      Left            =   4110
      Top             =   4125
      Width           =   615
   End
   Begin VB.Image h53 
      Height          =   435
      Left            =   2745
      Top             =   4140
      Width           =   630
   End
   Begin VB.Image h57 
      Height          =   435
      Left            =   5490
      Top             =   4155
      Width           =   615
   End
   Begin VB.Image h56 
      Height          =   435
      Left            =   4800
      Top             =   4140
      Width           =   615
   End
   Begin VB.Image h52 
      Height          =   435
      Left            =   2040
      Top             =   4155
      Width           =   645
   End
   Begin VB.Image h51 
      Height          =   435
      Left            =   1380
      Top             =   4140
      Width           =   600
   End
   Begin VB.Image h50 
      Height          =   435
      Left            =   7575
      Top             =   3600
      Width           =   675
   End
   Begin VB.Image h48 
      Height          =   435
      Left            =   6165
      Top             =   3600
      Width           =   660
   End
   Begin VB.Image h44 
      Height          =   450
      Left            =   3435
      Top             =   3600
      Width           =   600
   End
   Begin VB.Image h49 
      Height          =   435
      Left            =   6870
      Top             =   3600
      Width           =   645
   End
   Begin VB.Image h45 
      Height          =   450
      Left            =   4110
      Top             =   3600
      Width           =   645
   End
   Begin VB.Image h43 
      Height          =   435
      Left            =   2730
      Top             =   3585
      Width           =   630
   End
   Begin VB.Image h47 
      Height          =   435
      Left            =   5505
      Top             =   3600
      Width           =   615
   End
   Begin VB.Image h46 
      Height          =   450
      Left            =   4800
      Top             =   3585
      Width           =   615
   End
   Begin VB.Image h42 
      Height          =   435
      Left            =   2070
      Top             =   3600
      Width           =   600
   End
   Begin VB.Image h41 
      Height          =   435
      Left            =   1365
      Top             =   3585
      Width           =   615
   End
   Begin VB.Image h40 
      Height          =   435
      Left            =   7560
      Top             =   2895
      Width           =   675
   End
   Begin VB.Image h38 
      Height          =   435
      Left            =   6195
      Top             =   2895
      Width           =   615
   End
   Begin VB.Image h34 
      Height          =   435
      Left            =   3465
      Top             =   2895
      Width           =   600
   End
   Begin VB.Image h39 
      Height          =   435
      Left            =   6855
      Top             =   2895
      Width           =   630
   End
   Begin VB.Image h35 
      Height          =   435
      Left            =   4140
      Top             =   2895
      Width           =   615
   End
   Begin VB.Image h33 
      Height          =   435
      Left            =   2760
      Top             =   2895
      Width           =   630
   End
   Begin VB.Image h37 
      Height          =   435
      Left            =   5490
      Top             =   2880
      Width           =   645
   End
   Begin VB.Image h36 
      Height          =   435
      Left            =   4815
      Top             =   2895
      Width           =   615
   End
   Begin VB.Image h32 
      Height          =   435
      Left            =   2070
      Top             =   2895
      Width           =   600
   End
   Begin VB.Image h31 
      Height          =   435
      Left            =   1395
      Top             =   2895
      Width           =   600
   End
   Begin VB.Image h30 
      Height          =   435
      Left            =   7545
      Top             =   2370
      Width           =   675
   End
   Begin VB.Image h28 
      Height          =   435
      Left            =   6180
      Top             =   2370
      Width           =   615
   End
   Begin VB.Image h24 
      Height          =   435
      Left            =   3450
      Top             =   2370
      Width           =   600
   End
   Begin VB.Image h29 
      Height          =   435
      Left            =   6855
      Top             =   2355
      Width           =   630
   End
   Begin VB.Image h25 
      Height          =   435
      Left            =   4125
      Top             =   2370
      Width           =   615
   End
   Begin VB.Image h23 
      Height          =   435
      Left            =   2760
      Top             =   2355
      Width           =   630
   End
   Begin VB.Image h27 
      Height          =   435
      Left            =   5490
      Top             =   2370
      Width           =   615
   End
   Begin VB.Image h26 
      Height          =   435
      Left            =   4830
      Top             =   2370
      Width           =   615
   End
   Begin VB.Image h22 
      Height          =   435
      Left            =   2070
      Top             =   2370
      Width           =   600
   End
   Begin VB.Image h21 
      Height          =   435
      Left            =   1395
      Top             =   2370
      Width           =   600
   End
   Begin VB.Image h20 
      Height          =   435
      Left            =   7545
      Top             =   1830
      Width           =   675
   End
   Begin VB.Image h18 
      Height          =   435
      Left            =   6180
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image h14 
      Height          =   435
      Left            =   3450
      Top             =   1830
      Width           =   600
   End
   Begin VB.Image h19 
      Height          =   435
      Left            =   6855
      Top             =   1815
      Width           =   630
   End
   Begin VB.Image h15 
      Height          =   435
      Left            =   4125
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image h13 
      Height          =   435
      Left            =   2760
      Top             =   1830
      Width           =   630
   End
   Begin VB.Image h17 
      Height          =   435
      Left            =   5490
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image h16 
      Height          =   435
      Left            =   4815
      Top             =   1830
      Width           =   615
   End
   Begin VB.Image h12 
      Height          =   435
      Left            =   2070
      Top             =   1830
      Width           =   600
   End
   Begin VB.Image h11 
      Height          =   435
      Left            =   1395
      Top             =   1830
      Width           =   600
   End
   Begin VB.Image h10 
      Height          =   435
      Left            =   7545
      Top             =   1305
      Width           =   675
   End
   Begin VB.Image h8 
      Height          =   435
      Left            =   6195
      Top             =   1305
      Width           =   615
   End
   Begin VB.Image h4 
      Height          =   435
      Left            =   3450
      Top             =   1305
      Width           =   600
   End
   Begin VB.Image h9 
      Height          =   435
      Left            =   6855
      Top             =   1290
      Width           =   630
   End
   Begin VB.Image h5 
      Height          =   435
      Left            =   4125
      Top             =   1305
      Width           =   615
   End
   Begin VB.Image h3 
      Height          =   435
      Left            =   2760
      Top             =   1305
      Width           =   630
   End
   Begin VB.Image h7 
      Height          =   435
      Left            =   5490
      Top             =   1305
      Width           =   615
   End
   Begin VB.Image h6 
      Height          =   435
      Left            =   4815
      Top             =   1305
      Width           =   615
   End
   Begin VB.Image h2 
      Height          =   435
      Left            =   2070
      Top             =   1305
      Width           =   600
   End
   Begin VB.Image h1 
      Height          =   435
      Left            =   1395
      Top             =   1305
      Width           =   600
   End
   Begin VB.Image d80 
      Height          =   435
      Left            =   7575
      Top             =   5235
      Width           =   630
   End
   Begin VB.Image d78 
      Height          =   435
      Left            =   6195
      Top             =   5235
      Width           =   585
   End
   Begin VB.Image d74 
      Height          =   405
      Left            =   3435
      Top             =   5250
      Width           =   630
   End
   Begin VB.Image d79 
      Height          =   435
      Left            =   6855
      Top             =   5220
      Width           =   645
   End
   Begin VB.Image d75 
      Height          =   405
      Left            =   4110
      Top             =   5250
      Width           =   630
   End
   Begin VB.Image d73 
      Height          =   435
      Left            =   2730
      Top             =   5220
      Width           =   630
   End
   Begin VB.Image d77 
      Height          =   435
      Left            =   5505
      Top             =   5250
      Width           =   600
   End
   Begin VB.Image d76 
      Height          =   435
      Left            =   4815
      Top             =   5250
      Width           =   645
   End
   Begin VB.Image d72 
      Height          =   435
      Left            =   2040
      Top             =   5235
      Width           =   645
   End
   Begin VB.Image d71 
      Height          =   435
      Left            =   1365
      Top             =   5235
      Width           =   615
   End
   Begin VB.Image d70 
      Height          =   435
      Left            =   7545
      Top             =   4695
      Width           =   675
   End
   Begin VB.Image d68 
      Height          =   435
      Left            =   6180
      Top             =   4695
      Width           =   615
   End
   Begin VB.Image d64 
      Height          =   435
      Left            =   3420
      Top             =   4695
      Width           =   645
   End
   Begin VB.Image d69 
      Height          =   435
      Left            =   6855
      Top             =   4695
      Width           =   660
   End
   Begin VB.Image d65 
      Height          =   435
      Left            =   4095
      Top             =   4695
      Width           =   660
   End
   Begin VB.Image d63 
      Height          =   435
      Left            =   2730
      Top             =   4695
      Width           =   645
   End
   Begin VB.Image d67 
      Height          =   435
      Left            =   5490
      Top             =   4695
      Width           =   645
   End
   Begin VB.Image d66 
      Height          =   435
      Left            =   4785
      Top             =   4695
      Width           =   660
   End
   Begin VB.Image d62 
      Height          =   435
      Left            =   2040
      Top             =   4695
      Width           =   645
   End
   Begin VB.Image d61 
      Height          =   435
      Left            =   1365
      Top             =   4695
      Width           =   630
   End
   Begin VB.Image d60 
      Height          =   435
      Left            =   7560
      Top             =   4140
      Width           =   675
   End
   Begin VB.Image d58 
      Height          =   435
      Left            =   6195
      Top             =   4140
      Width           =   615
   End
   Begin VB.Image d54 
      Height          =   435
      Left            =   3450
      Top             =   4125
      Width           =   615
   End
   Begin VB.Image d59 
      Height          =   435
      Left            =   6870
      Top             =   4140
      Width           =   630
   End
   Begin VB.Image d55 
      Height          =   435
      Left            =   4125
      Top             =   4125
      Width           =   615
   End
   Begin VB.Image d53 
      Height          =   435
      Left            =   2745
      Top             =   4125
      Width           =   630
   End
   Begin VB.Image d57 
      Height          =   435
      Left            =   5490
      Top             =   4140
      Width           =   615
   End
   Begin VB.Image d56 
      Height          =   435
      Left            =   4800
      Top             =   4125
      Width           =   615
   End
   Begin VB.Image d52 
      Height          =   435
      Left            =   2040
      Top             =   4140
      Width           =   645
   End
   Begin VB.Image d51 
      Height          =   435
      Left            =   1380
      Top             =   4125
      Width           =   600
   End
   Begin VB.Image d50 
      Height          =   435
      Left            =   7575
      Top             =   3585
      Width           =   675
   End
   Begin VB.Image d48 
      Height          =   435
      Left            =   6165
      Top             =   3585
      Width           =   660
   End
   Begin VB.Image d44 
      Height          =   450
      Left            =   3435
      Top             =   3585
      Width           =   600
   End
   Begin VB.Image d49 
      Height          =   435
      Left            =   6870
      Top             =   3585
      Width           =   645
   End
   Begin VB.Image d45 
      Height          =   450
      Left            =   4110
      Top             =   3585
      Width           =   645
   End
   Begin VB.Image d43 
      Height          =   435
      Left            =   2730
      Top             =   3570
      Width           =   630
   End
   Begin VB.Image d47 
      Height          =   435
      Left            =   5505
      Top             =   3585
      Width           =   615
   End
   Begin VB.Image d46 
      Height          =   450
      Left            =   4800
      Top             =   3570
      Width           =   615
   End
   Begin VB.Image d42 
      Height          =   435
      Left            =   2070
      Top             =   3585
      Width           =   600
   End
   Begin VB.Image d41 
      Height          =   435
      Left            =   1365
      Top             =   3570
      Width           =   615
   End
   Begin VB.Image d40 
      Height          =   435
      Left            =   7560
      Top             =   2880
      Width           =   675
   End
   Begin VB.Image d38 
      Height          =   435
      Left            =   6195
      Top             =   2880
      Width           =   615
   End
   Begin VB.Image d34 
      Height          =   435
      Left            =   3465
      Top             =   2880
      Width           =   600
   End
   Begin VB.Image d39 
      Height          =   435
      Left            =   6855
      Top             =   2880
      Width           =   630
   End
   Begin VB.Image d35 
      Height          =   435
      Left            =   4140
      Top             =   2880
      Width           =   615
   End
   Begin VB.Image d33 
      Height          =   435
      Left            =   2760
      Top             =   2880
      Width           =   630
   End
   Begin VB.Image d37 
      Height          =   435
      Left            =   5490
      Top             =   2865
      Width           =   645
   End
   Begin VB.Image d36 
      Height          =   435
      Left            =   4815
      Top             =   2880
      Width           =   615
   End
   Begin VB.Image d32 
      Height          =   435
      Left            =   2070
      Top             =   2880
      Width           =   600
   End
   Begin VB.Image d31 
      Height          =   435
      Left            =   1395
      Top             =   2880
      Width           =   600
   End
   Begin VB.Image d30 
      Height          =   435
      Left            =   7545
      Top             =   2355
      Width           =   675
   End
   Begin VB.Image d28 
      Height          =   435
      Left            =   6180
      Top             =   2355
      Width           =   615
   End
   Begin VB.Image d24 
      Height          =   435
      Left            =   3450
      Top             =   2355
      Width           =   600
   End
   Begin VB.Image d29 
      Height          =   435
      Left            =   6855
      Top             =   2340
      Width           =   630
   End
   Begin VB.Image d25 
      Height          =   435
      Left            =   4125
      Top             =   2355
      Width           =   615
   End
   Begin VB.Image d23 
      Height          =   435
      Left            =   2760
      Top             =   2355
      Width           =   630
   End
   Begin VB.Image d27 
      Height          =   435
      Left            =   5490
      Top             =   2355
      Width           =   615
   End
   Begin VB.Image d26 
      Height          =   435
      Left            =   4830
      Top             =   2355
      Width           =   615
   End
   Begin VB.Image d22 
      Height          =   435
      Left            =   2070
      Top             =   2355
      Width           =   600
   End
   Begin VB.Image d21 
      Height          =   435
      Left            =   1395
      Top             =   2355
      Width           =   600
   End
   Begin VB.Image d20 
      Height          =   435
      Left            =   7545
      Top             =   1815
      Width           =   675
   End
   Begin VB.Image d18 
      Height          =   435
      Left            =   6180
      Top             =   1815
      Width           =   615
   End
   Begin VB.Image d14 
      Height          =   435
      Left            =   3450
      Top             =   1815
      Width           =   600
   End
   Begin VB.Image d19 
      Height          =   435
      Left            =   6855
      Top             =   1800
      Width           =   630
   End
   Begin VB.Image d15 
      Height          =   435
      Left            =   4125
      Top             =   1815
      Width           =   615
   End
   Begin VB.Image d13 
      Height          =   435
      Left            =   2760
      Top             =   1815
      Width           =   630
   End
   Begin VB.Image d17 
      Height          =   435
      Left            =   5490
      Top             =   1815
      Width           =   615
   End
   Begin VB.Image d16 
      Height          =   435
      Left            =   4815
      Top             =   1815
      Width           =   615
   End
   Begin VB.Image d12 
      Height          =   435
      Left            =   2070
      Top             =   1815
      Width           =   600
   End
   Begin VB.Image d11 
      Height          =   435
      Left            =   1395
      Top             =   1815
      Width           =   600
   End
   Begin VB.Image d10 
      Height          =   435
      Left            =   7545
      Top             =   1290
      Width           =   675
   End
   Begin VB.Image d8 
      Height          =   435
      Left            =   6180
      Top             =   1290
      Width           =   615
   End
   Begin VB.Image d4 
      Height          =   435
      Left            =   3450
      Top             =   1290
      Width           =   600
   End
   Begin VB.Image d9 
      Height          =   435
      Left            =   6855
      Top             =   1275
      Width           =   630
   End
   Begin VB.Image d5 
      Height          =   435
      Left            =   4125
      Top             =   1290
      Width           =   615
   End
   Begin VB.Image d3 
      Height          =   435
      Left            =   2760
      Top             =   1290
      Width           =   630
   End
   Begin VB.Image d7 
      Height          =   435
      Left            =   5490
      Top             =   1290
      Width           =   615
   End
   Begin VB.Image d6 
      Height          =   435
      Left            =   4815
      Top             =   1290
      Width           =   615
   End
   Begin VB.Image d2 
      Height          =   435
      Left            =   2070
      Top             =   1290
      Width           =   600
   End
   Begin VB.Image d1 
      Height          =   435
      Left            =   1395
      Top             =   1290
      Width           =   600
   End
End
Attribute VB_Name = "board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "Kernel32" () As Long
Private Sub betone_Click()
hits = 0
hits_label.Caption = hits
hits_label.Refresh
If amountbet = 0 Then
    Call enablechecks
    Call clearall
    End If
If amountbet >= 4 Or dollars = 0 Then
Exit Sub
End If
amountbet = amountbet + 1
coinsin_label.Caption = amountbet
dollars = dollars - 1
credits_label.Caption = dollars

If dollars >= 4 Then
play4_button.Picture = LoadResPicture("play4_lit", bitmap)
play4_button.Enabled = True
start_button.Enabled = True
erase_button.Enabled = True
play4_button.SetFocus
start_button.Picture = LoadResPicture("start_lit", bitmap)
erase_button.Picture = LoadResPicture("erase_lit", bitmap)
bottom_display.Picture = LoadResPicture("play4coins", bitmap)
Else
play4_button.Picture = LoadResPicture("play4_lit", bitmap)
play4_button.Enabled = True
start_button.Enabled = True
erase_button.Enabled = True
start_button.Picture = LoadResPicture("start_lit", bitmap)
erase_button.Picture = LoadResPicture("erase_lit", bitmap)
play4_button.Picture = LoadResPicture("play4_dim", bitmap)
play4_button.Enabled = False
bottom_display.Picture = LoadResPicture("play4coins", bitmap)
End If
End Sub

Private Sub change_button_Click()
options.bet_total.Caption = Format(bettotal, "#,###,###")
options.bet_total.Refresh
options.credits_won.Caption = Format(hopperempty, "#,###,###")
options.credits_won.Refresh
options.deals_label.Caption = Format(totaldeals, "#,###,###")
options.deals_label.Refresh
If bettotal > 0 And hopperempty > 0 Then
Call options.returnrate
End If
options.Show
options.cmdOK.Enabled = False
End Sub
Private Sub erase_button_Click()
Call clearchecks
Call clearall
boxes_checked = 0
spots_marked.Caption = boxes_checked
spots_marked.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
autostop = 1
End If


End Sub

Private Sub form_load()
If getgame = 2 Then
Call clearall
Call disablechecks
Call gamesettings
dollars = lastcredits
coinsin_label.Caption = amountbet
credits_label.Caption = dollars
spots_marked.Caption = boxes_checked
'init menu items
bottom_display.Picture = LoadResPicture("insert_coins", bitmap)
board.Picture = LoadResPicture("board2", bitmap)
start_button.Picture = LoadResPicture("start_dim", bitmap)
start_button.DownPicture = LoadResPicture("start_lit_down", bitmap)
erase_button.Picture = LoadResPicture("erase_dim", bitmap)
erase_button.DownPicture = LoadResPicture("erase_lit_down", bitmap)
betone.Picture = LoadResPicture("betone_lit", bitmap)
betone.DownPicture = LoadResPicture("betone_lit_down", bitmap)
play4_button.Picture = LoadResPicture("play4_lit", bitmap)
play4_button.DownPicture = LoadResPicture("play4_lit_down", bitmap)
play4_button.Enabled = False
change_button.Picture = LoadResPicture("change_button_up", bitmap)
change_button.DownPicture = LoadResPicture("change_button_down", bitmap)
End If
If getgame = 3 Then
Call clearall
Call disablechecks
Call gamesettings
coinsin_label.Caption = amountbet
credits_label.Caption = dollars
spots_marked.Caption = boxes_checked
bottom_display.Picture = LoadResPicture("insert_coins", bitmap)
board.Picture = LoadResPicture("board2", bitmap)
start_button.Picture = LoadResPicture("start_dim", bitmap)
start_button.DownPicture = LoadResPicture("start_lit_down", bitmap)
erase_button.Picture = LoadResPicture("erase_dim", bitmap)
erase_button.DownPicture = LoadResPicture("erase_lit_down", bitmap)
betone.Picture = LoadResPicture("betone_lit", bitmap)
betone.DownPicture = LoadResPicture("betone_lit_down", bitmap)
play4_button.Picture = LoadResPicture("play4_lit", bitmap)
play4_button.DownPicture = LoadResPicture("play4_lit_down", bitmap)
play4_button.Enabled = False
change_button.Picture = LoadResPicture("change_button_up", bitmap)
change_button.DownPicture = LoadResPicture("change_button_down", bitmap)
End If
If getgame = 1 Then
Call disablechecks
Call clearall
Call clearchecks
coinsin_label.Caption = amountbet
credits_label.Caption = dollars
spots_marked.Caption = boxes_checked
'init menu items
bottom_display.Picture = LoadResPicture("insert_coins", bitmap)
Me.Picture = LoadResPicture("board2", bitmap)
start_button.Picture = LoadResPicture("start_dim", bitmap)
start_button.DownPicture = LoadResPicture("start_lit_down", bitmap)
erase_button.Picture = LoadResPicture("erase_dim", bitmap)
erase_button.DownPicture = LoadResPicture("erase_lit_down", bitmap)
betone.Picture = LoadResPicture("betone_lit", bitmap)
betone.DownPicture = LoadResPicture("betone_lit_down", bitmap)
play4_button.Picture = LoadResPicture("play4_lit", bitmap)
play4_button.DownPicture = LoadResPicture("play4_lit_down", bitmap)
play4_button.Enabled = False
change_button.Picture = LoadResPicture("change_button_up", bitmap)
change_button.DownPicture = LoadResPicture("change_button_down", bitmap)
End If
If dollars <= 0 Then
    betone.Enabled = False
    End If
End Sub
Private Sub c1_Click()
d1.Picture = LoadPicture
h1.Picture = LoadPicture
If c1_checked = 0 Then
c1.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c1_checked = c1_checked + 1
Else
c1.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c1_checked = c1_checked - 1
End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c1.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c1_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Spot Marked Monitor"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub
Private Sub c10_Click()
d10.Picture = LoadPicture
h10.Picture = LoadPicture
If c10_checked = 0 Then
c10.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c10_checked = c10_checked + 1
Else
c10.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c10_checked = c10_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c10.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c10_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If

End Sub

Private Sub c11_Click()
d11.Picture = LoadPicture
h11.Picture = LoadPicture
If c11_checked = 0 Then
c11.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c11_checked = c11_checked + 1
Else
c11.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c11_checked = c11_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c11.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c11_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c12_Click()
d12.Picture = LoadPicture
h12.Picture = LoadPicture
If c12_checked = 0 Then
c12.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c12_checked = c12_checked + 1
Else
c12.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c12_checked = c12_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c12.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c12_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c13_Click()
d13.Picture = LoadPicture
h13.Picture = LoadPicture
If c13_checked = 0 Then
c13.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c13_checked = c13_checked + 1
Else
c13.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c13_checked = c13_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c13.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c13_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c14_Click()
d14.Picture = LoadPicture
h14.Picture = LoadPicture
If c14_checked = 0 Then
c14.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c14_checked = c14_checked + 1
Else
c14.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c14_checked = c14_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c14.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c14_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c15_Click()
d15.Picture = LoadPicture
h15.Picture = LoadPicture
If c15_checked = 0 Then
c15.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c15_checked = c15_checked + 1
Else
c15.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c15_checked = c15_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c15.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c15_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c16_Click()
d16.Picture = LoadPicture
h16.Picture = LoadPicture
If c16_checked = 0 Then
c16.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c16_checked = c16_checked + 1
Else
c16.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c16_checked = c16_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c16.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c16_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c17_Click()
d17.Picture = LoadPicture
h17.Picture = LoadPicture
If c17_checked = 0 Then
c17.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c17_checked = c17_checked + 1
Else
c17.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c17_checked = c17_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c17.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c17_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c18_Click()
d18.Picture = LoadPicture
h18.Picture = LoadPicture
If c18_checked = 0 Then
c18.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c18_checked = c18_checked + 1
Else
c18.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c18_checked = c18_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c18.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c18_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c19_Click()
d19.Picture = LoadPicture
h19.Picture = LoadPicture
If c19_checked = 0 Then
c19.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c19_checked = c19_checked + 1
Else
c19.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c19_checked = c19_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c19.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c19_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c2_Click()
d2.Picture = LoadPicture
h2.Picture = LoadPicture
If c2_checked = 0 Then
c2.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c2_checked = c2_checked + 1
Else
c2.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c2_checked = c2_checked - 1
End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c2.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c2_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c20_Click()
d20.Picture = LoadPicture
h20.Picture = LoadPicture
If c20_checked = 0 Then
c20.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c20_checked = c20_checked + 1
Else
c20.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c20_checked = c20_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c20.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c20_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c21_Click()
d21.Picture = LoadPicture
h21.Picture = LoadPicture
If c21_checked = 0 Then
c21.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c21_checked = c21_checked + 1
Else
c21.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c21_checked = c21_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c21.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c21_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c22_Click()
d22.Picture = LoadPicture
h22.Picture = LoadPicture
If c22_checked = 0 Then
c22.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c22_checked = c22_checked + 1
Else
c22.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c22_checked = c22_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c22.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c22_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c23_Click()
d23.Picture = LoadPicture
h23.Picture = LoadPicture
If c23_checked = 0 Then
c23.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c23_checked = c23_checked + 1
Else
c23.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c23_checked = c23_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c23.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c23_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c24_Click()
d24.Picture = LoadPicture
h24.Picture = LoadPicture
If c24_checked = 0 Then
c24.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c24_checked = c24_checked + 1
Else
c24.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c24_checked = c24_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c24.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c24_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c25_Click()
d25.Picture = LoadPicture
h25.Picture = LoadPicture
If c25_checked = 0 Then
c25.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c25_checked = c25_checked + 1
Else
c25.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c25_checked = c25_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c25.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c25_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c26_Click()
d26.Picture = LoadPicture
h26.Picture = LoadPicture
If c26_checked = 0 Then
c26.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c26_checked = c26_checked + 1
Else
c26.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c26_checked = c26_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c26.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c26_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c27_Click()
d27.Picture = LoadPicture
h27.Picture = LoadPicture
If c27_checked = 0 Then
c27.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c27_checked = c27_checked + 1
Else
c27.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c27_checked = c27_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c27.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c27_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c28_Click()
d28.Picture = LoadPicture
h28.Picture = LoadPicture
If c28_checked = 0 Then
c28.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c28_checked = c28_checked + 1
Else
c28.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c28_checked = c28_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c28.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c28_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c29_Click()
d29.Picture = LoadPicture
h29.Picture = LoadPicture
If c29_checked = 0 Then
c29.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c29_checked = c29_checked + 1
Else
c29.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c29_checked = c29_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c29.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c29_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c3_Click()
d3.Picture = LoadPicture
h3.Picture = LoadPicture
If c3_checked = 0 Then
c3.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c3_checked = c3_checked + 1
Else
c3.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c3_checked = c3_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c3.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c3_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c30_Click()
d30.Picture = LoadPicture
h30.Picture = LoadPicture
If c30_checked = 0 Then
c30.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c30_checked = c30_checked + 1
Else
c30.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c30_checked = c30_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c30.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c30_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c31_Click()
d31.Picture = LoadPicture
h31.Picture = LoadPicture
If c31_checked = 0 Then
c31.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c31_checked = c31_checked + 1
Else
c31.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c31_checked = c31_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c31.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c31_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c32_Click()
d32.Picture = LoadPicture
h32.Picture = LoadPicture
If c32_checked = 0 Then
c32.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c32_checked = c32_checked + 1
Else
c32.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c32_checked = c32_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c32.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c32_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c33_Click()
d33.Picture = LoadPicture
h33.Picture = LoadPicture
If c33_checked = 0 Then
c33.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c33_checked = c33_checked + 1
Else
c33.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c33_checked = c33_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c33.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c33_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c34_Click()
d34.Picture = LoadPicture
h34.Picture = LoadPicture
If c34_checked = 0 Then
c34.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c34_checked = c34_checked + 1
Else
c34.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c34_checked = c34_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c34.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c34_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c35_Click()
d35.Picture = LoadPicture
h35.Picture = LoadPicture
If c35_checked = 0 Then
c35.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c35_checked = c35_checked + 1
Else
c35.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c35_checked = c35_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c35.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c35_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c36_Click()
d36.Picture = LoadPicture
h36.Picture = LoadPicture
If c36_checked = 0 Then
c36.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c36_checked = c36_checked + 1
Else
c36.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c36_checked = c36_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c36.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c36_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c37_Click()
d37.Picture = LoadPicture
h37.Picture = LoadPicture
If c37_checked = 0 Then
c37.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c37_checked = c37_checked + 1
Else
c37.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c37_checked = c37_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c37.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c37_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c38_Click()
d20.Picture = LoadPicture
h38.Picture = LoadPicture
If c38_checked = 0 Then
c38.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c38_checked = c38_checked + 1
Else
c38.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c38_checked = c38_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c38Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c38_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c39_Click()
d39.Picture = LoadPicture
h39.Picture = LoadPicture
If c39_checked = 0 Then
c39.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c39_checked = c39_checked + 1
Else
c39.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c39_checked = c39_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c39.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c39_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c4_Click()
d4.Picture = LoadPicture
h4.Picture = LoadPicture
If c4_checked = 0 Then
c4.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c4_checked = c4_checked + 1
Else
c4.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c4_checked = c4_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c4.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c4_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c40_Click()
d40.Picture = LoadPicture
h40.Picture = LoadPicture
If c40_checked = 0 Then
c40.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c40_checked = c40_checked + 1
Else
c40.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c40_checked = c40_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c40.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c40_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c41_Click()
d41.Picture = LoadPicture
h41.Picture = LoadPicture
If c41_checked = 0 Then
c41.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c41_checked = c41_checked + 1
Else
c41.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c41_checked = c41_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c41.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c41_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c42_Click()
d42.Picture = LoadPicture
h42.Picture = LoadPicture
If c42_checked = 0 Then
c42.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c42_checked = c42_checked + 1
Else
c42.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c42_checked = c42_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c42.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c42_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c43_Click()
d43.Picture = LoadPicture
h43.Picture = LoadPicture
If c43_checked = 0 Then
c43.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c43_checked = c43_checked + 1
Else
c43.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c43_checked = c43_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c43.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c43_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c44_Click()
d44.Picture = LoadPicture
h44.Picture = LoadPicture
If c44_checked = 0 Then
c44.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c44_checked = c44_checked + 1
Else
c44.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c44_checked = c44_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c44.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c44_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c45_Click()
d45.Picture = LoadPicture
h45.Picture = LoadPicture
If c45_checked = 0 Then
c45.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c45_checked = c45_checked + 1
Else
c45.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c45_checked = c45_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c45.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c45_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c46_Click()
d46.Picture = LoadPicture
h46.Picture = LoadPicture
If c46_checked = 0 Then
c46.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c46_checked = c46_checked + 1
Else
c46.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c46_checked = c46_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c46.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c46_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c47_Click()
d47.Picture = LoadPicture
h47.Picture = LoadPicture
If c47_checked = 0 Then
c47.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c47_checked = c47_checked + 1
Else
c47.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c47_checked = c47_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c47.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c47_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c48_Click()
d48.Picture = LoadPicture
h48.Picture = LoadPicture
If c48_checked = 0 Then
c48.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c48_checked = c48_checked + 1
Else
c48.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c48_checked = c48_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c48.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c48_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub



Private Sub c49_Click()
d49.Picture = LoadPicture
h49.Picture = LoadPicture
If c49_checked = 0 Then
c49.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c49_checked = c49_checked + 1
Else
c49.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c49_checked = c49_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c49.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c49_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c5_Click()
d5.Picture = LoadPicture
h5.Picture = LoadPicture
If c5_checked = 0 Then
c5.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c5_checked = c5_checked + 1
Else
c5.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c5_checked = c5_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c5.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c5_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c50_Click()
d50.Picture = LoadPicture
h50.Picture = LoadPicture
If c50_checked = 0 Then
c50.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c50_checked = c50_checked + 1
Else
c50.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c50_checked = c50_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c50.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c50_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c51_Click()
d51.Picture = LoadPicture
h51.Picture = LoadPicture
If c51_checked = 0 Then
c51.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c51_checked = c51_checked + 1
Else
c51.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c51_checked = c51_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c51.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c51_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c52_Click()
d52.Picture = LoadPicture
h52.Picture = LoadPicture
If c52_checked = 0 Then
c52.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c52_checked = c52_checked + 1
Else
c52.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c52_checked = c52_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c52.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c52_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c53_Click()
d53.Picture = LoadPicture
h53.Picture = LoadPicture
If c53_checked = 0 Then
c53.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c53_checked = c53_checked + 1
Else
c53.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c53_checked = c53_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c53.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c53_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c54_Click()
d54.Picture = LoadPicture
h54.Picture = LoadPicture
If c54_checked = 0 Then
c54.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c54_checked = c54_checked + 1
Else
c54.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c54_checked = c54_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c54.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c54_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c55_Click()
d55.Picture = LoadPicture
h55.Picture = LoadPicture
If c55_checked = 0 Then
c55.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c55_checked = c55_checked + 1
Else
c55.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c55_checked = c55_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c55.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c55_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c56_Click()
d56.Picture = LoadPicture
h56.Picture = LoadPicture
If c56_checked = 0 Then
c56.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c56_checked = c56_checked + 1
Else
c56.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c56_checked = c56_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c56.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c56_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c57_Click()
d57.Picture = LoadPicture
h57.Picture = LoadPicture
If c57_checked = 0 Then
c57.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c57_checked = c57_checked + 1
Else
c57.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c57_checked = c57_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c57.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c57_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c58_Click()
d58.Picture = LoadPicture
h58.Picture = LoadPicture
If c58_checked = 0 Then
c58.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c58_checked = c58_checked + 1
Else
c58.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c58_checked = c58_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c58.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c58_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c59_Click()
d59.Picture = LoadPicture
h59.Picture = LoadPicture
If c59_checked = 0 Then
c59.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c59_checked = c59_checked + 1
Else
c59.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c59_checked = c59_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c59.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c59_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c6_Click()
d6.Picture = LoadPicture
h6.Picture = LoadPicture
If c6_checked = 0 Then
c6.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c6_checked = c6_checked + 1
Else
c6.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c6_checked = c6_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c6.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c6_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c60_Click()
d60.Picture = LoadPicture
h60.Picture = LoadPicture
If c60_checked = 0 Then
c60.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c60_checked = c60_checked + 1
Else
c60.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c60_checked = c60_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c60.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c60_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c61_Click()
d61.Picture = LoadPicture
h61.Picture = LoadPicture
If c61_checked = 0 Then
c61.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c61_checked = c61_checked + 1
Else
c61.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c61_checked = c61_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c61.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c61_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c62_Click()
d62.Picture = LoadPicture
h62.Picture = LoadPicture
If c62_checked = 0 Then
c62.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c62_checked = c62_checked + 1
Else
c62.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c62_checked = c62_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c62.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c62_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c63_Click()
d63.Picture = LoadPicture
h63.Picture = LoadPicture
If c63_checked = 0 Then
c63.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c63_checked = c63_checked + 1
Else
c63.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c63_checked = c63_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c63.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c63_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c64_Click()
d64.Picture = LoadPicture
h64.Picture = LoadPicture
If c64_checked = 0 Then
c64.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c64_checked = c64_checked + 1
Else
c64.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c64_checked = c64_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c64.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c64_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c65_Click()
d65.Picture = LoadPicture
h65.Picture = LoadPicture
If c65_checked = 0 Then
c65.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c65_checked = c65_checked + 1
Else
c65.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c65_checked = c65_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c65.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c65_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c66_Click()
d66.Picture = LoadPicture
h66.Picture = LoadPicture
If c66_checked = 0 Then
c66.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c66_checked = c66_checked + 1
Else
c66.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c66_checked = c66_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c66.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c66_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c67_Click()
d67.Picture = LoadPicture
h67.Picture = LoadPicture
If c67_checked = 0 Then
c67.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c67_checked = c67_checked + 1
Else
c67.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c67_checked = c67_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c67.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c67_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c68_Click()
d68.Picture = LoadPicture
h68.Picture = LoadPicture
If c68_checked = 0 Then
c68.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c68_checked = c68_checked + 1
Else
c68.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c68_checked = c68_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c68.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c68_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c69_Click()
d69.Picture = LoadPicture
h69.Picture = LoadPicture
If c69_checked = 0 Then
c69.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c69_checked = c69_checked + 1
Else
c69.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c69_checked = c69_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c69.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c69_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c7_Click()
d7.Picture = LoadPicture
h7.Picture = LoadPicture
If c7_checked = 0 Then
c7.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c7_checked = c7_checked + 1
Else
c7.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c7_checked = c7_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c7.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c7_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c70_Click()
d70.Picture = LoadPicture
h70.Picture = LoadPicture
If c70_checked = 0 Then
c70.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c70_checked = c70_checked + 1
Else
c70.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c70_checked = c70_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c70.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c70_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c71_Click()
d71.Picture = LoadPicture
h71.Picture = LoadPicture
If c71_checked = 0 Then
c71.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c71_checked = c71_checked + 1
Else
c71.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c71_checked = c71_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c71.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c71_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c72_Click()
d72.Picture = LoadPicture
h72.Picture = LoadPicture
If c72_checked = 0 Then
c72.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c72_checked = c72_checked + 1
Else
c72.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c72_checked = c72_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c72.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c72_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c73_Click()
d73.Picture = LoadPicture
h73.Picture = LoadPicture
If c73_checked = 0 Then
c73.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c73_checked = c73_checked + 1
Else
c73.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c73_checked = c73_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c73.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c73_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c74_Click()
d74.Picture = LoadPicture
h74.Picture = LoadPicture
If c74_checked = 0 Then
c74.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c74_checked = c74_checked + 1
Else
c74.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c74_checked = c74_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c74.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c74_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c75_Click()
d75.Picture = LoadPicture
h75.Picture = LoadPicture
If c75_checked = 0 Then
c75.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c75_checked = c75_checked + 1
Else
c75.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c75_checked = c75_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c75.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c75_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c76_Click()
d76.Picture = LoadPicture
h76.Picture = LoadPicture
If c76_checked = 0 Then
c76.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c76_checked = c76_checked + 1
Else
c76.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c76_checked = c76_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c76.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c76_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c77_Click()
d77.Picture = LoadPicture
h77.Picture = LoadPicture
If c77_checked = 0 Then
c77.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c77_checked = c77_checked + 1
Else
c77.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c77_checked = c77_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c77.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c77_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c78_Click()
d78.Picture = LoadPicture
h78.Picture = LoadPicture
If c78_checked = 0 Then
c78.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c78_checked = c78_checked + 1
Else
c78.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c78_checked = c78_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c78.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c78_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c79_Click()
d79.Picture = LoadPicture
h79.Picture = LoadPicture
If c79_checked = 0 Then
c79.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c79_checked = c79_checked + 1
Else
c79.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c79_checked = c79_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c79.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c79_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c8_Click()
d8.Picture = LoadPicture
h8.Picture = LoadPicture
If c8_checked = 0 Then
c8.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c8_checked = c8_checked + 1
Else
c8.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c8_checked = c8_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c8.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c8_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c80_Click()
d80.Picture = LoadPicture
h80.Picture = LoadPicture
If c80_checked = 0 Then
c80.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c80_checked = c80_checked + 1
Else
c80.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c80_checked = c80_checked - 1

End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c80.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c80_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub

Private Sub c9_Click()
d9.Picture = LoadPicture
h9.Picture = LoadPicture
If c9_checked = 0 Then
c9.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
c9_checked = c9_checked + 1
Else
c9.Picture = LoadPicture
boxes_checked = boxes_checked - 1
c9_checked = c9_checked - 1
End If
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Call checksound
If boxes_checked > 10 Then
    c9.Picture = LoadPicture
    boxes_checked = boxes_checked - 1
    c9_checked = 0
    MsgBox "No more than 10 spots allowed please.", vbOKOnly, "Crazy Clicker"
    spots_marked.Caption = boxes_checked
    spots_marked.Refresh
    Exit Sub
    End If
End Sub
Public Function go()
change_button.Enabled = False
previousgame = 1
If amountbet = 0 Then
    Exit Function
    End If
If amountbet > 4 Then
    amountbet = 4
    End If
bettotal = bettotal + amountbet
start_button.Picture = LoadResPicture("start_dim", bitmap)
start_button.Enabled = False
erase_button.Picture = LoadResPicture("erase_dim", bitmap)
erase_button.Enabled = False
betone.Picture = LoadResPicture("betone_dim", bitmap)
betone.Enabled = False
play4_button.Enabled = False
play4_button.Picture = LoadResPicture("play4_dim", bitmap)
bottom_display.Picture = LoadPicture
spots_marked.Caption = boxes_checked
spots_marked.Refresh
Dim casehit As Integer
Dim r As Integer
Dim v As Variant
Call clearall
totaldeals = totaldeals + 1

' declare hit variables
Dim c1_hit_already, c2_hit_already, c3_hit_already, c4_hit_already, c5_hit_already, c6_hit_already As Single
Dim c7_hit_already, c8_hit_already, c9_hit_already, c10_hit_already, c11_hit_already, c12_hit_already As Single
Dim c13_hit_already, c14_hit_already, c15_hit_already, c16_hit_already, c17_hit_already, c18_hit_already As Single
Dim c19_hit_already, c20_hit_already, c21_hit_already, c22_hit_already, c23_hit_already, c24_hit_already As Single
Dim c25_hit_already, c26_hit_already, c27_hit_already, c28_hit_already, c29_hit_already, c30_hit_already As Single
Dim c31_hit_already, c32_hit_already, c33_hit_already, c34_hit_already, c35_hit_already, c36_hit_already As Single
Dim c37_hit_already, c38_hit_already, c39_hit_already, c40_hit_already, c41_hit_already, c42_hit_already As Single
Dim c43_hit_already, c44_hit_already, c45_hit_already, c46_hit_already, c47_hit_already, c48_hit_already As Single
Dim c49_hit_already, c50_hit_already, c51_hit_already As Single
Dim c52_hit_already, c53_hit_already, c54_hit_already As Single
Dim c55_hit_already, c56_hit_already, c57_hit_already As Single
Dim c58_hit_already, c59_hit_already, c60_hit_already As Single
Dim c61_hit_already, c62_hit_already, c63_hit_already As Single
Dim c64_hit_already, c65_hit_already, c66_hit_already As Single
Dim c67_hit_already, c68_hit_already, c69_hit_already As Single
Dim c70_hit_already, c71_hit_already, c72_hit_already As Single
Dim c73_hit_already, c74_hit_already, c75_hit_already As Single
Dim c76_hit_already, c77_hit_already, c78_hit_already As Single
Dim c79_hit_already, c80_hit_already As Single
'reset already hit
c1_hit_already = 0
c2_hit_already = 0
c3_hit_already = 0
c4_hit_already = 0
c5_hit_already = 0
c6_hit_already = 0
c7_hit_already = 0
c8_hit_already = 0
c9_hit_already = 0
c10_hit_already = 0
c11_hit_already = 0
c12_hit_already = 0
c13_hit_already = 0
c14_hit_already = 0
c15_hit_already = 0
c16_hit_already = 0
c17_hit_already = 0
c18_hit_already = 0
c19_hit_already = 0
c20_hit_already = 0
c21_hit_already = 0
c22_hit_already = 0
c23_hit_already = 0
c24_hit_already = 0
c25_hit_already = 0
c26_hit_already = 0
c27_hit_already = 0
c28_hit_already = 0
c29_hit_already = 0
c30_hit_already = 0
c31_hit_already = 0
c32_hit_already = 0
c33_hit_already = 0
c34_hit_already = 0
c35_hit_already = 0
c36_hit_already = 0
c37_hit_already = 0
c38_hit_already = 0
c39_hit_already = 0
c40_hit_already = 0
c41_hit_already = 0
c42_hit_already = 0
c43_hit_already = 0
c44_hit_already = 0
c45_hit_already = 0
c46_hit_already = 0
c47_hit_already = 0
c48_hit_already = 0
c49_hit_already = 0
c50_hit_already = 0
c51_hit_already = 0
c52_hit_already = 0
c53_hit_already = 0
c54_hit_already = 0
c55_hit_already = 0
c56_hit_already = 0
c57_hit_already = 0
c58_hit_already = 0
c59_hit_already = 0
c60_hit_already = 0
c61_hit_already = 0
c62_hit_already = 0
c63_hit_already = 0
c64_hit_already = 0
c65_hit_already = 0
c66_hit_already = 0
c67_hit_already = 0
c68_hit_already = 0
c69_hit_already = 0
c70_hit_already = 0
c71_hit_already = 0
c72_hit_already = 0
c73_hit_already = 0
c74_hit_already = 0
c75_hit_already = 0
c76_hit_already = 0
c77_hit_already = 0
c78_hit_already = 0
c79_hit_already = 0
c80_hit_already = 0
 Dim a(79) As Integer ' Sets the maximum number to pick
    Dim b(79) As Integer  ' New Numbers
        MaxNumber = 79 ' Must equal above

    For seq = 0 To MaxNumber
        a(seq) = seq
    Next seq
             Randomize Int(CDbl((Now))) + Timer
        For MainLoop = MaxNumber To 0 Step -1
            ChosenNumber = Int(MainLoop * Rnd)
            b(MaxNumber - MainLoop) = a(ChosenNumber)
            a(ChosenNumber) = a(MainLoop)
        Next MainLoop
dealt = 0
r = 1
For r = 0 To 19 Step 1
    dealt = dealt + 1
    If dealt > 20 Then
    Exit For
        End If
        board.MousePointer = vbHourglass
                       casehit = b(r) + 1
' Debug.Print "Numbers drawn"; casehit; "#"; r
' If casehit > 80 Or casehit < 1 Then
' MsgBox "you suck", vbOKOnly, "Crazy Clicker"
' End If
Call display(casehit)
Call delay
hits_label.Caption = hits
hits_label.Refresh
Next r
If hits >= 1 Then
Call paytable(hits, boxes_checked)
End If
bottom_display.Picture = LoadResPicture("insert_coins", bitmap)
amountbet = 0
change_button.Enabled = True
If dollars < 1 Then
board.MousePointer = vbDefault
MsgBox "No More Credits." & Chr(13) & Chr(10) & "The Change Lady Has Been Called.", vbOKOnly, "Call the Change Lady."
play4_button.Enabled = False
betone.Enabled = False
options.Show
options.moremoney.Visible = True
options.moremoney.SetFocus
options.bet_total.Caption = bettotal
options.bet_total.Refresh
options.credits_won.Caption = hopperempty
options.credits_won.Refresh
options.deals_label.Caption = totaldeals
options.deals_label.Refresh
If bettotal > 0 And hopperempty > 0 Then
Call options.returnrate
End If
Call disablechecks
Exit Function
End If
betone.Picture = LoadResPicture("betone_lit", bitmap)
If dollars >= 4 Then
    play4_button.Enabled = True
    play4_button.SetFocus
    play4_button.Picture = LoadResPicture("play4_lit", bitmap)
    End If
betone.Enabled = True
If dollars < 4 Then
betone.SetFocus
End If
Call disablechecks
board.MousePointer = vbDefault
End Function
Private Function display(casehit As Integer)
'Casehit 1
If casehit = 1 And c1_checked = 0 And c1_hit_already < 1 Then
    d1.Picture = LoadResPicture("c1", bitmap)
    c1_hit_already = 1
    Call dealsound
    d1.Refresh
        ElseIf casehit = 1 And c1_checked = 1 And c1_hit_already < 1 Then
        c1.Visible = False
        h1.Picture = LoadResPicture("hit", bitmap)
        c1__hit_already = 1
        hits = hits + 1
        Call hitsound
        h1.Refresh
        End If
'Casehit 2
If casehit = 2 And c2_checked = 0 And c2_hit_already < 1 Then
    d2.Picture = LoadResPicture("c2", bitmap)
    c2_hit_already = 1
    Call dealsound
       d2.Refresh
    ElseIf casehit = 2 And c2_checked = 1 And c2_hit_already < 1 Then
        c2.Visible = False
        h2.Picture = LoadResPicture("hit", bitmap)
        c2__hit_already = 1
        hits = hits + 1
        Call hitsound
           h2.Refresh
                End If
                'Casehit 3
If casehit = 3 And c3_checked = 0 And c3_hit_already < 1 Then
    d3.Picture = LoadResPicture("c3", bitmap)
    c3_hit_already = 1
    Call dealsound
       d3.Refresh
    ElseIf casehit = 3 And c3_checked = 1 And c3_hit_already < 1 Then
        c3.Visible = False
        h3.Picture = LoadResPicture("hit", bitmap)
        c3__hit_already = 1
        hits = hits + 1
        Call hitsound
           h3.Refresh
                End If
                'Casehit 4
If casehit = 4 And c4_checked = 0 And c4_hit_already < 1 Then
    d4.Picture = LoadResPicture("c4", bitmap)
    c4_hit_already = 1
    Call dealsound
       d4.Refresh
    ElseIf casehit = 4 And c4_checked = 1 And c4_hit_already < 1 Then
        c4.Visible = False
        h4.Picture = LoadResPicture("hit", bitmap)
        c4__hit_already = 1
        hits = hits + 1
        Call hitsound
           h4.Refresh
                End If
                'Casehit 5
If casehit = 5 And c5_checked = 0 And c5_hit_already < 1 Then
    d5.Picture = LoadResPicture("c5", bitmap)
    c5_hit_already = 1
    Call dealsound
       d5.Refresh
    ElseIf casehit = 5 And c5_checked = 1 And c5_hit_already < 1 Then
        c5.Visible = False
        h5.Picture = LoadResPicture("hit", bitmap)
        c5__hit_already = 1
        hits = hits + 1
        Call hitsound
        h5.Refresh
                End If
                
                'Casehit 6
If casehit = 6 And c6_checked = 0 And c6_hit_already < 1 Then
    d6.Picture = LoadResPicture("c6", bitmap)
    c6_hit_already = 1
    Call dealsound
       d6.Refresh
    ElseIf casehit = 6 And c6_checked = 1 And c6_hit_already < 1 Then
        c6.Visible = False
        h6.Picture = LoadResPicture("hit", bitmap)
        c6__hit_already = 1
        hits = hits + 1
        Call hitsound
        h6.Refresh
                End If
                'Casehit 7
If casehit = 7 And c7_checked = 0 And c7_hit_already < 1 Then
    d7.Picture = LoadResPicture("c7", bitmap)
    c7_hit_already = 1
    Call dealsound
      d7.Refresh
    ElseIf casehit = 7 And c7_checked = 1 And c7_hit_already < 1 Then
        c7.Visible = False
        h7.Picture = LoadResPicture("hit", bitmap)
        c7__hit_already = 1
        hits = hits + 1
        Call hitsound
        h7.Refresh
                End If
                'Casehit 8
If casehit = 8 And c8_checked = 0 And c8_hit_already < 1 Then
    d8.Picture = LoadResPicture("c8", bitmap)
    c8_hit_already = 1
    Call dealsound
       d8.Refresh
    ElseIf casehit = 8 And c8_checked = 1 And c8_hit_already < 1 Then
        c8.Visible = False
        h8.Picture = LoadResPicture("hit", bitmap)
        c8__hit_already = 1
        hits = hits + 1
        Call hitsound
       h8.Refresh
                End If
                'Casehit 9
If casehit = 9 And c9_checked = 0 And c9_hit_already < 1 Then
    d9.Picture = LoadResPicture("c9", bitmap)
    c9_hit_already = 1
    Call dealsound
       d9.Refresh
    ElseIf casehit = 9 And c9_checked = 1 And c9_hit_already < 1 Then
        c9.Visible = False
        h9.Picture = LoadResPicture("hit", bitmap)
        c9__hit_already = 1
        hits = hits + 1
        Call hitsound
        h9.Refresh
                End If
                'Casehit 10
If casehit = 10 And c10_checked = 0 And c10_hit_already < 1 Then
    d10.Picture = LoadResPicture("c10", bitmap)
    c10_hit_already = 1
    Call dealsound
       d10.Refresh
    ElseIf casehit = 10 And c10_checked = 1 And c10_hit_already < 1 Then
        c10.Visible = False
        h10.Picture = LoadResPicture("hit", bitmap)
        c10__hit_already = 1
        hits = hits + 1
        Call hitsound
        h10.Refresh
                End If
                'Casehit 11
If casehit = 11 And c11_checked = 0 And c11_hit_already < 1 Then
    d11.Picture = LoadResPicture("c11", bitmap)
    c11_hit_already = 1
    Call dealsound
       d11.Refresh
    ElseIf casehit = 11 And c11_checked = 1 And c11_hit_already < 1 Then
        c11.Visible = False
        h11.Picture = LoadResPicture("hit", bitmap)
        c11__hit_already = 1
        hits = hits + 1
        Call hitsound
        h11.Refresh
                End If
                'Casehit 12
If casehit = 12 And c12_checked = 0 And c12_hit_already < 1 Then
    d12.Picture = LoadResPicture("c12", bitmap)
    c12_hit_already = 1
    Call dealsound
       d12.Refresh
    ElseIf casehit = 12 And c12_checked = 1 And c12_hit_already < 1 Then
        c12.Visible = False
        h12.Picture = LoadResPicture("hit", bitmap)
        c12__hit_already = 1
        hits = hits + 1
        Call hitsound
        h12.Refresh
                End If
                'Casehit 13
If casehit = 13 And c13_checked = 0 And c13_hit_already < 1 Then
    d13.Picture = LoadResPicture("c13", bitmap)
    c13_hit_already = 1
    Call dealsound
       d13.Refresh
    ElseIf casehit = 13 And c13_checked = 1 And c13_hit_already < 1 Then
        c13.Visible = False
        h13.Picture = LoadResPicture("hit", bitmap)
        c13__hit_already = 1
        hits = hits + 1
        Call hitsound
        h13.Refresh
                End If
                'Casehit 14
If casehit = 14 And c14_checked = 0 And c14_hit_already < 1 Then
    d14.Picture = LoadResPicture("c14", bitmap)
    c14_hit_already = 1
    Call dealsound
       d14.Refresh
    ElseIf casehit = 14 And c14_checked = 1 And c14_hit_already < 1 Then
        c14.Visible = False
        h14.Picture = LoadResPicture("hit", bitmap)
        c14__hit_already = 1
        hits = hits + 1
        Call hitsound
        h14.Refresh
                End If
                'Casehit 15
If casehit = 15 And c15_checked = 0 And c15_hit_already < 1 Then
    d15.Picture = LoadResPicture("c15", bitmap)
    c15_hit_already = 1
    Call dealsound
       d15.Refresh
    ElseIf casehit = 15 And c15_checked = 1 And c15_hit_already < 1 Then
        c15.Visible = False
        h15.Picture = LoadResPicture("hit", bitmap)
        c15__hit_already = 1
        hits = hits + 1
        Call hitsound
        h15.Refresh
                End If
                'Casehit 16
If casehit = 16 And c16_checked = 0 And c16_hit_already < 1 Then
    d16.Picture = LoadResPicture("c16", bitmap)
    c16_hit_already = 1
    Call dealsound
       d16.Refresh
    ElseIf casehit = 16 And c16_checked = 1 And c16_hit_already < 1 Then
        c16.Visible = False
        h16.Picture = LoadResPicture("hit", bitmap)
        c16__hit_already = 1
        hits = hits + 1
        Call hitsound
        h16.Refresh
                End If 'Casehit 17
If casehit = 17 And c17_checked = 0 And c17_hit_already < 1 Then
    d17.Picture = LoadResPicture("c17", bitmap)
    c17_hit_already = 1
    Call dealsound
       d17.Refresh
    ElseIf casehit = 17 And c17_checked = 1 And c17_hit_already < 1 Then
        c17.Visible = False
        h17.Picture = LoadResPicture("hit", bitmap)
        c17__hit_already = 1
        hits = hits + 1
        Call hitsound
        h17.Refresh
                End If 'Casehit 18
If casehit = 18 And c18_checked = 0 And c18_hit_already < 1 Then
    d18.Picture = LoadResPicture("c18", bitmap)
    c18_hit_already = 1
    Call dealsound
       d18.Refresh
    ElseIf casehit = 18 And c18_checked = 1 And c18_hit_already < 1 Then
        c18.Visible = False
        h18.Picture = LoadResPicture("hit", bitmap)
        c18__hit_already = 1
        hits = hits + 1
        Call hitsound
        h18.Refresh
                End If 'Casehit 19
If casehit = 19 And c19_checked = 0 And c19_hit_already < 1 Then
    d19.Picture = LoadResPicture("c19", bitmap)
    c19_hit_already = 1
    Call dealsound
       d19.Refresh
    ElseIf casehit = 19 And c19_checked = 1 And c19_hit_already < 1 Then
        c19.Visible = False
        h19.Picture = LoadResPicture("hit", bitmap)
        c19__hit_already = 1
        hits = hits + 1
        Call hitsound
        h19.Refresh
                End If 'Casehit 20
If casehit = 20 And c20_checked = 0 And c20_hit_already < 1 Then
    d20.Picture = LoadResPicture("c20", bitmap)
    c20_hit_already = 1
    Call dealsound
       d20.Refresh
    ElseIf casehit = 20 And c20_checked = 1 And c20_hit_already < 1 Then
        c20.Visible = False
        h20.Picture = LoadResPicture("hit", bitmap)
        c20__hit_already = 1
        hits = hits + 1
        Call hitsound
        h20.Refresh
                End If
                'Casehit 21
If casehit = 21 And c21_checked = 0 And c21_hit_already < 1 Then
    d21.Picture = LoadResPicture("c21", bitmap)
    c21_hit_already = 1
    Call dealsound
       d21.Refresh
    ElseIf casehit = 21 And c21_checked = 1 And c21_hit_already < 1 Then
        c21.Visible = False
        h21.Picture = LoadResPicture("hit", bitmap)
        c21__hit_already = 1
        hits = hits + 1
        Call hitsound
        h21.Refresh
                End If
                'Casehit 22
If casehit = 22 And c22_checked = 0 And c22_hit_already < 1 Then
    d22.Picture = LoadResPicture("c22", bitmap)
    c22_hit_already = 1
    Call dealsound
       d22.Refresh
    ElseIf casehit = 22 And c22_checked = 1 And c22_hit_already < 1 Then
        c22.Visible = False
        h22.Picture = LoadResPicture("hit", bitmap)
        c22__hit_already = 1
        hits = hits + 1
        Call hitsound
        h22.Refresh
                End If 'Casehit 23
If casehit = 23 And c23_checked = 0 And c23_hit_already < 1 Then
    d23.Picture = LoadResPicture("c23", bitmap)
    c23_hit_already = 1
    Call dealsound
       d23.Refresh
    ElseIf casehit = 23 And c23_checked = 1 And c23_hit_already < 1 Then
        c23.Visible = False
        h23.Picture = LoadResPicture("hit", bitmap)
        c23__hit_already = 1
        hits = hits + 1
        Call hitsound
        h23.Refresh
                End If
                'Casehit 24
If casehit = 24 And c24_checked = 0 And c24_hit_already < 1 Then
    d24.Picture = LoadResPicture("c24", bitmap)
    c24_hit_already = 1
    Call dealsound
       d24.Refresh
    ElseIf casehit = 24 And c24_checked = 1 And c24_hit_already < 1 Then
        c24.Visible = False
        h24.Picture = LoadResPicture("hit", bitmap)
        c24__hit_already = 1
        hits = hits + 1
        Call hitsound
        h24.Refresh
                End If
                'Casehit 25
If casehit = 25 And c25_checked = 0 And c25_hit_already < 1 Then
    d25.Picture = LoadResPicture("c25", bitmap)
    c25_hit_already = 1
    Call dealsound
       d25.Refresh
    ElseIf casehit = 25 And c25_checked = 1 And c25_hit_already < 1 Then
        c25.Visible = False
        h25.Picture = LoadResPicture("hit", bitmap)
        c25__hit_already = 1
        hits = hits + 1
        Call hitsound
        h25.Refresh
                End If
                'Casehit 26
If casehit = 26 And c26_checked = 0 And c26_hit_already < 1 Then
    d26.Picture = LoadResPicture("c26", bitmap)
    c26_hit_already = 1
    Call dealsound
       d26.Refresh
    ElseIf casehit = 26 And c26_checked = 1 And c26_hit_already < 1 Then
        c26.Visible = False
        h26.Picture = LoadResPicture("hit", bitmap)
        c26__hit_already = 1
        hits = hits + 1
        Call hitsound
        h26.Refresh
                End If
                'Casehit 27
If casehit = 27 And c27_checked = 0 And c27_hit_already < 1 Then
    d27.Picture = LoadResPicture("c27", bitmap)
    c27_hit_already = 1
    Call dealsound
       d27.Refresh
    ElseIf casehit = 27 And c27_checked = 1 And c27_hit_already < 1 Then
        c27.Visible = False
        h27.Picture = LoadResPicture("hit", bitmap)
        c27__hit_already = 1
        hits = hits + 1
        Call hitsound
        h27.Refresh
                End If
                'Casehit 28
If casehit = 28 And c28_checked = 0 And c28_hit_already < 1 Then
    d28.Picture = LoadResPicture("c28", bitmap)
    c28_hit_already = 1
    Call dealsound
       d28.Refresh
    ElseIf casehit = 28 And c28_checked = 1 And c28_hit_already < 1 Then
        c28.Visible = False
        h28.Picture = LoadResPicture("hit", bitmap)
        c28__hit_already = 1
        hits = hits + 1
        Call hitsound
        h28.Refresh
                End If
                'Casehit 29
If casehit = 29 And c29_checked = 0 And c29_hit_already < 1 Then
    d29.Picture = LoadResPicture("c29", bitmap)
    c29_hit_already = 1
    Call dealsound
       d29.Refresh
    ElseIf casehit = 29 And c29_checked = 1 And c29_hit_already < 1 Then
        c29.Visible = False
        h29.Picture = LoadResPicture("hit", bitmap)
        c29__hit_already = 1
        hits = hits + 1
        Call hitsound
        h29.Refresh
                End If
                'Casehit 30
If casehit = 30 And c30_checked = 0 And c30_hit_already < 1 Then
    d30.Picture = LoadResPicture("c30", bitmap)
    c30_hit_already = 1
    Call dealsound
      d30.Refresh
    ElseIf casehit = 30 And c30_checked = 1 And c30_hit_already < 1 Then
        c30.Visible = False
        h30.Picture = LoadResPicture("hit", bitmap)
        c30__hit_already = 1
        hits = hits + 1
        Call hitsound
        h30.Refresh
                End If
                'Casehit 31
If casehit = 31 And c31_checked = 0 And c31_hit_already < 1 Then
    d31.Picture = LoadResPicture("c31", bitmap)
    c31_hit_already = 1
    Call dealsound
       d31.Refresh
    ElseIf casehit = 31 And c31_checked = 1 And c31_hit_already < 1 Then
        c31.Visible = False
        h31.Picture = LoadResPicture("hit", bitmap)
        c31__hit_already = 1
        hits = hits + 1
        Call hitsound
        h31.Refresh
                End If
                'Casehit 32
If casehit = 32 And c32_checked = 0 And c32_hit_already < 1 Then
    d32.Picture = LoadResPicture("c32", bitmap)
    c32_hit_already = 1
    Call dealsound
       d32.Refresh
    ElseIf casehit = 32 And c32_checked = 1 And c32_hit_already < 1 Then
        c32.Visible = False
        h32.Picture = LoadResPicture("hit", bitmap)
        c32__hit_already = 1
        hits = hits + 1
        Call hitsound
        h32.Refresh
                End If
                'Casehit 33
If casehit = 33 And c33_checked = 0 And c33_hit_already < 1 Then
    d33.Picture = LoadResPicture("c33", bitmap)
    c33_hit_already = 1
    Call dealsound
       d33.Refresh
    ElseIf casehit = 33 And c33_checked = 1 And c33_hit_already < 1 Then
        c33.Visible = False
        h33.Picture = LoadResPicture("hit", bitmap)
        c33__hit_already = 1
        hits = hits + 1
        Call hitsound
        h33.Refresh
                End If
                'Casehit 34
If casehit = 34 And c34_checked = 0 And c34_hit_already < 1 Then
    d34.Picture = LoadResPicture("c34", bitmap)
    c34_hit_already = 1
    Call dealsound
       d34.Refresh
    ElseIf casehit = 34 And c34_checked = 1 And c34_hit_already < 1 Then
        c34.Visible = False
        h34.Picture = LoadResPicture("hit", bitmap)
        c34__hit_already = 1
        hits = hits + 1
        Call hitsound
        h34.Refresh
                End If
                'Casehit 35
If casehit = 35 And c35_checked = 0 And c35_hit_already < 1 Then
    d35.Picture = LoadResPicture("c35", bitmap)
    c35_hit_already = 1
    Call dealsound
       d35.Refresh
    ElseIf casehit = 35 And c35_checked = 1 And c35_hit_already < 1 Then
        c35.Visible = False
        h35.Picture = LoadResPicture("hit", bitmap)
        c35__hit_already = 1
        hits = hits + 1
        Call hitsound
        h35.Refresh
                End If
                'Casehit 36
If casehit = 36 And c36_checked = 0 And c36_hit_already < 1 Then
    d36.Picture = LoadResPicture("c36", bitmap)
    c36_hit_already = 1
    Call dealsound
       d36.Refresh
    ElseIf casehit = 36 And c36_checked = 1 And c36_hit_already < 1 Then
        c36.Visible = False
        h36.Picture = LoadResPicture("hit", bitmap)
        c36__hit_already = 1
        hits = hits + 1
        Call hitsound
        h36.Refresh
                End If
                'Casehit 37
If casehit = 37 And c37_checked = 0 And c37_hit_already < 1 Then
    d37.Picture = LoadResPicture("c37", bitmap)
    c37_hit_already = 1
    Call dealsound
       d37.Refresh
    ElseIf casehit = 37 And c37_checked = 1 And c37_hit_already < 1 Then
        c37.Visible = False
        h37.Picture = LoadResPicture("hit", bitmap)
        c37__hit_already = 1
        hits = hits + 1
        Call hitsound
        h37.Refresh
                End If
                'Casehit 38
If casehit = 38 And c38_checked = 0 And c38_hit_already < 1 Then
    d38.Picture = LoadResPicture("c38", bitmap)
    c38_hit_already = 1
    Call dealsound
       d38.Refresh
    ElseIf casehit = 38 And c38_checked = 1 And c38_hit_already < 1 Then
        c38.Visible = False
        h38.Picture = LoadResPicture("hit", bitmap)
        c38__hit_already = 1
        hits = hits + 1
        Call hitsound
        h38.Refresh
                End If
                'Casehit 39
If casehit = 39 And c39_checked = 0 And c39_hit_already < 1 Then
    d39.Picture = LoadResPicture("c39", bitmap)
    c39_hit_already = 1
    Call dealsound
       d39.Refresh
    ElseIf casehit = 39 And c39_checked = 1 And c39_hit_already < 1 Then
        c39.Visible = False
        h39.Picture = LoadResPicture("hit", bitmap)
        c39__hit_already = 1
        hits = hits + 1
        Call hitsound
        h39.Refresh
                End If
                'Casehit 40
If casehit = 40 And c40_checked = 0 And c40_hit_already < 1 Then
    d40.Picture = LoadResPicture("c40", bitmap)
    c40_hit_already = 1
    Call dealsound
       d40.Refresh
    ElseIf casehit = 40 And c40_checked = 1 And c40_hit_already < 1 Then
        c40.Visible = False
        h40.Picture = LoadResPicture("hit", bitmap)
        c40__hit_already = 1
        hits = hits + 1
        Call hitsound
        h40.Refresh
                End If
                'Casehit 41
If casehit = 41 And c41_checked = 0 And c41_hit_already < 1 Then
    d41.Picture = LoadResPicture("c41", bitmap)
    c41_hit_already = 1
    Call dealsound
       d41.Refresh
    ElseIf casehit = 41 And c41_checked = 1 And c41_hit_already < 1 Then
        c41.Visible = False
        h41.Picture = LoadResPicture("hit", bitmap)
        c41__hit_already = 1
        hits = hits + 1
        Call hitsound
        h41.Refresh
                End If
                'Casehit 42
If casehit = 42 And c42_checked = 0 And c42_hit_already < 1 Then
    d42.Picture = LoadResPicture("c42", bitmap)
    c42_hit_already = 1
    Call dealsound
       d42.Refresh
    ElseIf casehit = 42 And c42_checked = 1 And c42_hit_already < 1 Then
        c42.Visible = False
        h42.Picture = LoadResPicture("hit", bitmap)
        c42__hit_already = 1
        hits = hits + 1
        Call hitsound
        h42.Refresh
                End If
                'Casehit 43
If casehit = 43 And c43_checked = 0 And c43_hit_already < 1 Then
    d43.Picture = LoadResPicture("c43", bitmap)
    c43_hit_already = 1
    Call dealsound
       d43.Refresh
    ElseIf casehit = 43 And c43_checked = 1 And c43_hit_already < 1 Then
        c43.Visible = False
        h43.Picture = LoadResPicture("hit", bitmap)
        c43__hit_already = 1
        hits = hits + 1
        Call hitsound
        h43.Refresh
                End If
                'Casehit 44
If casehit = 44 And c44_checked = 0 And c44_hit_already < 1 Then
    d44.Picture = LoadResPicture("c44", bitmap)
    c44_hit_already = 1
    Call dealsound
       d44.Refresh
    ElseIf casehit = 44 And c44_checked = 1 And c44_hit_already < 1 Then
        c44.Visible = False
        h44.Picture = LoadResPicture("hit", bitmap)
        c44__hit_already = 1
        hits = hits + 1
        Call hitsound
        h44.Refresh
                End If
                'Casehit 45
If casehit = 45 And c45_checked = 0 And c45_hit_already < 1 Then
    d45.Picture = LoadResPicture("c45", bitmap)
    c45_hit_already = 1
    Call dealsound
       d45.Refresh
    ElseIf casehit = 45 And c45_checked = 1 And c45_hit_already < 1 Then
        c45.Visible = False
        h45.Picture = LoadResPicture("hit", bitmap)
        c45__hit_already = 1
        hits = hits + 1
        Call hitsound
        h45.Refresh
                End If
                'Casehit 46
If casehit = 46 And c46_checked = 0 And c46_hit_already < 1 Then
    d46.Picture = LoadResPicture("c46", bitmap)
    c46_hit_already = 1
    Call dealsound
       d46.Refresh
    ElseIf casehit = 46 And c46_checked = 1 And c46_hit_already < 1 Then
        c46.Visible = False
        h46.Picture = LoadResPicture("hit", bitmap)
        c46__hit_already = 1
        hits = hits + 1
        Call hitsound
        h46.Refresh
                End If
                'Casehit 47
If casehit = 47 And c47_checked = 0 And c47_hit_already < 1 Then
    d47.Picture = LoadResPicture("c47", bitmap)
    c47_hit_already = 1
    Call dealsound
       d47.Refresh
    ElseIf casehit = 47 And c47_checked = 1 And c47_hit_already < 1 Then
        c47.Visible = False
        h47.Picture = LoadResPicture("hit", bitmap)
        c47__hit_already = 1
        hits = hits + 1
        Call hitsound
        h47.Refresh
                End If
                  'Casehit 48
If casehit = 48 And c48_checked = 0 And c48_hit_already < 1 Then
    d48.Picture = LoadResPicture("c48", bitmap)
    c48_hit_already = 1
    Call dealsound
       d48.Refresh
    ElseIf casehit = 48 And c48_checked = 1 And c48_hit_already < 1 Then
        c48.Visible = False
        h48.Picture = LoadResPicture("hit", bitmap)
        c48__hit_already = 1
        hits = hits + 1
        Call hitsound
        h48.Refresh
                End If
                  'Casehit 49
If casehit = 49 And c49_checked = 0 And c49_hit_already < 1 Then
    d49.Picture = LoadResPicture("c49", bitmap)
    c49_hit_already = 1
    Call dealsound
       d49.Refresh
    ElseIf casehit = 49 And c49_checked = 1 And c49_hit_already < 1 Then
        c49.Visible = False
        h49.Picture = LoadResPicture("hit", bitmap)
        c49__hit_already = 1
        hits = hits + 1
        Call hitsound
        h49.Refresh
                End If
                  'Casehit 50
If casehit = 50 And c50_checked = 0 And c50_hit_already < 1 Then
    d50.Picture = LoadResPicture("c50", bitmap)
    c50_hit_already = 1
    Call dealsound
       d50.Refresh
    ElseIf casehit = 50 And c50_checked = 1 And c50_hit_already < 1 Then
        c50.Visible = False
        h50.Picture = LoadResPicture("hit", bitmap)
        c50__hit_already = 1
        hits = hits + 1
        Call hitsound
        h50.Refresh
                End If
                  'Casehit 51
If casehit = 51 And c51_checked = 0 And c51_hit_already < 1 Then
    d51.Picture = LoadResPicture("c51", bitmap)
    c51_hit_already = 1
    Call dealsound
       d51.Refresh
    ElseIf casehit = 51 And c51_checked = 1 And c51_hit_already < 1 Then
        c51.Visible = False
        h51.Picture = LoadResPicture("hit", bitmap)
        c51__hit_already = 1
        hits = hits + 1
        Call hitsound
        h51.Refresh
                End If
                  'Casehit 52
If casehit = 52 And c52_checked = 0 And c52_hit_already < 1 Then
    d52.Picture = LoadResPicture("c52", bitmap)
    c52_hit_already = 1
    Call dealsound
       d52.Refresh
    ElseIf casehit = 52 And c52_checked = 1 And c52_hit_already < 1 Then
        c52.Visible = False
        h52.Picture = LoadResPicture("hit", bitmap)
        c52__hit_already = 1
        hits = hits + 1
        Call hitsound
        h52.Refresh
                End If
                  'Casehit 53
If casehit = 53 And c53_checked = 0 And c53_hit_already < 1 Then
    d53.Picture = LoadResPicture("c53", bitmap)
    c53_hit_already = 1
    Call dealsound
       d53.Refresh
    ElseIf casehit = 53 And c53_checked = 1 And c53_hit_already < 1 Then
        c53.Visible = False
        h53.Picture = LoadResPicture("hit", bitmap)
        c53__hit_already = 1
        hits = hits + 1
        Call hitsound
        h53.Refresh
                End If
                  'Casehit 54
If casehit = 54 And c54_checked = 0 And c54_hit_already < 1 Then
    d54.Picture = LoadResPicture("c54", bitmap)
    c54_hit_already = 1
    Call dealsound
       d54.Refresh
    ElseIf casehit = 54 And c54_checked = 1 And c54_hit_already < 1 Then
        c54.Visible = False
        h54.Picture = LoadResPicture("hit", bitmap)
        c54__hit_already = 1
        hits = hits + 1
        Call hitsound
        h54.Refresh
                End If
                  'Casehit 55
If casehit = 55 And c55_checked = 0 And c55_hit_already < 1 Then
    d55.Picture = LoadResPicture("c55", bitmap)
    c55_hit_already = 1
    Call dealsound
       d55.Refresh
    ElseIf casehit = 55 And c55_checked = 1 And c55_hit_already < 1 Then
        c55.Visible = False
        h55.Picture = LoadResPicture("hit", bitmap)
        c55__hit_already = 1
        hits = hits + 1
        Call hitsound
        h55.Refresh
                End If
                  'Casehit 56
If casehit = 56 And c56_checked = 0 And c56_hit_already < 1 Then
    d56.Picture = LoadResPicture("c56", bitmap)
    c56_hit_already = 1
    Call dealsound
       d56.Refresh
    ElseIf casehit = 56 And c56_checked = 1 And c56_hit_already < 1 Then
        c56.Visible = False
        h56.Picture = LoadResPicture("hit", bitmap)
        c56__hit_already = 1
        hits = hits + 1
        Call hitsound
        h56.Refresh
                End If
                  'Casehit 57
If casehit = 57 And c57_checked = 0 And c57_hit_already < 1 Then
    d57.Picture = LoadResPicture("c57", bitmap)
    c57_hit_already = 1
    Call dealsound
       d57.Refresh
    ElseIf casehit = 57 And c57_checked = 1 And c57_hit_already < 1 Then
        c57.Visible = False
        h57.Picture = LoadResPicture("hit", bitmap)
        c57__hit_already = 1
        hits = hits + 1
        Call hitsound
        h57.Refresh
                End If
                  'Casehit 58
If casehit = 58 And c58_checked = 0 And c58_hit_already < 1 Then
    d58.Picture = LoadResPicture("c58", bitmap)
    c58_hit_already = 1
    Call dealsound
       d58.Refresh
    ElseIf casehit = 58 And c58_checked = 1 And c58_hit_already < 1 Then
        c58.Visible = False
        h58.Picture = LoadResPicture("hit", bitmap)
        c58__hit_already = 1
        hits = hits + 1
        Call hitsound
        h58.Refresh
                End If
                  'Casehit 59
If casehit = 59 And c59_checked = 0 And c59_hit_already < 1 Then
    d59.Picture = LoadResPicture("c59", bitmap)
    c59_hit_already = 1
    Call dealsound
       d59.Refresh
    ElseIf casehit = 59 And c59_checked = 1 And c59_hit_already < 1 Then
        c59.Visible = False
        h59.Picture = LoadResPicture("hit", bitmap)
        c59__hit_already = 1
        hits = hits + 1
        Call hitsound
        h59.Refresh
                End If
                  'Casehit 60
If casehit = 60 And c60_checked = 0 And c60_hit_already < 1 Then
    d60.Picture = LoadResPicture("c60", bitmap)
    c60_hit_already = 1
    Call dealsound
       d60.Refresh
    ElseIf casehit = 60 And c60_checked = 1 And c60_hit_already < 1 Then
        c60.Visible = False
        h60.Picture = LoadResPicture("hit", bitmap)
        c60__hit_already = 1
        hits = hits + 1
        Call hitsound
        h60.Refresh
                End If
                  'Casehit 61
If casehit = 61 And c61_checked = 0 And c61_hit_already < 1 Then
    d61.Picture = LoadResPicture("c61", bitmap)
    c61_hit_already = 1
    Call dealsound
       d61.Refresh
    ElseIf casehit = 61 And c61_checked = 1 And c61_hit_already < 1 Then
        c61.Visible = False
        h61.Picture = LoadResPicture("hit", bitmap)
        c61__hit_already = 1
        hits = hits + 1
        Call hitsound
        h61.Refresh
                End If
                  'Casehit 62
If casehit = 62 And c62_checked = 0 And c62_hit_already < 1 Then
    d62.Picture = LoadResPicture("c62", bitmap)
    c62_hit_already = 1
    Call dealsound
       d62.Refresh
    ElseIf casehit = 62 And c62_checked = 1 And c62_hit_already < 1 Then
        c62.Visible = False
        h62.Picture = LoadResPicture("hit", bitmap)
        c62__hit_already = 1
        hits = hits + 1
        Call hitsound
        h62.Refresh
                End If
                  'Casehit 63
If casehit = 63 And c63_checked = 0 And c63_hit_already < 1 Then
    d63.Picture = LoadResPicture("c63", bitmap)
    c63_hit_already = 1
    Call dealsound
       d63.Refresh
    ElseIf casehit = 63 And c63_checked = 1 And c63_hit_already < 1 Then
        c63.Visible = False
        h63.Picture = LoadResPicture("hit", bitmap)
        c63__hit_already = 1
        hits = hits + 1
        Call hitsound
        h63.Refresh
                End If
                  'Casehit 64
If casehit = 64 And c64_checked = 0 And c64_hit_already < 1 Then
    d64.Picture = LoadResPicture("c64", bitmap)
    c64_hit_already = 1
    Call dealsound
       d64.Refresh
    ElseIf casehit = 64 And c64_checked = 1 And c64_hit_already < 1 Then
        c64.Visible = False
        h64.Picture = LoadResPicture("hit", bitmap)
        c64__hit_already = 1
        hits = hits + 1
        Call hitsound
        h64.Refresh
                End If
                  'Casehit 65
If casehit = 65 And c65_checked = 0 And c65_hit_already < 1 Then
    d65.Picture = LoadResPicture("c65", bitmap)
    c65_hit_already = 1
    Call dealsound
       d65.Refresh
    ElseIf casehit = 65 And c65_checked = 1 And c65_hit_already < 1 Then
        c65.Visible = False
        h65.Picture = LoadResPicture("hit", bitmap)
        c65__hit_already = 1
        hits = hits + 1
        Call hitsound
        h65.Refresh
                End If
                  'Casehit 66
If casehit = 66 And c66_checked = 0 And c66_hit_already < 1 Then
    d66.Picture = LoadResPicture("c66", bitmap)
    c66_hit_already = 1
    Call dealsound
       d66.Refresh
    ElseIf casehit = 66 And c66_checked = 1 And c66_hit_already < 1 Then
        c66.Visible = False
        h66.Picture = LoadResPicture("hit", bitmap)
        c66__hit_already = 1
        hits = hits + 1
        Call hitsound
        h66.Refresh
                End If
                  'Casehit 67
If casehit = 67 And c67_checked = 0 And c67_hit_already < 1 Then
    d67.Picture = LoadResPicture("c67", bitmap)
    c67_hit_already = 1
    Call dealsound
       d67.Refresh
    ElseIf casehit = 67 And c67_checked = 1 And c67_hit_already < 1 Then
        c67.Visible = False
        h67.Picture = LoadResPicture("hit", bitmap)
        c67__hit_already = 1
        hits = hits + 1
        Call hitsound
        h67.Refresh
                End If
                  'Casehit 68
If casehit = 68 And c68_checked = 0 And c68_hit_already < 1 Then
    d68.Picture = LoadResPicture("c68", bitmap)
    c68_hit_already = 1
    Call dealsound
       d68.Refresh
    ElseIf casehit = 68 And c68_checked = 1 And c68_hit_already < 1 Then
        c68.Visible = False
        h68.Picture = LoadResPicture("hit", bitmap)
        c68__hit_already = 1
        hits = hits + 1
        Call hitsound
        h68.Refresh
                End If
                  'Casehit 69
If casehit = 69 And c69_checked = 0 And c69_hit_already < 1 Then
    d69.Picture = LoadResPicture("c69", bitmap)
    c69_hit_already = 1
    Call dealsound
       d69.Refresh
    ElseIf casehit = 69 And c69_checked = 1 And c69_hit_already < 1 Then
        c69.Visible = False
        h69.Picture = LoadResPicture("hit", bitmap)
        c69__hit_already = 1
        hits = hits + 1
        Call hitsound
        h69.Refresh
                End If
                  'Casehit 70
If casehit = 70 And c70_checked = 0 And c70_hit_already < 1 Then
    d70.Picture = LoadResPicture("c70", bitmap)
    c70_hit_already = 1
    Call dealsound
       d70.Refresh
    ElseIf casehit = 70 And c70_checked = 1 And c70_hit_already < 1 Then
        c70.Visible = False
        h70.Picture = LoadResPicture("hit", bitmap)
        c70__hit_already = 1
        hits = hits + 1
        Call hitsound
        h70.Refresh
                End If  'Casehit 71
If casehit = 71 And c71_checked = 0 And c71_hit_already < 1 Then
    d71.Picture = LoadResPicture("c71", bitmap)
    c71_hit_already = 1
    Call dealsound
       d71.Refresh
    ElseIf casehit = 71 And c71_checked = 1 And c71_hit_already < 1 Then
        c71.Visible = False
        h71.Picture = LoadResPicture("hit", bitmap)
        c71__hit_already = 1
        hits = hits + 1
        Call hitsound
        h71.Refresh
                End If
                  'Casehit 72
If casehit = 72 And c72_checked = 0 And c72_hit_already < 1 Then
    d72.Picture = LoadResPicture("c72", bitmap)
    c72_hit_already = 1
    Call dealsound
       d72.Refresh
    ElseIf casehit = 72 And c72_checked = 1 And c72_hit_already < 1 Then
        c72.Visible = False
        h72.Picture = LoadResPicture("hit", bitmap)
        c72__hit_already = 1
        hits = hits + 1
        Call hitsound
        h72.Refresh
                End If
                  'Casehit 73
If casehit = 73 And c73_checked = 0 And c73_hit_already < 1 Then
    d73.Picture = LoadResPicture("c73", bitmap)
    c73_hit_already = 1
    Call dealsound
       d73.Refresh
    ElseIf casehit = 73 And c73_checked = 1 And c73_hit_already < 1 Then
        c73.Visible = False
        h73.Picture = LoadResPicture("hit", bitmap)
        c73__hit_already = 1
        hits = hits + 1
        Call hitsound
        h73.Refresh
                End If
                  'Casehit 74
If casehit = 74 And c74_checked = 0 And c74_hit_already < 1 Then
    d74.Picture = LoadResPicture("c74", bitmap)
    c74_hit_already = 1
    Call dealsound
       d74.Refresh
    ElseIf casehit = 74 And c74_checked = 1 And c74_hit_already < 1 Then
        c74.Visible = False
        h74.Picture = LoadResPicture("hit", bitmap)
        c74__hit_already = 1
        hits = hits + 1
        Call hitsound
        h74.Refresh
                End If
                  'Casehit 75
If casehit = 75 And c75_checked = 0 And c75_hit_already < 1 Then
    d75.Picture = LoadResPicture("c75", bitmap)
    c75_hit_already = 1
    Call dealsound
       d75.Refresh
    ElseIf casehit = 75 And c75_checked = 1 And c75_hit_already < 1 Then
        c75.Visible = False
        h75.Picture = LoadResPicture("hit", bitmap)
        c75__hit_already = 1
        hits = hits + 1
        Call hitsound
        h75.Refresh
                End If
                  'Casehit 76
If casehit = 76 And c76_checked = 0 And c76_hit_already < 1 Then
    d76.Picture = LoadResPicture("c76", bitmap)
    c76_hit_already = 1
    Call dealsound
       d76.Refresh
    ElseIf casehit = 76 And c76_checked = 1 And c76_hit_already < 1 Then
        c76.Visible = False
        h76.Picture = LoadResPicture("hit", bitmap)
        c76__hit_already = 1
        hits = hits + 1
        Call hitsound
        h76.Refresh
                End If
                  'Casehit 77
If casehit = 77 And c77_checked = 0 And c77_hit_already < 1 Then
    d77.Picture = LoadResPicture("c77", bitmap)
    c77_hit_already = 1
    Call dealsound
       d77.Refresh
    ElseIf casehit = 77 And c77_checked = 1 And c77_hit_already < 1 Then
        c77.Visible = False
        h77.Picture = LoadResPicture("hit", bitmap)
        c77__hit_already = 1
        hits = hits + 1
        Call hitsound
        h77.Refresh
                End If
                  'Casehit 78
If casehit = 78 And c78_checked = 0 And c78_hit_already < 1 Then
    d78.Picture = LoadResPicture("c78", bitmap)
    c78_hit_already = 1
    Call dealsound
       d78.Refresh
    ElseIf casehit = 78 And c78_checked = 1 And c78_hit_already < 1 Then
        c78.Visible = False
        h78.Picture = LoadResPicture("hit", bitmap)
        c78__hit_already = 1
        hits = hits + 1
        Call hitsound
        h78.Refresh
                End If
                  'Casehit 79
If casehit = 79 And c79_checked = 0 And c79_hit_already < 1 Then
    d79.Picture = LoadResPicture("c79", bitmap)
    c79_hit_already = 1
    Call dealsound
       d79.Refresh
    ElseIf casehit = 79 And c79_checked = 1 And c79_hit_already < 1 Then
        c79.Visible = False
        h79.Picture = LoadResPicture("hit", bitmap)
        c79__hit_already = 1
        hits = hits + 1
        Call hitsound
        h79.Refresh
                End If
                  'Casehit 80
If casehit = 80 And c80_checked = 0 And c80_hit_already < 1 Then
    d80.Picture = LoadResPicture("c80", bitmap)
    c80_hit_already = 1
    Call dealsound
       d80.Refresh
    ElseIf casehit = 80 And c80_checked = 1 And c80_hit_already < 1 Then
        c80.Visible = False
        h80.Picture = LoadResPicture("hit", bitmap)
        c80__hit_already = 1
        hits = hits + 1
        Call hitsound
        h80.Refresh
                End If
End Function
Private Function clearall()
top_display.Picture = LoadPicture
paid_label.Caption = ""
hits = 0
d1.Picture = LoadPicture
c1.Visible = True
d2.Picture = LoadPicture
c2.Visible = True
d3.Picture = LoadPicture
c3.Visible = True
d4.Picture = LoadPicture
c4.Visible = True
d5.Picture = LoadPicture
c5.Visible = True
d6.Picture = LoadPicture
c6.Visible = True
d7.Picture = LoadPicture
c7.Visible = True
d8.Picture = LoadPicture
c8.Visible = True
d9.Picture = LoadPicture
c9.Visible = True
d10.Picture = LoadPicture
c10.Visible = True
d11.Picture = LoadPicture
c11.Visible = True
d12.Picture = LoadPicture
c12.Visible = True
d13.Picture = LoadPicture
c13.Visible = True
d14.Picture = LoadPicture
c14.Visible = True
d15.Picture = LoadPicture
c15.Visible = True
d16.Picture = LoadPicture
c16.Visible = True
d17.Picture = LoadPicture
c17.Visible = True
d18.Picture = LoadPicture
c18.Visible = True
d19.Picture = LoadPicture
c19.Visible = True
d20.Picture = LoadPicture
c20.Visible = True
d21.Picture = LoadPicture
c21.Visible = True
d22.Picture = LoadPicture
c22.Visible = True
d23.Picture = LoadPicture
c23.Visible = True
d24.Picture = LoadPicture
c24.Visible = True
d25.Picture = LoadPicture
c25.Visible = True
d26.Picture = LoadPicture
c26.Visible = True
d27.Picture = LoadPicture
c27.Visible = True
d28.Picture = LoadPicture
c28.Visible = True
d29.Picture = LoadPicture
c29.Visible = True
d30.Picture = LoadPicture
c30.Visible = True
d31.Picture = LoadPicture
c31.Visible = True
d32.Picture = LoadPicture
c32.Visible = True
d33.Picture = LoadPicture
c33.Visible = True
d34.Picture = LoadPicture
c34.Visible = True
d35.Picture = LoadPicture
c35.Visible = True
d36.Picture = LoadPicture
c36.Visible = True
d37.Picture = LoadPicture
c37.Visible = True
d38.Picture = LoadPicture
c38.Visible = True
d39.Picture = LoadPicture
c39.Visible = True
d40.Picture = LoadPicture
c40.Visible = True
d41.Picture = LoadPicture
c41.Visible = True
d42.Picture = LoadPicture
c42.Visible = True
d43.Picture = LoadPicture
c43.Visible = True
d44.Picture = LoadPicture
c44.Visible = True
d45.Picture = LoadPicture
c45.Visible = True
d46.Picture = LoadPicture
c46.Visible = True
d47.Picture = LoadPicture
c47.Visible = True
d48.Picture = LoadPicture
c48.Visible = True
d49.Picture = LoadPicture
c49.Visible = True
d50.Picture = LoadPicture
c50.Visible = True
d51.Picture = LoadPicture
c51.Visible = True
d52.Picture = LoadPicture
c52.Visible = True
d53.Picture = LoadPicture
c53.Visible = True
d54.Picture = LoadPicture
c54.Visible = True
d55.Picture = LoadPicture
c55.Visible = True
d56.Picture = LoadPicture
c56.Visible = True
d57.Picture = LoadPicture
c57.Visible = True
d58.Picture = LoadPicture
c58.Visible = True
d59.Picture = LoadPicture
c59.Visible = True
d60.Picture = LoadPicture
c60.Visible = True
d61.Picture = LoadPicture
c61.Visible = True
d62.Picture = LoadPicture
c62.Visible = True
d63.Picture = LoadPicture
c63.Visible = True
d64.Picture = LoadPicture
c64.Visible = True
d65.Picture = LoadPicture
c65.Visible = True
d66.Picture = LoadPicture
c66.Visible = True
d67.Picture = LoadPicture
c67.Visible = True
d68.Picture = LoadPicture
c68.Visible = True
d69.Picture = LoadPicture
c69.Visible = True
d70.Picture = LoadPicture
c70.Visible = True
d71.Picture = LoadPicture
c71.Visible = True
d72.Picture = LoadPicture
c72.Visible = True
d73.Picture = LoadPicture
c73.Visible = True
d74.Picture = LoadPicture
c74.Visible = True
d75.Picture = LoadPicture
c75.Visible = True
d76.Picture = LoadPicture
c76.Visible = True
d77.Picture = LoadPicture
c77.Visible = True
d78.Picture = LoadPicture
c78.Visible = True
d79.Picture = LoadPicture
c79.Visible = True
d80.Picture = LoadPicture
c80.Visible = True
End Function
Public Function delay()
Dim lngEnd As Long, lngNow As Long
    lngEnd = GetTickCount()
    lngEnd = GetTickCount + delaytime
    Do
        DoEvents
            lngNow = GetTickCount()
        Loop Until lngNow >= lngEnd

    'Dim g As Double
    'g = 0
    'For g = 1 To delaytime * 80000
    'Next g
End Function
Public Function delay2(z As Long)
Dim lngEnd As Long, lngNow As Long
    lngEnd = GetTickCount()
    lngEnd = GetTickCount + z
    Do
        DoEvents
            lngNow = GetTickCount()
        Loop Until lngNow >= lngEnd

End Function



Public Function clearchecks()
c1_checked = 0
c1.Picture = LoadPicture
d1.Picture = LoadPicture
h1.Picture = LoadPicture
c2_checked = 0
c2.Picture = LoadPicture
d2.Picture = LoadPicture
h2.Picture = LoadPicture
c3_checked = 0
c3.Picture = LoadPicture
d3.Picture = LoadPicture
h3.Picture = LoadPicture
c4_checked = 0
c4.Picture = LoadPicture
d4.Picture = LoadPicture
h4.Picture = LoadPicture
c5_checked = 0
c5.Picture = LoadPicture
d5.Picture = LoadPicture
h5.Picture = LoadPicture
c6_checked = 0
c6.Picture = LoadPicture
d6.Picture = LoadPicture
h6.Picture = LoadPicture
c7_checked = 0
c7.Picture = LoadPicture
d7.Picture = LoadPicture
h7.Picture = LoadPicture
c8_checked = 0
c8.Picture = LoadPicture
d8.Picture = LoadPicture
h8.Picture = LoadPicture
c9_checked = 0
c9.Picture = LoadPicture
d9.Picture = LoadPicture
h9.Picture = LoadPicture
c10_checked = 0
c10.Picture = LoadPicture
d10.Picture = LoadPicture
h10.Picture = LoadPicture
c11_checked = 0
c11.Picture = LoadPicture
d11.Picture = LoadPicture
h11.Picture = LoadPicture
c12_checked = 0
c12.Picture = LoadPicture
d12.Picture = LoadPicture
h12.Picture = LoadPicture
c13_checked = 0
c13.Picture = LoadPicture
d13.Picture = LoadPicture
h13.Picture = LoadPicture
c14_checked = 0
c14.Picture = LoadPicture
d14.Picture = LoadPicture
h14.Picture = LoadPicture
c15_checked = 0
c15.Picture = LoadPicture
d15.Picture = LoadPicture
h15.Picture = LoadPicture
c16_checked = 0
c16.Picture = LoadPicture
d16.Picture = LoadPicture
h16.Picture = LoadPicture
c17_checked = 0
c17.Picture = LoadPicture
d17.Picture = LoadPicture
h17.Picture = LoadPicture
c18_checked = 0
c18.Picture = LoadPicture
d18.Picture = LoadPicture
h18.Picture = LoadPicture
c19_checked = 0
c19.Picture = LoadPicture
d19.Picture = LoadPicture
h19.Picture = LoadPicture
c20_checked = 0
c20.Picture = LoadPicture
d20.Picture = LoadPicture
h20.Picture = LoadPicture
c21_checked = 0
c21.Picture = LoadPicture
d21.Picture = LoadPicture
h21.Picture = LoadPicture
c22_checked = 0
c22.Picture = LoadPicture
d22.Picture = LoadPicture
h22.Picture = LoadPicture
c23_checked = 0
c23.Picture = LoadPicture
d23.Picture = LoadPicture
h23.Picture = LoadPicture
c24_checked = 0
c24.Picture = LoadPicture
d24.Picture = LoadPicture
h24.Picture = LoadPicture
c25_checked = 0
c25.Picture = LoadPicture
d25.Picture = LoadPicture
h25.Picture = LoadPicture
c26_checked = 0
c26.Picture = LoadPicture
d26.Picture = LoadPicture
h26.Picture = LoadPicture
c27_checked = 0
c27.Picture = LoadPicture
d27.Picture = LoadPicture
h27.Picture = LoadPicture
c28_checked = 0
c28.Picture = LoadPicture
d28.Picture = LoadPicture
h28.Picture = LoadPicture
c29_checked = 0
c29.Picture = LoadPicture
d29.Picture = LoadPicture
h29.Picture = LoadPicture
c30_checked = 0
c30.Picture = LoadPicture
d30.Picture = LoadPicture
h30.Picture = LoadPicture
c31_checked = 0
c31.Picture = LoadPicture
d31.Picture = LoadPicture
h31.Picture = LoadPicture
c32_checked = 0
c32.Picture = LoadPicture
d32.Picture = LoadPicture
h32.Picture = LoadPicture
c33_checked = 0
c33.Picture = LoadPicture
d33.Picture = LoadPicture
h33.Picture = LoadPicture
c34_checked = 0
c34.Picture = LoadPicture
d34.Picture = LoadPicture
h34.Picture = LoadPicture
c35_checked = 0
c35.Picture = LoadPicture
d35.Picture = LoadPicture
h35.Picture = LoadPicture
c36_checked = 0
c36.Picture = LoadPicture
d36.Picture = LoadPicture
h36.Picture = LoadPicture
c37_checked = 0
c37.Picture = LoadPicture
d37.Picture = LoadPicture
h37.Picture = LoadPicture
c38_checked = 0
c38.Picture = LoadPicture
d38.Picture = LoadPicture
h38.Picture = LoadPicture
c39_checked = 0
c39.Picture = LoadPicture
d39.Picture = LoadPicture
h39.Picture = LoadPicture
c40_checked = 0
c40.Picture = LoadPicture
d40.Picture = LoadPicture
h40.Picture = LoadPicture
c41_checked = 0
c41.Picture = LoadPicture
d41.Picture = LoadPicture
h41.Picture = LoadPicture
c42_checked = 0
c42.Picture = LoadPicture
d42.Picture = LoadPicture
h42.Picture = LoadPicture
c43_checked = 0
c43.Picture = LoadPicture
d43.Picture = LoadPicture
h43.Picture = LoadPicture
c44_checked = 0
c44.Picture = LoadPicture
d44.Picture = LoadPicture
h44.Picture = LoadPicture
c45_checked = 0
c45.Picture = LoadPicture
d45.Picture = LoadPicture
h45.Picture = LoadPicture
c46_checked = 0
c46.Picture = LoadPicture
d46.Picture = LoadPicture
h46.Picture = LoadPicture
c47_checked = 0
c47.Picture = LoadPicture
d47.Picture = LoadPicture
h47.Picture = LoadPicture
c48_checked = 0
c48.Picture = LoadPicture
d48.Picture = LoadPicture
h48.Picture = LoadPicture
c49_checked = 0
c49.Picture = LoadPicture
d49.Picture = LoadPicture
h49.Picture = LoadPicture
c50_checked = 0
c50.Picture = LoadPicture
d50.Picture = LoadPicture
h50.Picture = LoadPicture
c51_checked = 0
c51.Picture = LoadPicture
d51.Picture = LoadPicture
h51.Picture = LoadPicture
c52_checked = 0
c52.Picture = LoadPicture
d52.Picture = LoadPicture
h52.Picture = LoadPicture
c53_checked = 0
c53.Picture = LoadPicture
d53.Picture = LoadPicture
h53.Picture = LoadPicture
c54_checked = 0
c54.Picture = LoadPicture
d54.Picture = LoadPicture
h54.Picture = LoadPicture
c55_checked = 0
c55.Picture = LoadPicture
d55.Picture = LoadPicture
h55.Picture = LoadPicture
c56_checked = 0
c56.Picture = LoadPicture
d56.Picture = LoadPicture
h56.Picture = LoadPicture
c57_checked = 0
c57.Picture = LoadPicture
d57.Picture = LoadPicture
h57.Picture = LoadPicture
c58_checked = 0
c58.Picture = LoadPicture
d58.Picture = LoadPicture
h58.Picture = LoadPicture
c59_checked = 0
c59.Picture = LoadPicture
d59.Picture = LoadPicture
h59.Picture = LoadPicture
c60_checked = 0
c60.Picture = LoadPicture
d60.Picture = LoadPicture
h60.Picture = LoadPicture
c61_checked = 0
c61.Picture = LoadPicture
d61.Picture = LoadPicture
h61.Picture = LoadPicture
c62_checked = 0
c62.Picture = LoadPicture
d62.Picture = LoadPicture
h62.Picture = LoadPicture
c63_checked = 0
c63.Picture = LoadPicture
d63.Picture = LoadPicture
h63.Picture = LoadPicture
c64_checked = 0
c64.Picture = LoadPicture
d64.Picture = LoadPicture
h64.Picture = LoadPicture
c65_checked = 0
c65.Picture = LoadPicture
d65.Picture = LoadPicture
h65.Picture = LoadPicture
c66_checked = 0
c66.Picture = LoadPicture
d66.Picture = LoadPicture
h66.Picture = LoadPicture
c67_checked = 0
c67.Picture = LoadPicture
d67.Picture = LoadPicture
h67.Picture = LoadPicture
c68_checked = 0
c68.Picture = LoadPicture
d68.Picture = LoadPicture
h68.Picture = LoadPicture
c69_checked = 0
c69.Picture = LoadPicture
d69.Picture = LoadPicture
h69.Picture = LoadPicture
c70_checked = 0
c70.Picture = LoadPicture
d70.Picture = LoadPicture
h70.Picture = LoadPicture
c71_checked = 0
c71.Picture = LoadPicture
d71.Picture = LoadPicture
h71.Picture = LoadPicture
c72_checked = 0
c72.Picture = LoadPicture
d72.Picture = LoadPicture
h72.Picture = LoadPicture
c73_checked = 0
c73.Picture = LoadPicture
d73.Picture = LoadPicture
h73.Picture = LoadPicture
c74_checked = 0
c74.Picture = LoadPicture
d74.Picture = LoadPicture
h74.Picture = LoadPicture
c75_checked = 0
c75.Picture = LoadPicture
d75.Picture = LoadPicture
h75.Picture = LoadPicture
c76_checked = 0
c76.Picture = LoadPicture
d76.Picture = LoadPicture
h76.Picture = LoadPicture
c77_checked = 0
c77.Picture = LoadPicture
d77.Picture = LoadPicture
h77.Picture = LoadPicture
c78_checked = 0
c78.Picture = LoadPicture
d78.Picture = LoadPicture
h78.Picture = LoadPicture
c79_checked = 0
c79.Picture = LoadPicture
d79.Picture = LoadPicture
h79.Picture = LoadPicture
c80_checked = 0
c80.Picture = LoadPicture
d80.Picture = LoadPicture
h80.Picture = LoadPicture
boxes_checked = 0
End Function
Private Function paytable(hits As Integer, boxes_checked As Integer)
Dim pay(9, 9) As Double
pay(0, 0) = 0
pay(0, 1) = 0
pay(0, 2) = 0
pay(0, 3) = 0
pay(0, 4) = 0
pay(0, 5) = 0
pay(0, 6) = 0
pay(0, 7) = 0
pay(0, 8) = 0
pay(0, 9) = 0
pay(1, 0) = 0
pay(1, 1) = 12
pay(1, 2) = 0
pay(1, 3) = 0
pay(1, 4) = 0
pay(1, 5) = 0
pay(1, 6) = 0
pay(1, 7) = 0
pay(1, 8) = 0
pay(1, 9) = 0
pay(2, 0) = 0
pay(2, 1) = 2
pay(2, 2) = 43
pay(2, 3) = 0
pay(2, 4) = 0
pay(2, 5) = 0
pay(2, 6) = 0
pay(2, 7) = 0
pay(2, 8) = 0
pay(2, 9) = 0
pay(3, 0) = 0
pay(3, 1) = 2
pay(3, 2) = 3
pay(3, 3) = 100
pay(3, 4) = 0
pay(3, 5) = 0
pay(3, 6) = 0
pay(3, 7) = 0
pay(3, 8) = 0
pay(3, 9) = 0
pay(4, 0) = 0
pay(4, 1) = 0
pay(4, 2) = 3
pay(4, 3) = 12
pay(4, 4) = 800
pay(4, 5) = 0
pay(4, 6) = 0
pay(4, 7) = 0
pay(4, 8) = 0
pay(4, 9) = 0
pay(5, 0) = 0
pay(5, 1) = 0
pay(5, 2) = 3
pay(5, 3) = 4
pay(5, 4) = 95
pay(5, 5) = 1500
pay(5, 6) = 0
pay(5, 7) = 0
pay(5, 8) = 0
pay(5, 9) = 0
pay(6, 0) = 0
pay(6, 1) = 0
pay(6, 2) = 1
pay(6, 3) = 2
pay(6, 4) = 25
pay(6, 5) = 349
pay(6, 6) = 7560
pay(6, 7) = 0
pay(6, 8) = 0
pay(6, 9) = 0
pay(7, 0) = 0
pay(7, 1) = 0
pay(7, 2) = 0
pay(7, 3) = 1
pay(7, 4) = 12
pay(7, 5) = 112
pay(7, 6) = 1500
pay(7, 7) = 8000
pay(7, 8) = 0
pay(7, 9) = 0
pay(8, 0) = 0
pay(8, 1) = 0
pay(8, 2) = 0
pay(8, 3) = 1
pay(8, 4) = 6
pay(8, 5) = 54
pay(8, 6) = 349
pay(8, 7) = 4700
pay(8, 8) = 9000
pay(8, 9) = 0
pay(9, 0) = 0
pay(9, 1) = 0
pay(9, 2) = 0
pay(9, 3) = 0
pay(9, 4) = 5
pay(9, 5) = 28
pay(9, 6) = 140
pay(9, 7) = 1000
pay(9, 8) = 4700
pay(9, 9) = 10000
Dim new_hits As Integer
Dim new_boxes As Integer
new_hits = 0
new_boxes = 0
Dim amount_won As Double
amount_won = 0
new_hits = hits - 1
new_boxes_checked = boxes_checked - 1
amount_won = pay(new_boxes_checked, new_hits) * amountbet
If amount_won > 0 Then
top_display.Picture = LoadResPicture("winner_paid", bitmap)
Call displaywinnings(amount_won)
End If
End Function
Private Function displaywinnings(amount_won As Double)
Dim a As Double
Dim s As Double
Dim remaining As Double
hopper2 = hopper2 + amount_won
hopperempty = hopperempty + amount_won
If amount_won > 600 And autokenomode = 0 Then
    remaining = amount_won - 600
    amount_won = 600
    bottom_display.Picture = LoadResPicture("jackpot", bitmap)
    For a = 1 To amount_won Step 1
    paid_label.Caption = a
    paid_label.Refresh
    credits_label.Caption = dollars + a
    credits_label.Refresh
    Call moneysound
    Next a
    bottom_display.Picture = LoadResPicture("handpay", bitmap)
    Dim jackpotmsg
    jackpotmsg = "Your remaining " & remaining & " credits will be paid by hand"
    MsgBox jackpotmsg, vbOKOnly
    paid_label.Caption = amount_won + remaining
    paid_label.Refresh
    dollars = dollars + amount_won + remaining
    credits_label.Caption = dollars + amount_won + remaining
    credits_label.Refresh
    dollars = amount_won + remaining
    Exit Function
    End If
    
For a = 1 To amount_won Step 1
   paid_label.Caption = a
paid_label.Refresh
credits_label.Caption = dollars + a
credits_label.Refresh
Call moneysound
Next a
dollars = amount_won + dollars
If hopper2 > 10000 And autokenomode = 0 Then
    bottom_display.Picture = LoadResPicture("hopperempty", bitmap)
    MsgBox "You Emptied The Hopper", vbOKOnly
    MsgBox "Please Wait While The Hopper is Refilled", vbSystemModal, "The Attendant Has Been Called"
   Call hoppersound
      Dim u As Double
        For u = 1 To 300
        Call moneysound
          Next u
          hopper2 = 0
       End If
End Function
Private Function scoreupdate()
Score.hits_label.Caption = hits
Score.spots_marked.Caption = boxes_checked
End Function
Private Function dealsound()
  If soundonoff = 0 Then
  Dim x%
 soundName$ = "deal.wav" ' The file to play
 wFlags% = SND_ASYNC Or SND_NODEFAULT
 x% = sndPlaySound(soundName$, uFlags%)
 Else
Exit Function
End If
End Function
Private Function hitsound()
   If soundonoff = 0 Then
  Dim x%
  soundName$ = "hit.wav" ' The file to play
   wFlags% = SND_ASYNC Or SND_NODEFAULT
    x% = sndPlaySound(soundName$, uFlags%)
Else
Exit Function
End If
End Function
Private Function checksound()
 If soundonoff = 0 Then
 
  Dim x%
   soundName$ = "check.wav" ' The file to play
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    x% = sndPlaySound(soundName$, uFlags%)
Else
Exit Function
End If
End Function
Public Function moneysound()
If soundonoff = 0 Then
Dim x%
   soundName$ = "money.wav" ' The file to play
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    
    x% = sndPlaySound(soundName$, uFlags%)
Else
Call delay2(50)
Exit Function
End If
End Function
Public Function hoppersound()
If soundonoff = 0 Then
Dim x%
   soundName$ = "open.wav" ' The file to play
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    
    x% = sndPlaySound(soundName$, uFlags%)
Else
Exit Function
End If
End Function

Private Sub play4_button_Click()
Dim addtobet As Integer
Dim betcorrector As Integer
betcorrector = 0
addtobet = 0
If amountbet = 0 Then
amountbet = 4
ElseIf amountbet = 1 Then
addtobet = 3
betcorrector = 1
ElseIf amountbet = 2 Then
addtobet = 2
betcorrector = 2
ElseIf amountbet = 3 Then
addtobet = 1
betcorrector = 3
ElseIf amountbet = 4 Then
amountbet = 4
End If
If boxes_checked > 10 Or boxes_checked < 2 Then
     MsgBox "Please Choose From 2-10 Spots", vbOKOnly, "Crazy Clicker"
    Exit Sub
    End If
    coinsin_label.Caption = "4"
    dollars = dollars - ((amountbet + addtobet) - betcorrector)
credits_label.Caption = dollars
amountbet = 4
Call go
End Sub

Private Sub start_button_Click()
If boxes_checked > 10 Or boxes_checked < 2 Then
    MsgBox "Please Choose From 2-10 Spots", vbOKOnly, "Crazy Clicker"
    Exit Sub
    End If
Call go
End Sub
Private Function disablechecks()
c1.Enabled = False
c2.Enabled = False
c3.Enabled = False
c4.Enabled = False
c5.Enabled = False
c6.Enabled = False
c7.Enabled = False
c8.Enabled = False
c9.Enabled = False
c10.Enabled = False
c11.Enabled = False
c12.Enabled = False
c13.Enabled = False
c14.Enabled = False
c15.Enabled = False
c16.Enabled = False
c17.Enabled = False
c18.Enabled = False
c19.Enabled = False
c20.Enabled = False
c21.Enabled = False
c22.Enabled = False
c23.Enabled = False
c24.Enabled = False
c25.Enabled = False
c26.Enabled = False
c27.Enabled = False
c28.Enabled = False
c29.Enabled = False
c30.Enabled = False
c31.Enabled = False
c32.Enabled = False
c33.Enabled = False
c34.Enabled = False
c35.Enabled = False
c36.Enabled = False
c37.Enabled = False
c38.Enabled = False
c39.Enabled = False
c40.Enabled = False
c41.Enabled = False
c42.Enabled = False
c43.Enabled = False
c44.Enabled = False
c45.Enabled = False
c46.Enabled = False
c47.Enabled = False
c48.Enabled = False
c49.Enabled = False
c50.Enabled = False
c51.Enabled = False
c52.Enabled = False
c53.Enabled = False
c54.Enabled = False
c55.Enabled = False
c56.Enabled = False
c57.Enabled = False
c58.Enabled = False
c59.Enabled = False
c60.Enabled = False
c61.Enabled = False
c62.Enabled = False
c63.Enabled = False
c64.Enabled = False
c65.Enabled = False
c66.Enabled = False
c67.Enabled = False
c68.Enabled = False
c69.Enabled = False
c70.Enabled = False
c71.Enabled = False
c72.Enabled = False
c73.Enabled = False
c74.Enabled = False
c75.Enabled = False
c76.Enabled = False
c77.Enabled = False
c78.Enabled = False
c79.Enabled = False
c80.Enabled = False
End Function
Private Function enablechecks()
c1.Enabled = True
c2.Enabled = True
c3.Enabled = True
c4.Enabled = True
c5.Enabled = True
c6.Enabled = True
c7.Enabled = True
c8.Enabled = True
c9.Enabled = True
c10.Enabled = True
c11.Enabled = True
c12.Enabled = True
c13.Enabled = True
c14.Enabled = True
c15.Enabled = True
c16.Enabled = True
c17.Enabled = True
c18.Enabled = True
c19.Enabled = True
c20.Enabled = True
c21.Enabled = True
c22.Enabled = True
c23.Enabled = True
c24.Enabled = True
c25.Enabled = True
c26.Enabled = True
c27.Enabled = True
c28.Enabled = True
c29.Enabled = True
c30.Enabled = True
c31.Enabled = True
c32.Enabled = True
c33.Enabled = True
c34.Enabled = True
c35.Enabled = True
c36.Enabled = True
c37.Enabled = True
c38.Enabled = True
c39.Enabled = True
c40.Enabled = True
c41.Enabled = True
c42.Enabled = True
c43.Enabled = True
c44.Enabled = True
c45.Enabled = True
c46.Enabled = True
c47.Enabled = True
c48.Enabled = True
c49.Enabled = True
c50.Enabled = True
c51.Enabled = True
c52.Enabled = True
c53.Enabled = True
c54.Enabled = True
c55.Enabled = True
c56.Enabled = True
c57.Enabled = True
c58.Enabled = True
c59.Enabled = True
c60.Enabled = True
c61.Enabled = True
c62.Enabled = True
c63.Enabled = True
c64.Enabled = True
c65.Enabled = True
c66.Enabled = True
c67.Enabled = True
c68.Enabled = True
c69.Enabled = True
c70.Enabled = True
c71.Enabled = True
c72.Enabled = True
c73.Enabled = True
c74.Enabled = True
c75.Enabled = True
c76.Enabled = True
c77.Enabled = True
c78.Enabled = True
c79.Enabled = True
c80.Enabled = True
End Function
Public Function gamesettings()
If c1_checked = 1 Then
c1.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c2_checked = 1 Then
c2.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If

If c3_checked = 1 Then
c3.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c4_checked = 1 Then
c4.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c5_checked = 1 Then
c5.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c6_checked = 1 Then
c6.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c7_checked = 1 Then
c7.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c8_checked = 1 Then
c8.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c9_checked = 1 Then
c9.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c10_checked = 1 Then
c10.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c11_checked = 1 Then
c11.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c12_checked = 1 Then
c12.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c13_checked = 1 Then
c13.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c14_checked = 1 Then
c14.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c15_checked = 1 Then
c15.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c16_checked = 1 Then
c16.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c17_checked = 1 Then
c17.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c18_checked = 1 Then
c18.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c19_checked = 1 Then
c19.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c20_checked = 1 Then
c20.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c21_checked = 1 Then
c21.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c22_checked = 1 Then
c22.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c23_checked = 1 Then
c23.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c24_checked = 1 Then
c24.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c25_checked = 1 Then
c25.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c26_checked = 1 Then
c26.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c27_checked = 1 Then
c27.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c28_checked = 1 Then
c28.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c29_checked = 1 Then
c29.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c30_checked = 1 Then
c21.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c31_checked = 1 Then
c31.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c32_checked = 1 Then
c32.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c33_checked = 1 Then
c33.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c34_checked = 1 Then
c34.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c35_checked = 1 Then
c35.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c36_checked = 1 Then
c36.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c37_checked = 1 Then
c37.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c38_checked = 1 Then
c38.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c39_checked = 1 Then
c39.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c40_checked = 1 Then
c40.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If

If c41_checked = 1 Then
c41.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c42_checked = 1 Then
c42.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c43_checked = 1 Then
c43.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c44_checked = 1 Then
c44.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c45_checked = 1 Then
c45.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c46_checked = 1 Then
c46.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c47_checked = 1 Then
c47.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c48_checked = 1 Then
c48.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c49_checked = 1 Then
c49.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c50_checked = 1 Then
c50.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c51_checked = 1 Then
c51.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c52_checked = 1 Then
c52.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c53_checked = 1 Then
c53.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c54_checked = 1 Then
c54.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c55_checked = 1 Then
c55.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c56_checked = 1 Then
c56.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c57_checked = 1 Then
c57.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c58_checked = 1 Then
c58.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c59_checked = 1 Then
c59.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c60_checked = 1 Then
c60.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c61_checked = 1 Then
c61.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c62_checked = 1 Then
c62.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c63_checked = 1 Then
c63.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c64_checked = 1 Then
c64.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c65_checked = 1 Then
c65.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c66_checked = 1 Then
c66.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c67_checked = 1 Then
c67.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c68_checked = 1 Then
c68.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c69_checked = 1 Then
c69.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c70_checked = 1 Then
c71.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c71_checked = 1 Then
c71.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c72_checked = 1 Then
c72.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c73_checked = 1 Then
c73.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c74_checked = 1 Then
c74.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c75_checked = 1 Then
c75.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c76_checked = 1 Then
c76.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c77_checked = 1 Then
c77.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c78_checked = 1 Then
c78.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c79_checked = 1 Then
c79.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If
If c80_checked = 1 Then
c80.Picture = LoadResPicture("check_mark", bitmap)
boxes_checked = boxes_checked + 1
End If

spots_marked.Caption = boxes_checked
spots_marked.Refresh
End Function

Public Function autokeno(x As Long, y As Long)
Dim autonumber As Long
autobar.Visible = True
autoplaylabel.Visible = True
autobar.Max = x
coinsin_label.Caption = y
coinsin_label.Refresh
autostop = 0
For autonumber = 1 To x Step 1
    autoplaylabel.Caption = "Auto Play Progress.   " & autonumber & " of " & x & " Boards Drawn.   Press Escape to cancel."
    
    If autostop = 1 Then
    Exit For
    End If
    amountbet = y
    autobar.Value = autonumber
    dollars = dollars - amountbet
    credits_label.Caption = dollars
Call go
Next autonumber
autostop = 0
autobar.Visible = False
autoplaylabel.Visible = False
End Function
