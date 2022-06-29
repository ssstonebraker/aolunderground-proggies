VERSION 5.00
Begin VB.Form GFX 
   Caption         =   "GFX"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3120
   LinkTopic       =   "Form2"
   ScaleHeight     =   4065
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Tree1sF2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":0000
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   102
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox Turf 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Index           =   4
      Left            =   3360
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   101
      Top             =   120
      Width           =   5710
   End
   Begin VB.PictureBox Turf 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Index           =   3
      Left            =   3360
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   100
      Top             =   120
      Width           =   5710
   End
   Begin VB.PictureBox Turf 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Index           =   2
      Left            =   3240
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   99
      Top             =   120
      Width           =   5710
   End
   Begin VB.PictureBox Turf 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Index           =   1
      Left            =   3120
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   98
      Top             =   0
      Width           =   5710
   End
   Begin VB.PictureBox CA1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   120
      Picture         =   "GFX.frx":024A
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   97
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox BirdM3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   150
      Left            =   1680
      Picture         =   "GFX.frx":0C84
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   96
      Top             =   3480
      Width           =   165
   End
   Begin VB.PictureBox BirdM2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   150
      Left            =   1800
      Picture         =   "GFX.frx":0CE6
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   95
      Top             =   3360
      Width           =   165
   End
   Begin VB.PictureBox BirdM1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   150
      Left            =   1680
      Picture         =   "GFX.frx":0D48
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   94
      Top             =   3360
      Width           =   165
   End
   Begin VB.PictureBox GC1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   5
      Left            =   2520
      Picture         =   "GFX.frx":0DAA
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   93
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox GC1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   4
      Left            =   2760
      Picture         =   "GFX.frx":0E28
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   92
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox GC1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   3
      Left            =   2520
      Picture         =   "GFX.frx":0EA6
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   91
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox GC1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   2
      Left            =   2760
      Picture         =   "GFX.frx":0F24
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   90
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox GC1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   1
      Left            =   2760
      Picture         =   "GFX.frx":116E
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   89
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox GC1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   2520
      Picture         =   "GFX.frx":13B8
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   88
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox BusS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":1602
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   87
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox BusM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2040
      Picture         =   "GFX.frx":184C
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   86
      Top             =   3240
      Width           =   255
   End
   Begin VB.PictureBox Tree2s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2520
      Picture         =   "GFX.frx":18CA
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   85
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Tree2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2760
      Picture         =   "GFX.frx":1B14
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   84
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Tree1sF 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2520
      Picture         =   "GFX.frx":1B92
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   83
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Tree1sSS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":1DDC
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   82
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Tree1sW 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2040
      Picture         =   "GFX.frx":2026
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   81
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Tree1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2760
      Picture         =   "GFX.frx":2270
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   80
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox RoadIm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2280
      Picture         =   "GFX.frx":22EE
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   79
      Top             =   2760
      Width           =   300
   End
   Begin VB.PictureBox RoadIs 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   2040
      Picture         =   "GFX.frx":2378
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   78
      Top             =   2760
      Width           =   300
   End
   Begin VB.PictureBox ES2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   1080
      Picture         =   "GFX.frx":26BA
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   77
      Top             =   2520
      Width           =   450
   End
   Begin VB.PictureBox ES1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   600
      Picture         =   "GFX.frx":2F6C
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   76
      Top             =   2520
      Width           =   450
   End
   Begin VB.PictureBox ESM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   120
      Picture         =   "GFX.frx":381E
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   75
      Top             =   2520
      Width           =   450
   End
   Begin VB.PictureBox BR4M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   450
      Left            =   600
      Picture         =   "GFX.frx":38D4
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   20
      Top             =   1560
      Width           =   450
   End
   Begin VB.PictureBox BR3M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   600
      Picture         =   "GFX.frx":3986
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   19
      Top             =   1080
      Width           =   450
   End
   Begin VB.PictureBox BR2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   600
      Picture         =   "GFX.frx":3A3C
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   18
      Top             =   600
      Width           =   450
   End
   Begin VB.PictureBox BR1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   600
      Picture         =   "GFX.frx":3AF2
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   17
      Top             =   120
      Width           =   450
   End
   Begin VB.PictureBox BI3M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   1560
      Picture         =   "GFX.frx":3BA8
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   16
      Top             =   2040
      Width           =   450
   End
   Begin VB.PictureBox BI2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   1560
      Picture         =   "GFX.frx":3C5E
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   15
      Top             =   1560
      Width           =   450
   End
   Begin VB.PictureBox BI1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   1560
      Picture         =   "GFX.frx":3D14
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   14
      Top             =   1080
      Width           =   450
   End
   Begin VB.PictureBox BC3M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   1560
      Picture         =   "GFX.frx":45C6
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   13
      Top             =   600
      Width           =   450
   End
   Begin VB.PictureBox BC2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   1560
      Picture         =   "GFX.frx":467C
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   12
      Top             =   120
      Width           =   450
   End
   Begin VB.PictureBox BC1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   600
      Picture         =   "GFX.frx":4732
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   11
      Top             =   2040
      Width           =   450
   End
   Begin VB.PictureBox BR4s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   450
      Left            =   120
      Picture         =   "GFX.frx":47E8
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   74
      Top             =   1560
      Width           =   450
   End
   Begin VB.PictureBox BR3s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   120
      Picture         =   "GFX.frx":504A
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   73
      Top             =   1080
      Width           =   450
   End
   Begin VB.PictureBox BR2s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   120
      Picture         =   "GFX.frx":58FC
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   72
      Top             =   600
      Width           =   450
   End
   Begin VB.PictureBox BR1s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   120
      Picture         =   "GFX.frx":61AE
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   71
      Top             =   120
      Width           =   450
   End
   Begin VB.PictureBox BI3s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   1080
      Picture         =   "GFX.frx":6A60
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   70
      Top             =   2040
      Width           =   450
   End
   Begin VB.PictureBox BI2s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   1080
      Picture         =   "GFX.frx":7312
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   69
      Top             =   1560
      Width           =   450
   End
   Begin VB.PictureBox BI1s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   1080
      Picture         =   "GFX.frx":7BC4
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   68
      Top             =   1080
      Width           =   450
   End
   Begin VB.PictureBox BC3s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   1080
      Picture         =   "GFX.frx":8476
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   67
      Top             =   600
      Width           =   450
   End
   Begin VB.PictureBox BC2s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   1080
      Picture         =   "GFX.frx":8D28
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   66
      Top             =   120
      Width           =   450
   End
   Begin VB.PictureBox BC1s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   465
      Left            =   120
      Picture         =   "GFX.frx":95DA
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   65
      Top             =   2040
      Width           =   450
   End
   Begin VB.PictureBox rUDM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2520
      Picture         =   "GFX.frx":9E8C
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   64
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox rLRM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2760
      Picture         =   "GFX.frx":9F0A
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   63
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox rC2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2520
      Picture         =   "GFX.frx":9F88
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   62
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox rC1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2520
      Picture         =   "GFX.frx":A006
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   61
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox rT2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2760
      Picture         =   "GFX.frx":A084
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   60
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox rT1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2760
      Picture         =   "GFX.frx":A102
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   59
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox rC3M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2520
      Picture         =   "GFX.frx":A180
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   58
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox rT3M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2760
      Picture         =   "GFX.frx":A1FE
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   57
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox rC4M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2520
      Picture         =   "GFX.frx":A27C
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   56
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox rT4M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2760
      Picture         =   "GFX.frx":A2FA
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   55
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox rUDs 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":A378
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   54
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox rLRs 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2040
      Picture         =   "GFX.frx":A5C2
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   53
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox rC2s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":A80C
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   52
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox rC1s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":AA56
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   51
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox rT2s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2040
      Picture         =   "GFX.frx":ACA0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   50
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox rT1s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2040
      Picture         =   "GFX.frx":AEEA
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   49
      Top             =   1560
      Width           =   255
   End
   Begin VB.PictureBox rC3s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":B134
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   48
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox rT3s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2040
      Picture         =   "GFX.frx":B37E
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   47
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox rC4s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":B5C8
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   46
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox rT4s 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2040
      Picture         =   "GFX.frx":B812
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   45
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox h2Sgr 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":BA5C
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   44
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox h1Sgr 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2040
      Picture         =   "GFX.frx":BCA6
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   43
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox h2Sbr 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":BF18
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   42
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox h1Sbr 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2040
      Picture         =   "GFX.frx":C162
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   41
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox c2S 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2760
      Picture         =   "GFX.frx":C3D4
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   40
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox c1S 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2760
      Picture         =   "GFX.frx":C646
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   39
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox i4S 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2520
      Picture         =   "GFX.frx":C8B8
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   38
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox i3S 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2520
      Picture         =   "GFX.frx":CB02
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   37
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox i2S 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2520
      Picture         =   "GFX.frx":CD4C
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   36
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox i1S 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2520
      Picture         =   "GFX.frx":CF96
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   35
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox c4S 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2760
      Picture         =   "GFX.frx":D1E0
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   34
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox c3S 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2760
      Picture         =   "GFX.frx":D452
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   33
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox h2SMon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":D6C4
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   32
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox h1SMon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2040
      Picture         =   "GFX.frx":D90E
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   31
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox c2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2280
      Picture         =   "GFX.frx":DB80
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   30
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox c1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2040
      Picture         =   "GFX.frx":DC02
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   29
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox i4M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2760
      Picture         =   "GFX.frx":DC84
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   28
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox i3M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2520
      Picture         =   "GFX.frx":DD02
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   27
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox i2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":DD80
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   26
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox i1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2040
      Picture         =   "GFX.frx":DDFE
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   25
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox c4M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2760
      Picture         =   "GFX.frx":DE7C
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   24
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox c3M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2520
      Picture         =   "GFX.frx":DEFE
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   23
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   2280
      Picture         =   "GFX.frx":DF80
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   22
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox h1M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2040
      Picture         =   "GFX.frx":DFFE
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   21
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox c42 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1800
      Picture         =   "GFX.frx":E080
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   10
      Top             =   3000
      Width           =   285
   End
   Begin VB.PictureBox c41 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1560
      Picture         =   "GFX.frx":E392
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   9
      Top             =   3000
      Width           =   285
   End
   Begin VB.PictureBox c32 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1800
      Picture         =   "GFX.frx":E6A4
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   8
      Top             =   2760
      Width           =   285
   End
   Begin VB.PictureBox c31 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1560
      Picture         =   "GFX.frx":E9B6
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   7
      Top             =   2760
      Width           =   285
   End
   Begin VB.PictureBox CA2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1800
      Picture         =   "GFX.frx":ECC8
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   6
      Top             =   2520
      Width           =   285
   End
   Begin VB.PictureBox cur2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   1560
      Picture         =   "GFX.frx":EFDA
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   5
      Top             =   2520
      Width           =   285
   End
   Begin VB.PictureBox C22 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   1080
      Picture         =   "GFX.frx":F060
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   4
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox C21 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   600
      Picture         =   "GFX.frx":FA9A
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   3
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox C12 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   1080
      Picture         =   "GFX.frx":104D4
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   2
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox C11 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   600
      Picture         =   "GFX.frx":10F0E
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   1
      Top             =   3000
      Width           =   495
   End
   Begin VB.PictureBox curM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   120
      Picture         =   "GFX.frx":11948
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   0
      Top             =   3000
      Width           =   495
   End
End
Attribute VB_Name = "GFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'bmp_rotate Picture2, Picture4(0), (p.a * 10) * (Pi / 180)
'bmp_rotate Picture3, Picture4(1), (p.a * 10) * (Pi / 180)

End Sub
