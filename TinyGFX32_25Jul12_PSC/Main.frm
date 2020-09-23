VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8835
   ClientLeft      =   165
   ClientTop       =   135
   ClientWidth     =   17790
   ForeColor       =   &H00000000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   589
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1186
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdORGCul 
      BackColor       =   &H00FFFFFF&
      DownPicture     =   "Main.frx":107A
      Height          =   240
      Left            =   1305
      Picture         =   "Main.frx":11C4
      Style           =   1  'Graphical
      TabIndex        =   172
      ToolTipText     =   " Show Original colors "
      Top             =   825
      Width           =   360
   End
   Begin VB.CommandButton cmdAlpha 
      DownPicture     =   "Main.frx":1306
      Height          =   240
      Left            =   870
      Picture         =   "Main.frx":1450
      Style           =   1  'Graphical
      TabIndex        =   164
      ToolTipText     =   " Show Alphas for 32bpp images (White is transparent) "
      Top             =   825
      Width           =   360
   End
   Begin VB.PictureBox picSEL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   2
      Left            =   11640
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   163
      Top             =   2745
      Width           =   960
   End
   Begin VB.PictureBox picSEL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   1
      Left            =   11625
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   162
      Top             =   1740
      Width           =   960
   End
   Begin VB.PictureBox picSEL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   0
      Left            =   11610
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   161
      Top             =   720
      Width           =   960
   End
   Begin VB.PictureBox picElements 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   5160
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   146
      TabIndex        =   160
      ToolTipText     =   " on whole image or selection (RC to fix except B & LR) "
      Top             =   45
      Width           =   2220
   End
   Begin VB.PictureBox picToolBar2 
      BackColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   1065
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   258
      TabIndex        =   147
      Top             =   15
      Width           =   3930
      Begin VB.OptionButton optTools 
         Height          =   300
         Index           =   33
         Left            =   3495
         Picture         =   "Main.frx":1592
         Style           =   1  'Graphical
         TabIndex        =   174
         ToolTipText     =   " Box Horz & Vert shading "
         Top             =   -15
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   300
         Index           =   32
         Left            =   3180
         Picture         =   "Main.frx":16C4
         Style           =   1  'Graphical
         TabIndex        =   157
         ToolTipText     =   " Box horz center shading "
         Top             =   0
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   300
         Index           =   31
         Left            =   2880
         Picture         =   "Main.frx":1BFE
         Style           =   1  'Graphical
         TabIndex        =   156
         ToolTipText     =   " Box vert center shading "
         Top             =   15
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   300
         Index           =   30
         Left            =   2535
         Picture         =   "Main.frx":2138
         Style           =   1  'Graphical
         TabIndex        =   155
         ToolTipText     =   " Box diag / shading "
         Top             =   -15
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   300
         Index           =   29
         Left            =   2190
         Picture         =   "Main.frx":2672
         Style           =   1  'Graphical
         TabIndex        =   154
         ToolTipText     =   " Box diag \ shading "
         Top             =   0
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   300
         Index           =   28
         Left            =   1815
         Picture         =   "Main.frx":2BAC
         Style           =   1  'Graphical
         TabIndex        =   153
         ToolTipText     =   " Ellipse Horz Shading "
         Top             =   -15
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   300
         Index           =   27
         Left            =   1500
         Picture         =   "Main.frx":310E
         Style           =   1  'Graphical
         TabIndex        =   152
         ToolTipText     =   " Ellipse Vert Shading "
         Top             =   0
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   300
         Index           =   26
         Left            =   1080
         Picture         =   "Main.frx":3670
         Style           =   1  'Graphical
         TabIndex        =   151
         ToolTipText     =   " Ellipse Center Shading "
         Top             =   0
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   300
         Index           =   25
         Left            =   750
         Picture         =   "Main.frx":3BD2
         Style           =   1  'Graphical
         TabIndex        =   150
         ToolTipText     =   " Box Horz Shading "
         Top             =   0
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   300
         Index           =   24
         Left            =   390
         Picture         =   "Main.frx":4134
         Style           =   1  'Graphical
         TabIndex        =   149
         ToolTipText     =   " Box Vert Shading "
         Top             =   0
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   300
         Index           =   23
         Left            =   60
         Picture         =   "Main.frx":4696
         Style           =   1  'Graphical
         TabIndex        =   148
         ToolTipText     =   " Box Center Shading  "
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C0C0C0&
      Height          =   1740
      Left            =   9735
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   140
      Top             =   5145
      Width           =   555
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Index           =   19
         Left            =   90
         Picture         =   "Main.frx":47C8
         Style           =   1  'Graphical
         TabIndex        =   144
         ToolTipText     =   " Reflect to Right "
         Top             =   75
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Index           =   20
         Left            =   90
         Picture         =   "Main.frx":48CA
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   " Reflect to Left "
         Top             =   480
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Index           =   21
         Left            =   90
         Picture         =   "Main.frx":49CC
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   " Reflect to Below "
         Top             =   885
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Index           =   22
         Left            =   90
         Picture         =   "Main.frx":4ACE
         Style           =   1  'Graphical
         TabIndex        =   141
         ToolTipText     =   " Reflect to Above "
         Top             =   1290
         Width           =   330
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C0C0&
      Height          =   1605
      Left            =   9510
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   126
      Top             =   300
      Width           =   1020
      Begin VB.CommandButton cmdFlash 
         Appearance      =   0  'Flat
         DownPicture     =   "Main.frx":4BD0
         Height          =   195
         Left            =   525
         Picture         =   "Main.frx":4CA2
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   " Show Hotspot "
         Top             =   1335
         Width           =   180
      End
      Begin VB.VScrollBar scrHotY 
         Height          =   900
         LargeChange     =   2
         Left            =   45
         Max             =   31
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   585
         Width           =   240
      End
      Begin VB.HScrollBar scrHotX 
         Height          =   195
         LargeChange     =   2
         Left            =   45
         Max             =   31
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "HotXY    for"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   390
         TabIndex        =   133
         Top             =   600
         Width           =   495
      End
      Begin VB.Label LabImCurNum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   525
         TabIndex        =   131
         Top             =   1035
         Width           =   180
      End
      Begin VB.Label LabHotXY 
         BackColor       =   &H00C0C0C0&
         Caption         =   "X=0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   130
         Top             =   75
         Width           =   390
      End
      Begin VB.Label LabHotXY 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Y=0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   129
         Top             =   75
         Width           =   405
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0C0C0&
      Height          =   765
      Left            =   2850
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   508
      TabIndex        =   110
      Top             =   6930
      Width           =   7680
      Begin VB.CheckBox chkGridOnOff 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Grid On"
         Height          =   195
         Left            =   6705
         TabIndex        =   125
         Top             =   375
         Width           =   930
      End
      Begin VB.CommandButton cmdSwapLR 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   510
         Picture         =   "Main.frx":4D74
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   " Swap Left & Right Colors "
         Top             =   60
         Width           =   675
      End
      Begin VB.CommandButton cmdPAL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4"
         Height          =   225
         Index           =   4
         Left            =   4830
         TabIndex        =   118
         ToolTipText     =   " Centered palette "
         Top             =   45
         Width           =   300
      End
      Begin VB.CommandButton cmdPAL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "3"
         Height          =   225
         Index           =   3
         Left            =   4500
         TabIndex        =   117
         ToolTipText     =   " Grey palette "
         Top             =   45
         Width           =   300
      End
      Begin VB.CommandButton cmdPAL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2"
         Height          =   225
         Index           =   2
         Left            =   4140
         TabIndex        =   116
         ToolTipText     =   " Long banded palette "
         Top             =   60
         Width           =   300
      End
      Begin VB.CommandButton cmdPAL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1"
         Height          =   225
         Index           =   1
         Left            =   3780
         TabIndex        =   115
         ToolTipText     =   " Short banded palette "
         Top             =   60
         Width           =   300
      End
      Begin VB.CommandButton cmdPAL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   225
         Index           =   0
         Left            =   3450
         TabIndex        =   114
         ToolTipText     =   " QB palette "
         Top             =   60
         Width           =   300
      End
      Begin VB.PictureBox picPAL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   75
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   385
         TabIndex        =   113
         Top             =   360
         Width           =   5805
      End
      Begin VB.PictureBox picErase 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5985
         Picture         =   "Main.frx":4FFE
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   112
         ToolTipText     =   " Make Left or Right color Erase with transparent color "
         Top             =   375
         Width           =   300
      End
      Begin VB.CommandButton cmdPAL 
         BackColor       =   &H00C0C0C0&
         Height          =   225
         Index           =   5
         Left            =   5160
         Picture         =   "Main.frx":5148
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   " System color picker "
         Top             =   45
         Width           =   300
      End
      Begin VB.Label LabRGB 
         BackColor       =   &H00C0C0C0&
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   6285
         TabIndex        =   138
         Top             =   75
         Width           =   255
      End
      Begin VB.Label LabRGB 
         BackColor       =   &H00C0C0C0&
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   180
         Index           =   1
         Left            =   5985
         TabIndex        =   137
         Top             =   75
         Width           =   255
      End
      Begin VB.Label LabRGB 
         BackColor       =   &H00C0C0C0&
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   0
         Left            =   5685
         TabIndex        =   136
         Top             =   75
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Palettes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2775
         TabIndex        =   124
         Top             =   75
         Width           =   615
      End
      Begin VB.Label LabErase 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Erase color on"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   1725
         TabIndex        =   123
         Top             =   75
         Width           =   1185
      End
      Begin VB.Label LabColor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   1215
         TabIndex        =   122
         ToolTipText     =   " Right color "
         Top             =   30
         Width           =   375
      End
      Begin VB.Label LabColor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   75
         TabIndex        =   121
         ToolTipText     =   " Left color "
         Top             =   30
         Width           =   390
      End
      Begin VB.Label LabDropper 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6720
         TabIndex        =   120
         ToolTipText     =   " Color show "
         Top             =   45
         Width           =   840
      End
      Begin VB.Line LineErase 
         BorderColor     =   &H000000FF&
         BorderWidth     =   4
         X1              =   424
         X2              =   425
         Y1              =   24
         Y2              =   40
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   1740
      Left            =   8835
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   105
      Top             =   5145
      Width           =   555
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Index           =   18
         Left            =   90
         Picture         =   "Main.frx":568E
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   " Mirror bottom "
         Top             =   1290
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Index           =   17
         Left            =   90
         Picture         =   "Main.frx":5790
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   " Mirror top "
         Top             =   885
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Index           =   16
         Left            =   90
         Picture         =   "Main.frx":5892
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   " Mirror right "
         Top             =   480
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Index           =   15
         Left            =   90
         Picture         =   "Main.frx":5994
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   " Mirror left "
         Top             =   75
         Width           =   330
      End
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   2
      Left            =   12765
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   103
      Top             =   5925
      Width           =   960
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   1
      Left            =   12780
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   102
      Top             =   4890
      Width           =   960
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   0
      Left            =   12780
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   101
      Top             =   3855
      Width           =   960
   End
   Begin VB.CommandButton cmdMask 
      DownPicture     =   "Main.frx":5A96
      Height          =   240
      Left            =   435
      Picture         =   "Main.frx":5BE0
      Style           =   1  'Graphical
      TabIndex        =   100
      ToolTipText     =   " Show Masks (White is transparent) "
      Top             =   825
      Width           =   360
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   4470
      Left            =   8775
      ScaleHeight     =   294
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   84
      Top             =   300
      Width           =   690
      Begin VB.OptionButton optTools2 
         Height          =   330
         Index           =   0
         Left            =   150
         Picture         =   "Main.frx":5D2A
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   " Blur tool "
         Top             =   45
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         Height          =   330
         Index           =   1
         Left            =   150
         Picture         =   "Main.frx":62B4
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   " Grey  tool "
         Top             =   405
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         Height          =   330
         Index           =   2
         Left            =   150
         Picture         =   "Main.frx":63B6
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   " Speckle Invert tool "
         Top             =   765
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         Height          =   330
         Index           =   3
         Left            =   150
         Picture         =   "Main.frx":64B8
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   " Brighter tool "
         Top             =   1125
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         Height          =   330
         Index           =   4
         Left            =   150
         Picture         =   "Main.frx":69FA
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   " Darker tool "
         Top             =   1455
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         Height          =   330
         Index           =   5
         Left            =   150
         Picture         =   "Main.frx":6F3C
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   " Checkerboard "
         Top             =   1815
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         Height          =   330
         Index           =   6
         Left            =   150
         Picture         =   "Main.frx":6FC6
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   " Horz lines "
         Top             =   2175
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         Height          =   330
         Index           =   7
         Left            =   150
         Picture         =   "Main.frx":7050
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   " Vert lines "
         Top             =   2535
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         Height          =   330
         Index           =   8
         Left            =   150
         Picture         =   "Main.frx":70DA
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   " Random strip "
         Top             =   2895
         Width           =   330
      End
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   9
         Left            =   75
         Picture         =   "Main.frx":761C
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   " More blue "
         Top             =   3285
         Width           =   240
      End
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00000000&
         Height          =   330
         Index           =   10
         Left            =   345
         Picture         =   "Main.frx":76A6
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   " Less blue "
         Top             =   3285
         Width           =   240
      End
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   11
         Left            =   75
         Picture         =   "Main.frx":7730
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   " More green "
         Top             =   3675
         Width           =   240
      End
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00000000&
         Height          =   330
         Index           =   12
         Left            =   345
         Picture         =   "Main.frx":77BA
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   " Less green "
         Top             =   3675
         Width           =   240
      End
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   13
         Left            =   75
         Picture         =   "Main.frx":7844
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   " More red "
         Top             =   4065
         Width           =   240
      End
      Begin VB.OptionButton optTools2 
         BackColor       =   &H00000000&
         Height          =   330
         Index           =   14
         Left            =   345
         Picture         =   "Main.frx":78CE
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   " Less red "
         Top             =   4065
         Width           =   240
      End
   End
   Begin VB.PictureBox picC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5865
      Left            =   2835
      ScaleHeight     =   389
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   389
      TabIndex        =   40
      Top             =   1020
      Width           =   5865
      Begin VB.PictureBox picPANEL 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1920
         Left            =   465
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   41
         Top             =   285
         Width           =   1920
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            Height          =   480
            Left            =   240
            Top             =   1335
            Width           =   585
         End
         Begin VB.Shape shpCirc 
            BorderColor     =   &H00FFFFFF&
            DrawMode        =   7  'Invert
            Height          =   270
            Left            =   1125
            Shape           =   3  'Circle
            Top             =   1260
            Width           =   255
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   15  'Merge Pen Not
            Index           =   1
            X1              =   79
            X2              =   105
            Y1              =   36
            Y2              =   51
         End
         Begin VB.Shape Ellipse1 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   15  'Merge Pen Not
            Height          =   300
            Index           =   1
            Left            =   720
            Shape           =   2  'Oval
            Top             =   570
            Width           =   345
         End
         Begin VB.Shape Box1 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   15  'Merge Pen Not
            Height          =   315
            Index           =   1
            Left            =   225
            Top             =   825
            Width           =   345
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   6  'Inside Solid
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            Index           =   0
            X1              =   78
            X2              =   104
            Y1              =   13
            Y2              =   28
         End
         Begin VB.Shape Ellipse1 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   6  'Inside Solid
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            Height          =   300
            Index           =   0
            Left            =   660
            Shape           =   2  'Oval
            Top             =   135
            Width           =   345
         End
         Begin VB.Shape Box1 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   6  'Inside Solid
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            Height          =   315
            Index           =   0
            Left            =   225
            Top             =   435
            Width           =   345
         End
      End
   End
   Begin VB.PictureBox picORG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   2
      Left            =   12765
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   63
      Top             =   2745
      Width           =   960
   End
   Begin VB.PictureBox picORG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   1
      Left            =   12780
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   62
      Top             =   1725
      Width           =   960
   End
   Begin VB.PictureBox picORG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   0
      Left            =   12765
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   61
      Top             =   705
      Width           =   960
   End
   Begin VB.PictureBox picSmallBU 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   11640
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   52
      Top             =   5910
      Width           =   960
   End
   Begin VB.PictureBox picfraWH 
      BackColor       =   &H00C0C0C0&
      Height          =   1755
      Left            =   375
      ScaleHeight     =   1695
      ScaleWidth      =   1995
      TabIndex        =   44
      Top             =   5925
      Width           =   2055
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "Backup"
         Height          =   225
         Left            =   540
         TabIndex        =   171
         Top             =   1470
         Width           =   930
      End
      Begin VB.HScrollBar scrWidth 
         Height          =   285
         LargeChange     =   2
         Left            =   630
         Max             =   64
         Min             =   1
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1110
         Value           =   8
         Width           =   1200
      End
      Begin VB.VScrollBar scrHeight 
         Height          =   1365
         LargeChange     =   2
         Left            =   240
         Max             =   64
         Min             =   1
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   30
         Value           =   64
         Width           =   285
      End
      Begin VB.Label LabWidth 
         BackColor       =   &H00C0C0C0&
         Caption         =   " Width = 8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   750
         TabIndex        =   50
         Top             =   780
         Width           =   960
      End
      Begin VB.Label LabHeight 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Height = 8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   750
         TabIndex        =   49
         Top             =   480
         Width           =   960
      End
      Begin VB.Label LabImageNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1605
         TabIndex        =   48
         Top             =   60
         Width           =   195
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "for Image"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   735
         TabIndex        =   47
         Top             =   90
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdLRUD 
      Height          =   345
      Index           =   5
      Left            =   2595
      Picture         =   "Main.frx":7958
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   " Flip vert "
      Top             =   1785
      Width           =   195
   End
   Begin VB.CommandButton cmdLRUD 
      Height          =   195
      Index           =   4
      Left            =   3615
      Picture         =   "Main.frx":79EA
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   " Flip Horz "
      Top             =   780
      Width           =   345
   End
   Begin VB.CommandButton cmdLRUD 
      Height          =   345
      Index           =   3
      Left            =   2595
      Picture         =   "Main.frx":7A5C
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   " Scroll down "
      Top             =   1380
      Width           =   195
   End
   Begin VB.CommandButton cmdLRUD 
      Height          =   345
      Index           =   2
      Left            =   2595
      Picture         =   "Main.frx":7AEE
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   " Scroll up "
      Top             =   1005
      Width           =   195
   End
   Begin VB.CommandButton cmdLRUD 
      Height          =   195
      Index           =   1
      Left            =   3225
      Picture         =   "Main.frx":7B80
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   " Scroll right "
      Top             =   780
      Width           =   345
   End
   Begin VB.CommandButton cmdLRUD 
      Height          =   195
      Index           =   0
      Left            =   2850
      Picture         =   "Main.frx":7BF2
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   " Scroll left "
      Top             =   780
      Width           =   345
   End
   Begin VB.PictureBox picSmallFrame 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   4470
      Left            =   375
      ScaleHeight     =   294
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   3
      Top             =   1185
      Width           =   2055
      Begin VB.CommandButton cmdUndoALL 
         BackColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   2
         Left            =   1530
         Picture         =   "Main.frx":7C64
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   " Undo all & fix current image "
         Top             =   3945
         Width           =   360
      End
      Begin VB.CommandButton cmdUndoALL 
         BackColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   1
         Left            =   1530
         Picture         =   "Main.frx":7DAE
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   " Undo all & fix current image "
         Top             =   2475
         Width           =   360
      End
      Begin VB.CommandButton cmdUndoALL 
         BackColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   0
         Left            =   1530
         Picture         =   "Main.frx":7EF8
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   " Undo all & fix current image "
         Top             =   990
         Width           =   360
      End
      Begin VB.CommandButton cmdReload 
         BackColor       =   &H00E0E0E0&
         Caption         =   "R"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   73
         ToolTipText     =   " Reload original image "
         Top             =   3645
         Width           =   210
      End
      Begin VB.CommandButton cmdReload 
         BackColor       =   &H00E0E0E0&
         Caption         =   "R"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   72
         ToolTipText     =   " Reload original image "
         Top             =   2175
         Width           =   210
      End
      Begin VB.CommandButton cmdReload 
         BackColor       =   &H00E0E0E0&
         Caption         =   "R"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   71
         ToolTipText     =   " Reload original image "
         Top             =   675
         Width           =   210
      End
      Begin VB.CommandButton cmdRedo 
         BackColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   2
         Left            =   1530
         Picture         =   "Main.frx":8042
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   " Redo "
         Top             =   3660
         Width           =   360
      End
      Begin VB.CommandButton cmdRedo 
         BackColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   1
         Left            =   1530
         Picture         =   "Main.frx":818C
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   " Redo "
         Top             =   2190
         Width           =   360
      End
      Begin VB.CommandButton cmdRedo 
         BackColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   0
         Left            =   1530
         Picture         =   "Main.frx":82D6
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   " Redo "
         Top             =   705
         Width           =   360
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   2
         Left            =   75
         Picture         =   "Main.frx":8420
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   " Clear image "
         Top             =   3315
         Width           =   210
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   1
         Left            =   75
         Picture         =   "Main.frx":8502
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   " Clear image "
         Top             =   1845
         Width           =   210
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   75
         Picture         =   "Main.frx":85E4
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   " Clear image "
         Top             =   360
         Width           =   210
      End
      Begin VB.CommandButton cmdUndo 
         BackColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   2
         Left            =   1530
         Picture         =   "Main.frx":86C6
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   " Undo "
         Top             =   3405
         Width           =   360
      End
      Begin VB.CommandButton cmdUndo 
         BackColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   1
         Left            =   1530
         Picture         =   "Main.frx":8810
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   " Undo "
         Top             =   1935
         Width           =   360
      End
      Begin VB.CommandButton cmdUndo 
         BackColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   0
         Left            =   1530
         Picture         =   "Main.frx":895A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   " Undo "
         Top             =   450
         Width           =   360
      End
      Begin VB.CommandButton cmdSelect 
         Height          =   225
         Index           =   2
         Left            =   1530
         Picture         =   "Main.frx":8AA4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   " Edit "
         Top             =   3090
         Width           =   360
      End
      Begin VB.CommandButton cmdSelect 
         Height          =   225
         Index           =   1
         Left            =   1530
         Picture         =   "Main.frx":8DF2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   " Edit "
         Top             =   1620
         Width           =   360
      End
      Begin VB.CommandButton cmdSelect 
         Height          =   225
         Index           =   0
         Left            =   1530
         Picture         =   "Main.frx":9140
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   " Edit "
         Top             =   135
         Width           =   360
      End
      Begin VB.PictureBox picSmall 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   2
         Left            =   405
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   6
         Top             =   3015
         Width           =   960
      End
      Begin VB.PictureBox picSmall 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   1
         Left            =   390
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   5
         Top             =   1545
         Width           =   960
      End
      Begin VB.PictureBox picSmall 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   0
         Left            =   405
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   4
         Top             =   60
         Width           =   960
      End
      Begin VB.Label LabMaxGen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   2
         Left            =   1605
         TabIndex        =   170
         ToolTipText     =   " Number of Undos for Image 3 "
         Top             =   4155
         Width           =   345
      End
      Begin VB.Label LabMaxGen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   1
         Left            =   1605
         TabIndex        =   169
         ToolTipText     =   " Number of Undos for Image 2 "
         Top             =   2700
         Width           =   345
      End
      Begin VB.Label LabMaxGen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   0
         Left            =   1605
         TabIndex        =   168
         ToolTipText     =   " Number of Undos for Image 1 "
         Top             =   1215
         Width           =   345
      End
      Begin VB.Label LabCurGen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   1260
         TabIndex        =   167
         ToolTipText     =   " Undo number for Image 3 "
         Top             =   4155
         Width           =   345
      End
      Begin VB.Label LabCurGen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1260
         TabIndex        =   166
         ToolTipText     =   " Undo number for Image 2"
         Top             =   2700
         Width           =   345
      End
      Begin VB.Label LabCurGen 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   1260
         TabIndex        =   165
         ToolTipText     =   " Undo number for Image1 "
         Top             =   1215
         Width           =   345
      End
      Begin VB.Label LabWH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "W x H = 48 x 48"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   15
         TabIndex        =   18
         Top             =   4155
         Width           =   1245
      End
      Begin VB.Label LabWH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "W x H = 48 x 48"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   17
         Top             =   2700
         Width           =   1245
      End
      Begin VB.Label LabWH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "W x H = 48 x 48"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   16
         Top             =   1215
         Width           =   1245
      End
      Begin VB.Label LabNumColors 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Colors = 123456"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   2
         Left            =   15
         TabIndex        =   15
         Top             =   3975
         Width           =   900
      End
      Begin VB.Label LabNumColors 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Colors = 123456"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   1
         Left            =   30
         TabIndex        =   14
         Top             =   2520
         Width           =   900
      End
      Begin VB.Label LabNumColors 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Colors = 123456"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   15
         TabIndex        =   13
         Top             =   1035
         Width           =   900
      End
      Begin VB.Label LabNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&3"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   2
         Left            =   75
         TabIndex        =   9
         Top             =   3000
         Width           =   210
      End
      Begin VB.Label LabNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   1
         Left            =   75
         TabIndex        =   8
         Top             =   1530
         Width           =   210
      End
      Begin VB.Label LabNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   0
         Left            =   75
         TabIndex        =   7
         Top             =   45
         Width           =   210
      End
   End
   Begin VB.PictureBox picToolbar 
      BackColor       =   &H00C0C0C0&
      Height          =   405
      Left            =   375
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   551
      TabIndex        =   0
      Top             =   315
      Width           =   8325
      Begin VB.OptionButton optTools 
         Height          =   330
         Index           =   18
         Left            =   6375
         Picture         =   "Main.frx":948E
         Style           =   1  'Graphical
         TabIndex        =   158
         ToolTipText     =   " Border"
         Top             =   30
         Width           =   315
      End
      Begin VB.OptionButton optTools 
         Height          =   330
         Index           =   22
         Left            =   7815
         Picture         =   "Main.frx":95D8
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Copy selection: Right button on selection to fix "
         Top             =   15
         Width           =   315
      End
      Begin VB.OptionButton optTools 
         Height          =   330
         Index           =   21
         Left            =   7485
         Picture         =   "Main.frx":96DA
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   " Move selection: Right button on selection to fix "
         Top             =   15
         Width           =   315
      End
      Begin VB.OptionButton optTools 
         Height          =   330
         Index           =   20
         Left            =   7155
         Picture         =   "Main.frx":97DC
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   " To draw selection box (or Cancel operation) "
         Top             =   15
         Width           =   315
      End
      Begin VB.OptionButton optTools 
         Height          =   330
         Index           =   17
         Left            =   6060
         Picture         =   "Main.frx":9866
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   " Darker "
         Top             =   30
         Width           =   315
      End
      Begin VB.OptionButton optTools 
         Height          =   330
         Index           =   16
         Left            =   5745
         Picture         =   "Main.frx":9DA8
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   " Brighter "
         Top             =   30
         Width           =   315
      End
      Begin VB.OptionButton optTools 
         Height          =   330
         Index           =   15
         Left            =   5400
         Picture         =   "Main.frx":A2EA
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   " Invert "
         Top             =   30
         Width           =   315
      End
      Begin VB.OptionButton optTools 
         Height          =   330
         Index           =   14
         Left            =   5070
         Picture         =   "Main.frx":A374
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   " Left color-Relief "
         Top             =   30
         Width           =   315
      End
      Begin VB.OptionButton optTools 
         Height          =   330
         Index           =   13
         Left            =   4740
         Picture         =   "Main.frx":A8B6
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   " Blur "
         Top             =   30
         Width           =   315
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   4
         Left            =   1395
         Picture         =   "Main.frx":AE40
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   " Box: fill with Right color, border with Left color "
         Top             =   75
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   7
         Left            =   2415
         Picture         =   "Main.frx":B3CA
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Ellipse: fill with Right color, border with Left color "
         Top             =   90
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   330
         Index           =   19
         Left            =   6675
         Picture         =   "Main.frx":B514
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   " Replace Left by Right color "
         Top             =   30
         Width           =   315
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   9
         Left            =   3180
         Picture         =   "Main.frx":B616
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   " Color picker "
         Top             =   45
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   315
         Index           =   12
         Left            =   4185
         Picture         =   "Main.frx":B7C0
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   " Text "
         Top             =   30
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   315
         Index           =   11
         Left            =   3855
         Picture         =   "Main.frx":BD4A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   " Rotate 90 anti-clockwise "
         Top             =   45
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         Height          =   315
         Index           =   10
         Left            =   3540
         Picture         =   "Main.frx":BE4C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   " Rotate 90 clockwise "
         Top             =   45
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   8
         Left            =   2790
         Picture         =   "Main.frx":BF4E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   " Fill "
         Top             =   75
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   6
         Left            =   2100
         Picture         =   "Main.frx":C0F8
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   " Solid ellipse "
         Top             =   75
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   5
         Left            =   1755
         Picture         =   "Main.frx":C682
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   " Ellipse "
         Top             =   75
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   3
         Left            =   1110
         Picture         =   "Main.frx":CC0C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   " Solid box "
         Top             =   75
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Index           =   2
         Left            =   750
         Picture         =   "Main.frx":D196
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   " Box "
         Top             =   60
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   1
         Left            =   450
         Picture         =   "Main.frx":D720
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " Line "
         Top             =   60
         Width           =   300
      End
      Begin VB.OptionButton optTools 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   0
         Left            =   60
         Picture         =   "Main.frx":DCAA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   " Dot "
         Top             =   30
         Width           =   300
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   195
         X2              =   195
         Y1              =   6
         Y2              =   28
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   193
         X2              =   193
         Y1              =   6
         Y2              =   28
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   2895
      Left            =   9510
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   53
      Top             =   1875
      Width           =   1020
      Begin VB.PictureBox picVisColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1905
         Left            =   165
         ScaleHeight     =   127
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   74
         ToolTipText     =   " Temporarily show visibility against background colors. "
         Top             =   360
         Width           =   225
      End
      Begin VB.PictureBox picGrid 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1905
         Left            =   615
         ScaleHeight     =   127
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   54
         ToolTipText     =   " Set grid color. "
         Top             =   540
         Width           =   225
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   34
         X2              =   0
         Y1              =   159
         Y2              =   159
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   34
         X2              =   34
         Y1              =   27
         Y2              =   159
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   61
         X2              =   32
         Y1              =   28
         Y2              =   28
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Visibility"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   1
         Left            =   45
         TabIndex        =   135
         ToolTipText     =   " Click to show visibility against background colors "
         Top             =   75
         Width           =   900
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grid color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   134
         Top             =   2550
         Width           =   900
      End
   End
   Begin VB.Label LabDate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver: 25 Jul 2012"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9630
      TabIndex        =   173
      Top             =   8070
      Width           =   1185
   End
   Begin VB.Label LabTest 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   -150
      TabIndex        =   159
      Top             =   7965
      Width           =   11115
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Transparent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7845
      TabIndex        =   146
      Top             =   750
      Width           =   960
   End
   Begin VB.Image imTMarker 
      Height          =   150
      Left            =   7680
      Stretch         =   -1  'True
      ToolTipText     =   " Change transparency marker "
      Top             =   780
      Width           =   150
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reflectors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   2
      Left            =   9555
      TabIndex        =   145
      Top             =   4890
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cursor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   2
      Left            =   9510
      TabIndex        =   132
      ToolTipText     =   " on whole image or selection "
      Top             =   45
      Width           =   1020
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mirrors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   1
      Left            =   8775
      TabIndex        =   104
      ToolTipText     =   " on whole image or selection "
      Top             =   4890
      Width           =   690
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tools 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   0
      Left            =   8775
      TabIndex        =   83
      Top             =   45
      Width           =   690
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   2
      Left            =   7530
      TabIndex        =   78
      Top             =   45
      Width           =   1050
   End
   Begin VB.Label LabSpec 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   75
      Top             =   7725
      Width           =   10935
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   0
      Left            =   390
      TabIndex        =   57
      Top             =   45
      Width           =   600
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Set Width && Height"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   510
      TabIndex        =   51
      Top             =   5685
      Width           =   1770
   End
   Begin VB.Label LabTool 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   5340
      TabIndex        =   43
      ToolTipText     =   " Selected Tool "
      Top             =   765
      Width           =   2160
   End
   Begin VB.Label LabXY 
      BackColor       =   &H00E0E0E0&
      Caption         =   "X = 42  Y = 42"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4095
      TabIndex        =   39
      Top             =   765
      Width           =   1155
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   210
      Left            =   2595
      Picture         =   "Main.frx":E234
      ToolTipText     =   " Scrollers (whole image or a selection) "
      Top             =   780
      Width           =   210
   End
   Begin VB.Label LabMAC 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Images"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1755
      TabIndex        =   19
      Top             =   780
      Width           =   675
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Begin VB.Menu mnuOpenIntoImage 
            Caption         =   "Into Image 1"
            Index           =   0
         End
         Begin VB.Menu mnuOpenIntoImage 
            Caption         =   "Into Image 2"
            Index           =   1
         End
         Begin VB.Menu mnuOpenIntoImage 
            Caption         =   "Into Image 3"
            Index           =   2
         End
         Begin VB.Menu mnuOpenIntoImage 
            Caption         =   "(Will fill from Image 1 if multi-icon)"
            Index           =   3
         End
         Begin VB.Menu mnuOpenIntoImage 
            Caption         =   "(Alpha restricted)"
            Index           =   4
         End
      End
      Begin VB.Menu mnuPalette 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuPalette 
         Caption         =   "Open palette bmp"
         Index           =   1
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RFile2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RFile3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RFile4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RFile5"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RFile6"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RFile7"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "RFile8"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu ALine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveBMP 
         Caption         =   "Save As BMP"
         Begin VB.Menu mnuSaveBMPImage 
            Caption         =   "Save BMP Image 1"
            Index           =   0
         End
         Begin VB.Menu mnuSaveBMPImage 
            Caption         =   "Save BMP Image 2"
            Index           =   1
         End
         Begin VB.Menu mnuSaveBMPImage 
            Caption         =   "Save BMP Image 3"
            Index           =   2
         End
         Begin VB.Menu mnuSaveBMPImage 
            Caption         =   "Save Drawing panel (24bpp)"
            Index           =   3
         End
         Begin VB.Menu mnuSaveBMPImage 
            Caption         =   "(Optimized saving)"
            Index           =   4
         End
      End
      Begin VB.Menu mnuSaveICO 
         Caption         =   "Save As ICO (Standard 16, 24, 32, 48)"
         Begin VB.Menu mnuSaveICOImage 
            Caption         =   "Save ICO Image 1"
            Index           =   0
         End
         Begin VB.Menu mnuSaveICOImage 
            Caption         =   "Save ICO Image 2"
            Index           =   1
         End
         Begin VB.Menu mnuSaveICOImage 
            Caption         =   "Save ICO Image 3"
            Index           =   2
         End
         Begin VB.Menu mnuSaveICOImage 
            Caption         =   "Save ICO Images 1 && 2"
            Index           =   3
         End
         Begin VB.Menu mnuSaveICOImage 
            Caption         =   "Save ICO Images 1, 2 && 3"
            Index           =   4
         End
         Begin VB.Menu mnuSaveICOImage 
            Caption         =   "(Optimized saving)"
            Index           =   5
         End
      End
      Begin VB.Menu mnuSaveCUR 
         Caption         =   "Save As CUR (32 x 32 or 48 x 48 only)"
         Begin VB.Menu mnuSaveCURImage 
            Caption         =   "Save CUR Image 1"
            Index           =   0
         End
         Begin VB.Menu mnuSaveCURImage 
            Caption         =   "Save CUR Image 2"
            Index           =   1
         End
         Begin VB.Menu mnuSaveCURImage 
            Caption         =   "Save CUR Image 3"
            Index           =   2
         End
         Begin VB.Menu mnuSaveCURImage 
            Caption         =   "(Optimized saving)"
            Index           =   3
         End
      End
      Begin VB.Menu Brk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload original image"
         Begin VB.Menu mnuReloadImage 
            Caption         =   "Reload Image 1"
            Index           =   0
         End
         Begin VB.Menu mnuReloadImage 
            Caption         =   "Reload Image 2"
            Index           =   1
         End
         Begin VB.Menu mnuReloadImage 
            Caption         =   "Reload Image 3"
            Index           =   2
         End
      End
      Begin VB.Menu Brk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPrefs 
      Caption         =   "&Preferences"
      Begin VB.Menu mnuPreferences 
         Caption         =   "Optimized saving (according to number of colors)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "No optimization (BMP 24bpp, the rest 8bpp)"
         Index           =   1
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "HALFTONE for Stretch/Shrink && Clipboard"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "COLORONCOLOR for Stretch/Shrink && Clipboard"
         Index           =   4
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Keep aspect ratio for Capture && Clipboard pasting"
         Checked         =   -1  'True
         Index           =   6
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Ignore aspect ratio for Capture && Clipboard pasting"
         Index           =   7
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Marked transparent pixels-Square"
         Checked         =   -1  'True
         Index           =   9
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Marked transparent pixels-Diagonal"
         Index           =   10
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Unmarked transparent pixels"
         Index           =   11
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Back up all images from Extractor (<=20)"
         Index           =   13
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Just back up the final Extractor images"
         Checked         =   -1  'True
         Index           =   14
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Restrict alpha to opaque for 32bpp images"
         Index           =   16
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Query original colors on saving 32bpp images"
         Index           =   18
      End
   End
   Begin VB.Menu mnuSwap 
      Caption         =   "&Swaps"
      Begin VB.Menu mnuSwapImages 
         Caption         =   "Swap images 1 && 2"
         Index           =   0
      End
      Begin VB.Menu mnuSwapImages 
         Caption         =   "Swap images 1 && 3"
         Index           =   1
      End
      Begin VB.Menu mnuSwapImages 
         Caption         =   "Swap images 2 && 3"
         Index           =   2
      End
   End
   Begin VB.Menu mnuStretchShrink 
      Caption         =   "Stretch&/Shrink"
      Begin VB.Menu mnuStretchImage 
         Caption         =   "image 1 to 2"
         Index           =   0
      End
      Begin VB.Menu mnuStretchImage 
         Caption         =   "image 1 to 3"
         Index           =   1
      End
      Begin VB.Menu mnuStretchImage 
         Caption         =   "image 2 to 1"
         Index           =   2
      End
      Begin VB.Menu mnuStretchImage 
         Caption         =   "image 3 to 1"
         Index           =   3
      End
      Begin VB.Menu mnuStretchImage 
         Caption         =   "image 2 to 3"
         Index           =   4
      End
      Begin VB.Menu mnuStretchImage 
         Caption         =   "image 3 to 2"
         Index           =   5
      End
      Begin VB.Menu mnuStretchImage 
         Caption         =   "(HALFTONE)"
         Index           =   6
      End
   End
   Begin VB.Menu mnuCursor 
      Caption         =   "&Cursor"
      Begin VB.Menu mnuTestCursor 
         Caption         =   "Test cursor (or icon)"
         Index           =   0
      End
      Begin VB.Menu mnuTestCursor 
         Caption         =   "Cancel cursor"
         Index           =   1
      End
   End
   Begin VB.Menu mnuCapture 
      Caption         =   "C&apture"
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "to Image 1"
         Index           =   0
      End
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "to Image 2"
         Index           =   1
      End
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "to Image 3"
         Index           =   2
      End
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "Mag = x 1"
         Checked         =   -1  'True
         Index           =   4
      End
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "Mag = x 2"
         Index           =   5
      End
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "Mag = x 4"
         Index           =   6
      End
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "Drawn rectangle to image 1"
         Index           =   9
      End
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "Drawn rectangle to image 2"
         Index           =   10
      End
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "Drawn rectangle to image 3"
         Index           =   11
      End
      Begin VB.Menu mnuCaptureToImage 
         Caption         =   "(Keep aspect ratio)"
         Index           =   12
      End
   End
   Begin VB.Menu mnuClipBoard 
      Caption         =   "Clip&board"
      Begin VB.Menu mnuCLIP 
         Caption         =   "Copy from Image 1"
         Index           =   0
      End
      Begin VB.Menu mnuCLIP 
         Caption         =   "Copy from image 2"
         Index           =   1
      End
      Begin VB.Menu mnuCLIP 
         Caption         =   "Copy from image 3"
         Index           =   2
      End
      Begin VB.Menu mnuCLIP 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCLIP 
         Caption         =   "Paste to image 1"
         Index           =   4
      End
      Begin VB.Menu mnuCLIP 
         Caption         =   "Paste to image 2"
         Index           =   5
      End
      Begin VB.Menu mnuCLIP 
         Caption         =   "Paste to image 3"
         Index           =   6
      End
      Begin VB.Menu mnuCLIP 
         Caption         =   "(HALFTONE)"
         Index           =   7
      End
      Begin VB.Menu mnuCLIP 
         Caption         =   "(Keep aspect ratio)"
         Index           =   8
      End
      Begin VB.Menu mnuCLIP 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuCLIP 
         Caption         =   "View ClipBoard"
         Index           =   10
      End
      Begin VB.Menu mnuCLIP 
         Caption         =   "Clear Clipboard"
         Index           =   11
      End
   End
   Begin VB.Menu mnuExtractor 
      Caption         =   "&Extractor"
   End
   Begin VB.Menu mnuRotator 
      Caption         =   "&Rotator"
   End
   Begin VB.Menu mnuAlpha 
      Caption         =   "A&lpha"
      Begin VB.Menu mnuAlphaEdit 
         Caption         =   "Alpha edit Image 1"
         Index           =   0
      End
      Begin VB.Menu mnuAlphaEdit 
         Caption         =   "Alpha edit Image 2"
         Index           =   1
      End
      Begin VB.Menu mnuAlphaEdit 
         Caption         =   "Alpha edit Image 3"
         Index           =   2
      End
      Begin VB.Menu mnuAlphaEdit 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuAlphaEdit 
         Caption         =   "Restrict Alpha for Image 1"
         Index           =   4
      End
      Begin VB.Menu mnuAlphaEdit 
         Caption         =   "Restrict Alpha for Image 2"
         Index           =   5
      End
      Begin VB.Menu mnuAlphaEdit 
         Caption         =   "Restrict Alpha for Image 3"
         Index           =   6
      End
      Begin VB.Menu mnuAlphaEdit 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuAlphaEdit 
         Caption         =   "Create Alpha for Image 1"
         Index           =   8
      End
      Begin VB.Menu mnuAlphaEdit 
         Caption         =   "Create Alpha for Image2"
         Index           =   9
      End
      Begin VB.Menu mnuAlphaEdit 
         Caption         =   "Create Alpha for Image 3"
         Index           =   10
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TinyGFX32 with Alpha Edit

' 6/5/07

' 25 Jul 2012
'1. Adjust Sub ShowNine to avoid a mismatch RTE

' 23 Sep 2011
' 1. Faint dividing line to separate Visbility from Grid palette.


' 14 July 2011
' 1. Added a cross shading box in ModDRAW. Varying patterns when image not square!

' 5 July 2011
' 1. Avoid another RTE for less/more colors with 1 pixel size images

' 4 July 2011

'1. Avoid RTE for shifting and flipping 1 pixel size images.

' 7 June 2011
' Added to ModImage.Function Blur(...):-
' If ImageWidth(ImageNum) = 1 Or ImageHeight(ImageNum) = 1 Then Exit Function
' to avoid blurring corners when height or width = 1

' 10 May 2011
' 1. Put ToolTipText on Visibility and Grid color palettes.

' 28 Apr 2011
' Ignore Function CheckSize

' 10 Apr 2011
' 1. Added choice for standard icons only in Extractor

' 24/Mar/10
' 1. Allow 48x48 cursors as well as 32x32.

Option Explicit

' For chm Help file
Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_CLOSE_ALL = &H12

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
(ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Private hwndHelp As Long

' Drawing
Private aMouseDown As Boolean
Private TheButton As Integer
Private DrawCul As Long

'Public
'Private xStart As Single, xend As Single
'Private yStart As Single, yend As Single

Private xleft As Single, xright As Single
Private ytop As Single, ybelow As Single
Private aDropper As Boolean

' For cmdLRUD ' Scroll/Flips
Private picDummy() As Long
Private Temp() As Long

' For W & H Scroll bars
Private aScroll As Boolean

' For Undo/Redo
Private BUSpec$
'Public CurrentGen(0 To 3) As Long
Private MaxGen(0 To 3) As Long
Private MaxGenAllowed(0 To 3)

' For drag/drop
Private Jumper As Long

' For Grid on/off
Private aGrid As Long

'' For showing with other transparent colors
'Private VisColor As Long   ' Test different TColorBGR
Private VisMouse As Boolean

' For Stretch/Shrink & Clipboard
Private aClipBoard As Boolean

Private From_mnuNew As Boolean

' Recent files
Private Const MaxFileCount As Long = 8
Private FileCount As Long
Private NumRecentFiles As Long
Private FileArray$(1 To 8)

Private CommonDialog1 As cOSDialog
Dim Environment As New cEnvironment  'For IDE or ENVIRONMENT



'Public ImageNum As Integer
'Public ImageWidth() As Long
'Public ImageHeight() As Long
'Public OptimizeNumber As Integer
'Public GridMult As Long
'Public GridLinesCul As Long

' Tool Enums in ModImage.bas

Private Sub Form_Initialize()
Dim k As Long
   
   For k = 0 To 2
      MaxGenAllowed(k) = 20    ' Arbitrary number of Undo/Redo files, could be changed
      CurrentGen(k) = 0
      MaxGen(k) = 0
   Next k
   
   ImageNum = 0
   ' Default preferences set in Sub GetINI_Info
'   OptimizeNumber = 0      ' (Optimized saving)
'   ToneNumber = 3          ' (HALFTONE)
'   AspectNumber = 6        ' (Keep Aspect ratio)
'   MarkPixels = 9          ' (Mark transparent pixels)
'   ExtractorBackups = 12   ' (Back up all Extractor images)
   
   GridMult = 8 '  6 for >36 8 <= 36
   MaxWidth = 64
   MaxHeight = 64 ' These can go to 128 but too big for image views
   
   
   'GridLinesCul = RGB(128, 128, 128) ' Now from ini
   ReDim ImageWidth(0 To 2)
   ReDim ImageHeight(0 To 2)
   ' Starting Image Widths & Heights
   ImageWidth(0) = 16
   ImageHeight(0) = 16
   ImageWidth(1) = 32
   ImageHeight(1) = 32
   ImageWidth(2) = 48
   ImageHeight(2) = 48
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   ' Starting font for Text
   aText = False
   TFontname = "Arial"
   TFontsize = 8
   TFontBold = False
   TFontItalic = False

   ' To be sure ?
   aMouseDown = False
   aDropper = False
   aScroll = False
   aSelect = False
   aSelectDrawn = False
   aMoveSEL = False
   aTestCursor = False
   aClipBoard = False
   
   frmRotateLeft = 1200
   frmRotateTop = 1500

End Sub

'#### Form Load ####

Private Sub Form_Load()
Dim bm As BITMAP
Dim k As Long
   
   'LabDate = "14 Aug 2009"
   
   aShow = False

   Me.Width = 10935  ' Comment to see off form picboxes
   Me.Height = 9210 '8760
   
   AppPathSpec$ = App.Path
   If Right$(AppPathSpec$, 1) <> "\" Then AppPathSpec$ = AppPathSpec$ & "\"
   CurrPath$ = AppPathSpec$
   CPath$ = AppPathSpec$
   
   IniTitle$ = "GFX"
   FileCount = 0
   aAlphaRestricted = False   ' Alpha ON Unchecked
   
   GetINI_Info
   
   aGrid = 2
   chkGridOnOff.Value = Checked
   aGrid = 1
   
   MarkPixels = MarkPixels - 1
   imTMarker_Click
   LocateCtrls_Init
   
   LabErase.Visible = False
   
   ImageNum = 0
   
   Show
   
   aShow = True
   
   picC.Cls
   
   GetObject Image, Len(bm), bm
   If bm.bmBitsPixel < 24 Then
      MsgBox Str$(bm.bmBitsPixel) & "bpp. Sorry needs True Color Setting (24 or 32 bpp)", vbInformation, "tinyGFX"
      Unload Me
      End
   End If
   
   Shape1.Visible = False
   
   For ImageNum = 2 To 0 Step -1
      cmdUNDO(ImageNum).Enabled = False
      cmdRedo(ImageNum).Enabled = False
      cmdClear(ImageNum).Enabled = False
      mnuReloadImage(ImageNum).Enabled = False
      cmdReload(ImageNum).Enabled = False
      mnuAlphaEdit(ImageNum).Enabled = False
      mnuAlphaEdit(ImageNum + 4).Enabled = False
      ImBPP(ImageNum) = 1
   Next ImageNum
   
   ' Create Alphas
   mnuAlphaEdit(8).Enabled = True
   mnuAlphaEdit(9).Enabled = True
   mnuAlphaEdit(10).Enabled = True
   
   ImageNum = 0
   ' Start with Dots
   optTools(0).Value = True
   optTools_MouseUp 0, 1, 0, 0, 0
   LabTool = "Dots"
   ' Init Color count
   For ImageNum = 2 To 0 Step -1
      GetTheBitsLong picSmall(ImageNum)
      k = ColorCount(picSmallDATA())
      LabNumColors(ImageNum) = "Colors =" & Str$(k)
   Next ImageNum
   ImageNum = 0
   
   SaveSystemCursor
   
   Caption = " Tiny GFX32"
   
   For k = 0 To 2
      KILLSAVS k
   Next k

   
   TURNSELECTOFF
   
   picSmall(0).OLEDropMode = 1
   picSmall(1).OLEDropMode = 1
   picSmall(2).OLEDropMode = 1

   If Command$ <> "" Then   ' Loading pic on to exe
      CommandLineIN
   End If
   
   From_mnuNew = False
   
   cmdSelect_MouseUp 0, 1, 0, 0, 0
   
   Form1.KeyPreview = True

End Sub

Private Sub CommandLineIN()
Dim Infile$
      
   If Left$(Command$, 1) = Chr(34) Then ' Strip off quotes
       Infile$ = Mid$(Command$, 2, Len(Command$) - 2)
   Else
       Infile$ = Command$
   End If
   FileSpec$ = Infile$
   CPath$ = FileSpec$
   SavePath$ = FileSpec$
   Jumper = 1
   mnuOpenintoImage_Click 0
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Index As Integer
   Index = KeyCode - 49
   If Index >= 0 And Index <= 2 Then
      Select Case LabMAC.Caption
      Case "Masks"
         cmdMask_MouseUp 1, 0, 0, 0
         cmdSelect_MouseUp Index, 1, 0, 0, 0
         cmdMask_MouseDown 1, 0, 0, 0
      Case "Alphas"
         cmdAlpha_MouseUp 1, 0, 0, 0
         cmdSelect_MouseUp Index, 1, 0, 0, 0
         cmdAlpha_MouseDown 1, 0, 0, 0
      Case "Original colors"
         cmdORGCul_MouseUp 1, 0, 0, 0
         cmdSelect_MouseUp Index, 1, 0, 0, 0
         cmdORGCul_MouseDown 1, 0, 0, 0
      Case "Images", "No Alphas"
         cmdSelect_MouseUp Index, 1, 0, 0, 0
      End Select
   End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
End Sub

Private Sub imTMarker_Click()
 MarkPixels = MarkPixels + 1
 If MarkPixels > 11 Then MarkPixels = 9
 
 mnuPreferences(9).Checked = False
 mnuPreferences(10).Checked = False
 mnuPreferences(11).Checked = False
 
 Select Case MarkPixels
 Case 11    'Unmarked 101
   mnuPreferences(11).Checked = True
   imTMarker.Picture = LoadResPicture(101, vbResBitmap)
 Case 10    ' Diagonal 102
   mnuPreferences(10).Checked = True
   imTMarker.Picture = LoadResPicture(102, vbResBitmap)
 Case 9     'Square   103
   mnuPreferences(9).Checked = True
   imTMarker.Picture = LoadResPicture(103, vbResBitmap)
 End Select
 DrawGrid
End Sub

Private Sub mnuAlpha_Click()
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
End Sub

Private Sub mnuAlphaEdit_Click(Index As Integer)
Dim BMIH As BITMAPINFOHEADER
Dim ix As Long, iy As Long
Dim Cul As Long
Dim ABYTE As Byte
Dim svImBPP As Long


   If aSelectDrawn Then
      CancelSelection
   End If
   
   ' 0,1,2
   Select Case Index
   Case 0, 1, 2   ' Alpha Edit
      If ImBPP(Index) = 32 Then
         cmdSelect_MouseUp CInt(Index), 1, 0, 0, 0
         EditCreate = 0
         'BackUp ImageNum
         frmAlphaEdit.Show vbModal
         
         If aAlphaEdit Then
            ' TURNSELECTOFF  resets picSmall(ImqageNum) from picSmallBU
            BitBlt picSmallBU.hdc, 0, 0, picSmallBU.Width, picSmallBU.Height, _
                picSmall(ImageNum).hdc, 0, 0, vbSrcCopy     ' dest,src
            picSmallBU.Picture = picSmallBU.Image
            
            cmdSelect_MouseUp CInt(Index), 1, 0, 0, 0
            DrawGrid
            BackUp ImageNum
            LabSpec = "(Bpp =" & Str$(ImBPP(ImageNum)) & ")"
         End If
      End If
   Case 3   ' Break
   Case 4, 5, 6 ' Restrict Alpha for Images -4 = 0,1,2
      If ImBPP(Index - 4) = 32 Then
         ImageNum = (Index - 4)
         cmdSelect_MouseUp CInt(ImageNum), 1, 0, 0, 0
         EditCreate = 0
         ' Make ImageNum opaque
         picSmallBU.Width = ImageWidth(ImageNum)
         picSmallBU.Height = ImageHeight(ImageNum)
         picSmallBU.Picture = LoadPicture
         
         For iy = 0 To ImageHeight(ImageNum) - 1
         For ix = 0 To ImageWidth(ImageNum) - 1
            Select Case ImageNum
            Case 0
               If DATACUL0(3, ix, iy) <> 255 Then
                  'DATACUL0(3, ix, iy) = 0
                  picSmallBU.PSet (ix, ImageHeight(ImageNum) - 1 - iy), TColorBGR
               Else
                  picSmallBU.PSet (ix, ImageHeight(ImageNum) - 1 - iy), RGB(DATACUL0(2, ix, iy), DATACUL0(1, ix, iy), DATACUL0(0, ix, iy))
               End If
            Case 1
               If DATACUL1(3, ix, iy) <> 255 Then
                  'DATACUL1(3, ix, iy) = 0
                  picSmallBU.PSet (ix, ImageHeight(ImageNum) - 1 - iy), TColorBGR
               Else
                  picSmallBU.PSet (ix, ImageHeight(ImageNum) - 1 - iy), RGB(DATACUL1(2, ix, iy), DATACUL1(1, ix, iy), DATACUL1(0, ix, iy))
               End If
            Case 2
               If DATACUL2(3, ix, iy) <> 255 Then
                  'DATACUL2(3, ix, iy) = 0
                  picSmallBU.PSet (ix, ImageHeight(ImageNum) - 1 - iy), TColorBGR
               Else
                  picSmallBU.PSet (ix, ImageHeight(ImageNum) - 1 - iy), RGB(DATACUL2(2, ix, iy), DATACUL2(1, ix, iy), DATACUL2(0, ix, iy))
               End If
            End Select
         Next ix
         Next iy
         
         BitBlt picSmall(ImageNum).hdc, 0, 0, picSmall(ImageNum).Width, picSmall(ImageNum).Height, _
             picSmallBU.hdc, 0, 0, vbSrcCopy     ' dest,src
         picSmall(ImageNum).Picture = picSmall(ImageNum).Image
         
         ImBPP(ImageNum) = 24
         DrawGrid
         ShowTheColorCount ImageNum
         LabSpec = "(Bpp =" & Str$(ImBPP(ImageNum)) & ")"
         mnuAlphaEdit(ImageNum + 4).Enabled = False
         BackUp ImageNum
         NameSpec$(ImageNum) = ""
      End If
   Case 7   ' Break
   Case 8, 9, 10   ' Create Alpha -8 = 0,1,2
      If ImBPP(Index - 8) = 32 Then
         MsgBox "32 bpp image with Alpha already ?!", vbInformation, "Creating Alpha image"
         cmdSelect_MouseUp Index - 8, 1, 0, 0, 0
         Exit Sub
      End If
      
      If IsBlank(Index - 8) Then
         MsgBox "Blank image" & Str$(Index - 7) & " ?!", vbInformation, "Creating Alpha image"
         cmdSelect_MouseUp Index - 8, 1, 0, 0, 0
         Exit Sub
      End If
      
      
      ImageNum = (Index - 8)
      If ImBPP(ImageNum) < 32 Then
         svImBPP = ImBPP(ImageNum)
         cmdSelect_MouseUp CInt(ImageNum), 1, 0, 0, 0
         EditCreate = 1
         ' make picDATA DATACUL#
         '      Alpha# = Mask#
   
         aDIBError = False
         With BMIH
            .biSize = 40
            .biPlanes = 1
            .biWidth = ImageWidth(ImageNum)
            .biHeight = ImageHeight(ImageNum)
            .biBitCount = 32
            '.biSizeImage = 4 * W * H
         End With
         SetStretchBltMode picSmall(ImageNum).hdc, COLORONCOLOR
         ReDim picSmallDATA(0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         If GetDIBits(Form1.hdc, picSmall(ImageNum).Image, 0, ImageHeight(ImageNum), picSmallDATA(0, 0), BMIH, 0) = 0 Then
            MsgBox "DIB ERROR   ", vbCritical, "Making picSmallDATA()"
            Exit Sub
         End If
               
         ReDim pMaskData(0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         picTemp(0).Width = ImageWidth(ImageNum)
         picTemp(0).Height = ImageHeight(ImageNum)
         ShowMask picSmall(ImageNum), picTemp(0)
         With BMIH
            .biSize = 40
            .biPlanes = 1
            .biWidth = ImageWidth(ImageNum)
            .biHeight = ImageHeight(ImageNum)
            .biBitCount = 32
            '.biSizeImage = 4 * W * H
         End With
         If GetDIBits(Form1.hdc, picTemp(0).Image, 0, ImageHeight(ImageNum), pMaskData(0, 0), BMIH, 0) = 0 Then
            MsgBox "DIB ERROR   ", vbCritical, "Making Mask Data"
            Exit Sub
         End If
         
         ' Have picSmallDATA() & picTemp(0) (ie picMask)
         Select Case ImageNum
         Case 0
            ReDim DATACUL0(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         Case 1
            ReDim DATACUL1(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         Case 2
            ReDim DATACUL2(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         End Select
         For iy = 0 To ImageHeight(ImageNum) - 1
         For ix = 0 To ImageWidth(ImageNum) - 1
            If picTemp(0).Point(ix, iy) = vbWhite Then ABYTE = 0 Else ABYTE = 255
     'If ABYTE = 255 Then Stop
            Cul = picSmall(ImageNum).Point(ix, iy)
            SetAlphas ix, iy, ABYTE
            CulTo32bppArrays Cul, ix, iy, ABYTE
         Next ix
         Next iy
         
         ImBPP(ImageNum) = 32 ' NOW
         
         frmAlphaEdit.Show vbModal
         
         If aAlphaEdit Then
            ImBPP(ImageNum) = 32
            cmdSelect_MouseUp CInt(ImageNum), 1, 0, 0, 0
            DrawGrid
            BackUp ImageNum
            LabSpec = "(Bpp =" & Str$(ImBPP(ImageNum)) & ")"
         Else ' Cancelled
            ImBPP(ImageNum) = svImBPP
            LabSpec = "(Bpp =" & Str$(ImBPP(ImageNum)) & ")"
         End If
      Else
         Exit Sub
      End If
   End Select
   ShowTheColorCount ImageNum
End Sub

'#### Clipboard ####
Private Sub mnuClipBoard_Click()
   If aSelect Then TURNSELECTOFF
End Sub

Private Sub mnuCLIP_Click(Index As Integer)
Dim BMIH As BITMAPINFOHEADER
Dim sw As Long, sh As Long
Dim ImNum As Long
Dim Ret As Long
Dim f$
' For aspect ratio for Drawn rectangle Capture and Clipboard pasting
Dim CCbm As BITMAP
Dim CCW As Long, CCH As Long
Dim BitsPP As Long
Dim CCpw As Long
Dim CCph As Long
Dim zAspectCC As Single

   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
   
   Select Case Index
   Case 0, 1, 2 ' Copy from image 1,2,3 (0,1,2) to ClipBoard
      ImNum = CLng(Index)
      Clipboard.Clear
      Clipboard.SetData picSmall(ImNum).Image
      aClipBoard = True
   Case 3   ' break
   Case 4, 5, 6 ' Paste to image 1,2,3 (0,1,2)
      On Error GoTo PasteERROR
      ImNum = CLng(Index - 4)
      
      cmdSelect_MouseUp CInt(ImNum), 1, 0, 0, 0
      If CurrentGen(ImNum) = 0 Then
         BackUp ImNum
      End If
      '' Get Clipboard image size
      ''Clipboard.Clear   ' In IDE there's a Clipboard image put in by VB
      GetObject Clipboard.GetData(vbCFBitmap), Len(CCbm), CCbm
      CCW = CCbm.bmWidth
      CCH = CCbm.bmHeight
      BitsPP = CCbm.bmBitsPixel   ' Could calc image size on Clipboard
      If CCW = 0 Or CCH = 0 Then
         MsgBox "No image on Clipboard", vbInformation, "Clipboard"
         Exit Sub
      End If
      
      '------------------------------
      If AspectNumber = 6 Then   ' Keep aspect ratio
         zAspectCC = CCW / CCH
         If CCW >= CCH Then  ' zASpectCC >= 1
            CCpw = ImageWidth(ImNum)
            CCph = ImageWidth(ImNum) / zAspectCC
         Else  ' W < H   ' zAspect < 1
            CCph = ImageHeight(ImNum)
            CCpw = ImageHeight(ImNum) * zAspectCC
         End If
         If CCpw < 8 Or CCph < 8 Then
            MsgBox "Keeping aspect makes width or height < 8", vbInformation, "Clipboard"
            Exit Sub
         End If
         picSmallBU.Picture = LoadPicture
         picSmall(ImNum).Picture = LoadPicture
         picSmallBU.Width = CCpw
         picSmallBU.Height = CCph
         picSmall(ImNum).Width = CCpw
         picSmall(ImNum).Height = CCph
         ImageWidth(ImNum) = CCpw
         ImageHeight(ImNum) = CCph
      End If
      '------------------------------
      
      picSmallBU.Picture = Clipboard.GetData(vbCFBitmap)
      sw = picSmallBU.Width
      sh = picSmallBU.Height
      
      With BMIH
         .biSize = 40
         .biPlanes = 1
         .biWidth = sw
         .biHeight = sh
         .biBitCount = 32
         '.biSizeImage = 4 * W * H
      End With
      ReDim picSmallDATA(0 To sw - 1, 0 To sh - 1)
      If GetDIBits(Form1.hdc, picSmallBU.Image, 0, sh, picSmallDATA(0, 0), BMIH, 0) = 0 Then
         MsgBox "DIB Error", vbCritical, "Clipboard"
         picSmallBU.Picture = LoadPicture
         picSmallBU.Width = 4
         picSmallBU.Height = 4
         Exit Sub
      End If
      
      If ToneNumber = 3 Then  ' Halftone
         SetStretchBltMode picSmall(ImNum).hdc, 4
      Else  ' Coloroncolor
         SetStretchBltMode picSmall(ImNum).hdc, 3
      End If

      StretchDIBits picSmall(ImNum).hdc, 0, 0, picSmall(ImNum).Width, picSmall(ImNum).Height, _
      0, 0, sw, sh, picSmallDATA(0, 0), BMIH, 0, vbSrcCopy
      picSmall(ImNum).Picture = picSmall(ImNum).Image
      
      picSmallBU.Picture = LoadPicture
      picSmallBU.Width = 4
      picSmallBU.Height = 4
      
'      cmdSelect_MouseUp CInt(ImNum), 1, 0, 0, 0
      BackUp ImNum
      ImBPP(ImNum) = 0
      ShowTheColorCount ImNum
      NameSpec$(ImageNum) = ""
      ImFileSpec$(ImageNum) = ""
      cmdSelect_MouseUp CInt(ImNum), 1, 0, 0, 0
      
      Case 7   ' (HALFTONE Pasting) or (COLORONCOLOR Pasting)
      Case 8   ' (Keep or Ignore aspect ratio)
      Case 9   ' Break
      Case 10  ' View Clipboard
         f$ = Space$(255)
         Ret = GetSystemDirectory(f$, 255)
         f$ = Left$(f$, Ret) & "\clipbrd.exe"
         If FileExists(f$) Then
            Shell f$, vbMaximizedFocus
         Else
            f$ = Space$(255)
            Ret = GetWindowsDirectory(f$, 255)
            f$ = Left$(f$, Ret) & "\clipbrd.exe"
            If FileExists(f$) Then
               Shell f$, vbMaximizedFocus
            Else
               MsgBox f$ & "  " & vbCrLf & " Not there!", vbInformation, "View Clipboard"
            End If
         End If
      Case 11   ' Clear Clipboard
         Clipboard.Clear
         aClipBoard = False
   End Select
   Exit Sub
'===========
PasteERROR:
MsgBox "Error when pasting", vbCritical, "PASTING"
End Sub

'#### Screen Capture ####

Private Sub mnuCapture_Click()
   If aSelect Then TURNSELECTOFF
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

End Sub

Private Sub mnuCaptureToImage_Click(Index As Integer)
   mnuCaptureToImage(4).Checked = False
   mnuCaptureToImage(5).Checked = False
   mnuCaptureToImage(6).Checked = False
   Select Case MagCAP
   Case 1: mnuCaptureToImage(4).Checked = True
   Case 2: mnuCaptureToImage(5).Checked = True
   Case 4: mnuCaptureToImage(6).Checked = True
   End Select
   
   Select Case Index
   Case 0, 1, 2 ' Capture
   Case 3       ' Break
   Case 4: MagCAP = 1
      mnuCaptureToImage(4).Checked = True
      mnuCaptureToImage(5).Checked = False
      mnuCaptureToImage(6).Checked = False
      Exit Sub
   Case 5: MagCAP = 2
      mnuCaptureToImage(5).Checked = True
      mnuCaptureToImage(4).Checked = False
      mnuCaptureToImage(6).Checked = False
      Exit Sub
   Case 6: MagCAP = 4
      mnuCaptureToImage(6).Checked = True
      mnuCaptureToImage(4).Checked = False
      mnuCaptureToImage(5).Checked = False
      Exit Sub
      
   Case 7, 8 ' Breaks
   Case 9     ' Drawn rectangle to image 1
   Case 10    ' Drawn rectangle to image 2
   Case 11    ' Drawn rectangle to image 3
   Case 12     '(Keep or ignore zAspect)
   End Select
   
   If MagCAP = 0 Then   ' Shouldn't happen ?
      MagCAP = 1
      mnuCaptureToImage(4).Checked = True
      mnuCaptureToImage(5).Checked = False
      mnuCaptureToImage(6).Checked = False
   End If
   
   If aSelect Then TURNSELECTOFF ' Already done ?
   
   Select Case Index
   Case 0, 1, 2
      ImageNum = Index
      If CurrentGen(ImageNum) = 0 Then
         If Tools <> [Dropper] And Tools <> -1 Then
            BackUp ImageNum
         End If
      End If
      
      Form1.WindowState = vbMinimized
      DoEvents
   
      frmCAP.Show vbModal
   
   Case 9, 10, 11
      ImageNum = Index - 9
      If CurrentGen(ImageNum) = 0 Then
         If Tools <> [Dropper] And Tools <> -1 Then
            BackUp ImageNum
         End If
      End If
      
      Form1.WindowState = vbMinimized
      DoEvents
      frmCAP2.Show vbModal
      
   End Select
   
   Form1.WindowState = vbNormal
   Form1.SetFocus
   ShowTheColorCount ImageNum
   NameSpec$(ImageNum) = ""
   ImFileSpec$(ImageNum) = ""
   cmdSelect_MouseUp CInt(ImageNum), 1, 0, 0, 0
   If aCAP Then BackUp CLng(ImageNum)
End Sub

Private Sub mnuExtractor_Click()
Dim i As Integer
   
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
   
   frmExtractor.Show vbModal
   ' Select all ending up with ImageNum = 0
   
   If aAlphaRestricted Then
      mnuPreferences(16).Checked = True
   Else
      mnuPreferences(16).Checked = False
   End If
   For i = 2 To 0 Step -1
      cmdSelect_MouseUp i, 1, 0, 0, 0   ' Also sets ImageNum = 2,1 lastly 0
      ShowTheColorCount CLng(i)
   Next i
   BackUp ImageNum
End Sub

Private Sub mnuNew_Click()
Dim i As Integer
Dim resp As Long
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
   
   For i = 2 To 0 Step -1
      If cmdClear(i).Enabled = True Then Exit For
   Next i
   If i = -1 Then Exit Sub  'ie all Disabled
   
   resp = MsgBox("Clear all images and" & vbCrLf & _
          "clear any Undos/Redos." & vbCrLf & "Sure?", vbQuestion + vbYesNo, "Clear All")
   If resp = vbNo Then Exit Sub
   
   From_mnuNew = True
   For i = 2 To 0 Step -1
      If cmdClear(i).Enabled = True Then
         cmdClear_Click i
      End If
   Next i
   For i = 2 To 0 Step -1
      cmdClear(i).Enabled = False
   Next i
   From_mnuNew = False
End Sub

Private Sub mnuPalette_Click(Index As Integer)
   Select Case Index
   Case 0   ' Break
   Case 1   ' Open palette bmp
      frmPalette.Show 0

   End Select
End Sub

Private Sub mnuRecentFiles_Click(Index As Integer)
   
   If NumRecentFiles > 0 Then
      'FileSpec$ = mnuRecentFiles(Index).Caption
      FileSpec$ = FileArray(Index)
      
      If Not FileExists(FileSpec$) Then
         MsgBox FileSpec$ & vbCrLf & "File not found!  ", vbOKOnly + vbInformation, "Recent File"
         Exit Sub
      End If
      
      frmGetImNum.Show 1
      
      If TempImageNum = -1 Then
         FileSpec$ = ""
      End If
         
      ProcessInFile
   End If
End Sub

Private Sub mnuReload_Click()
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

End Sub

Private Sub mnuRotator_Click()
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
   If aMoveSEL Then
     TURNSELECTOFF
   End If
   If Not Set_LOHI_XY() Then
     TURNSELECTOFF
     Exit Sub
   End If
     
   'FillTransparentAreas  'vbMagenta default black
   
   frmRotator.Show vbModal
   
   If aTransfer Then
      If ImBPP(ImageNum) = 32 Then
         Reconcile picSmall(ImageNum), ImageNum  ' Ensure Alpha and Image aligned
      End If
   End If
   
   cmdSelect_MouseUp CInt(ImageNum), 1, 0, 0, 0
   ShowTheColorCount ImageNum
   If aTransfer Then
      BackUp ImageNum
   End If
End Sub

Private Sub mnuSaveBMP_Click()
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

End Sub

Private Sub mnuSaveCUR_Click()
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

End Sub

Private Sub mnuSaveICO_Click()
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

End Sub

Private Sub picC_Click()
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

End Sub

Private Sub picSmall_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
   cmdSelect_MouseUp Index, 1, 0, 0, 0
End Sub

'#### Drag/Drop ####

Private Sub picSmall_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Infile$
   
   On Error GoTo NoPic
   Infile$ = Data.Files(1)
   
      FileSpec$ = Infile$
      Jumper = 1
      mnuOpenintoImage_Click Index
   
   On Error GoTo 0
   Exit Sub
'=========
NoPic:
  MsgBox Err.Description & vbCrLf & vbCrLf & "Drag drop error", vbExclamation, "ERR# " & Err
End Sub

'#### Tools & Drawing ####

Private Sub ToolsCaption()
   Select Case Tools
   Case -1:              LabTool = "None"
   Case [Dots]:          LabTool = "Dots"
   Case [Lines]:         LabTool = "Line"
   Case [Boxes]:         LabTool = "Box"
   Case [BoxesSolid]:    LabTool = "Solid Box"
   Case [BoxesFilled]:   LabTool = "Filled box"
   Case [Ellipse]:       LabTool = "Ellipse"
   Case [EllipseSolid]:  LabTool = "Solid Ellipse"
   Case [EllipseFilled]: LabTool = "Filled ellipse"
   Case [Fill]:          LabTool = "Fill"
   Case [Dropper]:       LabTool = "Dropper"
   Case [Rot90CW]:       LabTool = "Rot 90 CW"
   Case [Rot90ACW]:      LabTool = "Rot 90 ACW"
   Case [Text]:          LabTool = "Text"
   
   Case [BoxCenShade]:       LabTool = "Box center shading"
   Case [BoxVertShade]:      LabTool = "Box vert shading"
   Case [BoxHorzShade]:      LabTool = "Box horz shading"
   Case [EllipseCenShade]:   LabTool = "Ellipse center shading"
   Case [EllipseVertShade]:  LabTool = "Ellipse vert shading"
   Case [EllipseHorzShade]:  LabTool = "Ellipse horz shading"
   Case [BoxDiagTLBRShade]:  LabTool = "Box diag \ shading"
   Case [BoxDiagTLBRShade]:  LabTool = "Box vert center shading"
   Case [BoxDiagBLTRShade]:  LabTool = "Box diag / shading"
   Case [BoxVertCenShade]:  LabTool = "Box vert center shading"
   Case [BoxHorzCenShade]:  LabTool = "Box horz center shading"
   Case [BoxHVCenShade]:    LabTool = "Box Horz && Vert shading"
   
   Case [BlurIm]:        LabTool = "Blur"
   Case [LeftColorReliefIm]:  LabTool = "Left color-Relief"
   Case [InvertIm]:      LabTool = "Invert"
   Case [BrighterIm]:    LabTool = "Brighter"
   Case [DarkerIm]:      LabTool = "Darker"
   Case [BorderIm]:      LabTool = "Border"
   Case [ReplaceLbyR]:   LabTool = "Replace L by R"
   
   Case [SelectON]:      LabTool = "Select ON"
   Case [MoveSEL]:       LabTool = "Move selection"
   Case [MoveCOPY]:      LabTool = "Copy selection"
   
   Case [Tools2Blur]:      LabTool = "Blur Tool"
   Case [Tools2Grey]:      LabTool = "Grey Tool"
   Case [Tools2Invert]:    LabTool = "Speckle Invert Tool"
   Case [Tools2Bright]:    LabTool = "Brighter Tool"
   Case [Tools2Dark]:      LabTool = "Darker Tool"
   Case [Tools2Checker]:   LabTool = "Checkerboard Tool"
   Case [Tools2HLines]:    LabTool = "Horz Lines Tool"
   Case [Tools2VLines]:    LabTool = "Vert Lines Tool"
   Case [Tools2Random]:    LabTool = "Random Tool"
   Case [Tools2MoreBlue]:  LabTool = "More blue"
   Case [Tools2LessBlue]:  LabTool = "Less blue"
   Case [Tools2MoreGreen]: LabTool = "More green"
   Case [Tools2LessGreen]: LabTool = "Less green"
   Case [Tools2MoreRed]:   LabTool = "More red"
   Case [Tools2LessRed]:   LabTool = "Less red"
   
   Case [MirrorL]:   LabTool = "Mirror left"
   Case [MirrorR]:   LabTool = "Mirror right"
   Case [MirrorT]:   LabTool = "Mirror top"
   Case [MirrorB]:   LabTool = "Mirror bottom"
   Case [ReflectLeft]:   LabTool = "Reflect left"
   Case [ReflectRight]:  LabTool = "Reflect right"
   Case [ReflectTop]:    LabTool = "Reflect top"
   Case [ReflectBottom]: LabTool = "Reflect bottom"

   End Select
End Sub

Private Sub optTools_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Long
Dim BMIH As BITMAPINFOHEADER
' For [MoveSEL]
Dim xSrc As Single, ySrc As Single
Dim sw As Single, sh As Single
Dim ix As Long, iy As Long
Dim Cul As Long
Dim ab As Boolean
Dim ABYTE As Byte

   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

   
   ' Switch off OptTools2s [Tools2Blur] to [ReflectBottom]
   For k = 0 To optTools2.Count - 1
      optTools2(k).Value = False
   Next k
      ' Switch off all OptTools
   For k = 0 To optTools.Count - 1
      optTools(k).Value = False
   Next k
   
   ' Set Tools
   Tools = Index
   ToolsCaption
   optTools(Index).Value = True
   
   picPANEL.SetFocus

   shpCirc.Visible = False
   
   optTools([MoveSEL]).Enabled = False
   optTools([MoveCOPY]).Enabled = False
   
   Select Case Tools
   Case [Dots] To [Text], [BoxCenShade] To [BoxHVCenShade]
      LOX = 0: HIX = ImageWidth(ImageNum) - 1
      LOY = 0: HIY = ImageHeight(ImageNum) - 1
      If aSelect Then   ' bar Effects Tools
         TURNSELECTOFF  ' does Dots
         Tools = Index  ' turn actual tool back on
         ToolsCaption
      End If
   End Select
   
   If CurrentGen(ImageNum) = 0 Then ' IE an empty image, no back ups
      Select Case Tools
      Case [Rot90CW], [Rot90ACW], [BorderIm] ' Can act on an empty image
         BackUp ImageNum
      Case [BlurIm] To [DarkerIm], [ReplaceLbyR] ' No point in backing up
         TURNSELECTOFF
         Tools = [Dots]
         optTools(0).Value = True
         ToolsCaption
         Exit Sub
      End Select
   End If

   Select Case Tools 'Index
   Case [Dropper]
      LabTool = "Dropper"
      aDropper = True
      If Environment = EnvironIDE Then
         picPANEL.MouseIcon = LoadResPicture("DROPPER", vbResCursor)
      Else
         picPANEL.MouseIcon = LoadResPicture("DROPPER32", vbResCursor)
      End If
      picPANEL.MousePointer = vbCustom
   Case [Fill]
      picPANEL.MouseIcon = LoadResPicture("FILLER", vbResCursor)
      picPANEL.MousePointer = vbCustom
   Case Else
      picPANEL.MouseIcon = LoadResPicture("PEN", vbResCursor)
      picPANEL.MousePointer = vbCustom
   End Select
   
   Select Case Tools 'Index
   
   Case [Rot90CW], [Rot90ACW]  ' Rot90 Clockwise, Rot90 Anti-Clockwise
         
      If aMoveSEL Then
         TURNSELECTOFF
         Exit Sub
      End If
      ' Input:  pic = picSmall(ImageNum) or picSmallTest
      GetTheBitsLong picSmall(ImageNum)
      ' Output: picSmallDATA(0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
      ' Input: picSmallDATA()
      Rotate90 Index
      ' Output: picSmallDATA(0 To ImageHeight(ImageNum) - 1, 0 To ImageWidth(ImageNum) - 1)
      
      ' Swap W & H values
      k = ImageHeight(ImageNum)
      ImageHeight(ImageNum) = ImageWidth(ImageNum)
      ImageWidth(ImageNum) = k
      picSmall(ImageNum).Height = ImageHeight(ImageNum)
      picSmall(ImageNum).Width = ImageWidth(ImageNum)
      
      With BMIH
         .biSize = 40
         .biPlanes = 1
         .biWidth = ImageWidth(ImageNum)
         .biHeight = ImageHeight(ImageNum)
         .biBitCount = 32
      End With
      If SetDIBits(picSmall(ImageNum).hdc, picSmall(ImageNum).Image, _
         0, ImageHeight(ImageNum), picSmallDATA(0, 0), BMIH, 0) = 0 Then
         MsgBox "SetDIBits Error", vbCritical, "optTools"
      End If
      picSmall(ImageNum).Picture = picSmall(ImageNum).Image
      ' Set new scroll values
      scrWidth.Value = ImageWidth(ImageNum)
      scrHeight.Value = 65 - ImageHeight(ImageNum)
      DrawGrid
      ' Lift Rotate buttons
      optTools(Index).Value = False
      BackUp ImageNum
      ShowTheColorCount ImageNum
      ToolsCaption
   
   Case [Text]  ' Text
      aText = True
      optTools(11).Value = False
      
      frmText.Show vbModal
      
      If aTextOK Then
         DrawGrid
         BackUp ImageNum
      End If
      ShowTheColorCount ImageNum
      Tools = [Dots]
      optTools(0).Value = True
      ToolsCaption
   
   ' ==== EFFECTS ===============
   Case [BlurIm]  ' Blur
      If Not Set_LOHI_XY Then Exit Sub
      If aMoveSEL Then
         TURNSELECTOFF
         Tools = Index
         ToolsCaption
         optTools(Index).Value = True
         Exit Sub
      End If
      
      If Blur(picSmall(ImageNum), 0) Then
         DisplayEffects Index    ' picDATAREG() to picSMALL(ImageNum)
         
         If aSelectDrawn Then ' So that picPANEL_MouseDown
            aEffects = True   ' can react to keep or cancel
            Tools = [SelectON]
         Else
            BackUp ImageNum
         End If
      Else
         TURNSELECTOFF
         Tools = [Dots]
         optTools(0).Value = True
         ToolsCaption
         Exit Sub
      End If
      
   Case [LeftColorReliefIm]
      If Not Set_LOHI_XY Then Exit Sub
      If aMoveSEL Then
         TURNSELECTOFF
         Tools = Index
         ToolsCaption
         optTools(Index).Value = True
         Exit Sub
      End If
      
      If LeftColorRelief(picSmall(ImageNum)) Then
         DisplayEffects Index
      
         If aSelectDrawn Then
            aEffects = True
            Tools = [SelectON]
         Else
            BackUp ImageNum
         End If
      Else
         TURNSELECTOFF
         Tools = [Dots]
         optTools(0).Value = True
         ToolsCaption
         Exit Sub
      End If
   
   Case [InvertIm]
      If Not Set_LOHI_XY Then Exit Sub
      If aMoveSEL Then
         TURNSELECTOFF
         Tools = Index
         ToolsCaption
         optTools(Index).Value = True
         Exit Sub
      End If
      
      If Invert(picSmall(ImageNum)) Then
         DisplayEffects Index
         
         If aSelectDrawn Then
            aEffects = True
            Tools = [SelectON]
         Else
            BackUp ImageNum
         End If
      Else
         TURNSELECTOFF
         Tools = [Dots]
         optTools(0).Value = True
         ToolsCaption
         Exit Sub
      End If
   
   Case [BrighterIm], [DarkerIm]
      If Not Set_LOHI_XY Then Exit Sub
      If aMoveSEL Then
         TURNSELECTOFF
         Tools = Index
         ToolsCaption
         optTools(Index).Value = True
         Exit Sub
      End If
      ab = False
      If Index = [BrighterIm] Then
         If BrighterDarker(picSmall(ImageNum)) Then ab = True
      Else  ' Index = [DarkerIm]
         If BrighterDarker(picSmall(ImageNum)) Then ab = True
      End If
      If ab Then
         DisplayEffects Index
      
         If aSelectDrawn Then
            aEffects = True
            Tools = [SelectON]
         Else
            BackUp ImageNum
         End If
      Else
         TURNSELECTOFF
         Tools = [Dots]
         optTools(0).Value = True
         ToolsCaption
         Exit Sub
      End If
   Case [BorderIm]
      If aMoveSEL Then
         TURNSELECTOFF
         Exit Sub
      End If
      If Not Set_LOHI_XY Then Exit Sub
      
      Border picSmall(ImageNum)
      DisplayEffects Index
      ShowTheColorCount ImageNum
      
      BackUp ImageNum
      
      If aSelectDrawn Then
         CancelSelection
      End If
         
      Tools = [Dots]
      optTools(0).Value = True
      ToolsCaption
      
   Case [ReplaceLbyR]  ' Replace Left by Right color
      If aMoveSEL Then
         TURNSELECTOFF
         Exit Sub
      End If
      If Not Set_LOHI_XY Then Exit Sub
      
      ReplaceLbyRColor picSmall(ImageNum)
      DisplayEffects Index
      ShowTheColorCount ImageNum
      BackUp ImageNum
      
      If aSelectDrawn Then
         CancelSelection
      End If
      Tools = [Dots]
      optTools(0).Value = True
      ToolsCaption
   
   ' ==== EFFECTS END ===========
      
   ' ==== SELECTION ===========
   Case [SelectON]
      If Not aSelect Then ' Switch on
         If IsBlank(CInt(ImageNum)) Then
            picSmallBU.Picture = LoadPicture
         'If CurrentGen(ImageNum) = 0 Then
            TURNSELECTOFF
            Exit Sub
         End If
         ' Backup DATACUL#()
         If ImBPP(ImageNum) = 32 Then
            ReDim DATACULBU(0 To 3, ImageWidth(ImageNum) - 1, ImageHeight(ImageNum) - 1)
            Select Case ImageNum
            Case 0: FILL3D DATACULBU(), DATACUL0()
            Case 1: FILL3D DATACULBU(), DATACUL1()
            Case 2: FILL3D DATACULBU(), DATACUL2()
            End Select
         End If
         
         picPANEL.MouseIcon = LoadResPicture("CROSS", vbResCursor)
         picPANEL.MousePointer = vbCustom
         ToolsCaption
         picSmallToPicSmallBU
         aSelect = True
      Else  ' But Select already On therefore cancel
         TURNSELECTOFF
         If ImBPP(ImageNum) = 32 Then
            Select Case ImageNum
            Case 0: FILL3D DATACUL0(), DATACULBU()
            Case 1: FILL3D DATACUL1(), DATACULBU()
            Case 2: FILL3D DATACUL2(), DATACULBU()
            End Select
         End If
         Erase DATACULBU()
         
         'If CurrentGen(ImageNum) <> 0 Then
         '   cmdUndo_Click CInt(ImageNum)
         'End If
      End If
      
   Case [MoveSEL], [MoveCOPY]
      aMoveSEL = True
      '''''''''''''''''''''''' see picPANEL Mouse_Move
      Shape1.Move Box1(0).Left, Box1(0).Top, Box1(0).Width, Box1(0).Height
      Shape1.Visible = True
      '''''''''''''''''''''''
      DisableEffects
      
      sw = Box1(0).Width \ GridMult
      sh = Box1(0).Height \ GridMult
      xSrc = Box1(0).Left \ GridMult
      ySrc = Box1(0).Top \ GridMult
      
      Box1(0).Visible = False
      Box1(1).Visible = False
      
      picSEL(ImageNum).Picture = LoadPicture
      
      picSEL(ImageNum).Move Form1.Width / STX + xSrc, ySrc, sw, sh
      
      BitBlt picSEL(ImageNum).hdc, 0, 0, sw, sh, _
         picSmall(ImageNum).hdc, xSrc, ySrc, vbSrcCopy
      picSEL(ImageNum).Picture = picSEL(ImageNum).Image
      
      'picSEL(ImageNum).Visible = True
      
      If Index = [MoveSEL] Then   ' Else MoveCOPY done at picPANEL_MouseUp
         ' Leave behind RColor
         For iy = ySrc To ySrc + sh - 1
         For ix = xSrc To xSrc + sw - 1
            SetPixelV picSmall(ImageNum).hdc, ix, iy, RColor 'TColorBGR
         Next ix
         Next iy
         
         If ImBPP(ImageNum) = 32 Then
            For iy = ySrc To ySrc + sh - 1
            For ix = xSrc To xSrc + sw - 1
               If RColor <> TColorBGR Then ABYTE = 255 Else ABYTE = 0
               Cul = picSmall(ImageNum).Point(ix, iy)
               SetAlphas ix, iy, ABYTE
               CulTo32bppArrays Cul, ix, iy, ABYTE
            Next ix
            Next iy
         End If
      End If
      
      DrawGrid
      TransferpicSEL picSEL(ImageNum), picPANEL, xSrc * GridMult, ySrc * GridMult
      picPANEL.MouseIcon = LoadResPicture("4WAY", vbResCursor)
      picPANEL.MousePointer = vbCustom
   ' ==== SELECTION END =======
   
   End Select
End Sub

Private Sub optTools2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Long
   
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

   
   ' Switch other Tools off
   If Tools < [Tools2Blur] And Tools <> -1 Then
      optTools(CInt(Tools)).Value = False
   End If
   
   For k = 0 To optTools.Count - 1
      optTools(k).Value = False
   Next k
   
   ' Switch off OptTools2s [Tools2Blur] to [ReflectBottom]
   For k = 0 To optTools2.Count - 1
      optTools2(k).Value = False
   Next k
   
   Tools = [Tools2Blur] + Index ' 33+Index For continuos Tools numbering
   
   picPANEL.MouseIcon = LoadResPicture("PEN", vbResCursor)
   picPANEL.MousePointer = vbCustom
   ToolsCaption
   optTools2(Index).Value = True
   
   Select Case Tools
   Case [MirrorL], [MirrorR], [MirrorT], [MirrorB]
      If CurrentGen(ImageNum) = 0 Then Exit Sub
      
      shpCirc.Visible = False
      If Not Set_LOHI_XY Then Exit Sub
      
      Mirrors picSmall(ImageNum)
      
      ShowTheColorCount ImageNum
      If aSelectDrawn Then
         aEffects = True
         Tools = [SelectON]
      Else
         BackUp ImageNum
      End If

   Case [ReflectLeft] To [ReflectBottom]
   '   Tools
   '   [ReflectLeft]
   '   [ReflectRight]
   '   [ReflectTop]
   '   [ReflectBottom]
      If CurrentGen(ImageNum) = 0 Then Exit Sub
      
      optTools2(Index).Value = False
      shpCirc.Visible = False
      If Not Set_LOHI_XY Then Exit Sub
      
      Reflectors picSmall(ImageNum)
      
      ShowTheColorCount ImageNum
      If aSelectDrawn Then
         aEffects = True
         Tools = [SelectON]
      End If
      
   Case Else   ' [Tools2Blur] to [Tools2LessRed]
      If aSelect Then
         TURNSELECTOFF  ' Dose Dots
         Tools = [Tools2Blur] + Index
         ToolsCaption
         aEffects = False
      End If
   End Select
End Sub

Private Sub picSmallToPicSmallBU()
   picSmallBU.Picture = LoadPicture
   With picSmallBU
      .Width = picSmall(ImageNum).Width
      .Height = picSmall(ImageNum).Height
   End With

   BitBlt picSmallBU.hdc, 0, 0, picSmallBU.Width, picSmallBU.Height, _
       picSmall(ImageNum).hdc, 0, 0, vbSrcCopy
   picSmallBU.Picture = picSmallBU.Image
End Sub

Private Sub CancelSelection()
Dim k As Long
   ' if aEffects
   optTools([MoveSEL]).Enabled = False
   optTools([MoveCOPY]).Enabled = False
   Box1(0).Visible = False
   Box1(1).Visible = False
   Shape1.Visible = False
   picSEL(ImageNum).Visible = False
   aSelect = False
   aMoveSEL = False
   EnableEffects
   ' Disable Reflectors
   Label4(2).Enabled = False
   For k = 19 To 22
      optTools2(k).Enabled = False
   Next k
   aEffects = False
   aSelectDrawn = False
End Sub

Private Sub TURNSELECTOFF()
   CancelSelection
   picPANEL.MouseIcon = LoadResPicture("PEN", vbResCursor)
   picPANEL.MousePointer = vbCustom
   
   BitBlt picSmall(ImageNum).hdc, 0, 0, picSmallBU.Width, picSmallBU.Height, _
       picSmallBU.hdc, 0, 0, vbSrcCopy
   picSmall(ImageNum).Picture = picSmall(ImageNum).Image
   DrawGrid
   If Tools < [Tools2Blur] Then
      Tools = [Dots]
      optTools(0).Value = True
      ToolsCaption
   End If
End Sub

Private Sub DisableEffects()
Dim i As Integer
   
   For i = [BlurIm] To [ReplaceLbyR]
      optTools(i).Enabled = False
   Next i
   For i = 0 To optTools2.Count - 1
      optTools2(i).Enabled = False
   Next i
End Sub

Private Sub EnableEffects()
Dim i As Integer
   
   For i = [BlurIm] To [ReplaceLbyR]
      optTools(i).Enabled = True
   Next i
   For i = 0 To optTools2.Count - 1
      optTools2(i).Enabled = True
   Next i
End Sub

Private Function Set_LOHI_XY() As Boolean
' For Effects
   Set_LOHI_XY = False     ' Not nec
   
   LOX = 0: HIX = ImageWidth(ImageNum) - 1
   LOY = 0: HIY = ImageHeight(ImageNum) - 1

   If aSelect Then
      If aSelectDrawn = True Then
         LOX = Box1(0).Left \ GridMult
         HIX = LOX + Box1(0).Width \ GridMult - 1
         LOY = ImageHeight(ImageNum) - Box1(0).Height \ GridMult - Box1(0).Top \ GridMult
         HIY = ImageHeight(ImageNum) - Box1(0).Top \ GridMult - 1
         
         If HIX > ImageWidth(ImageNum) - 1 Then HIX = ImageWidth(ImageNum) - 1
         If HIY > ImageHeight(ImageNum) - 1 Then HIY = ImageHeight(ImageNum) - 1
         If LOX < 0 Then LOX = 0
         If LOY < 0 Then LOY = 0
      Else
         TURNSELECTOFF
      End If
   End If
   
   If LOY > HIY Or LOX > HIX Then
      MsgBox "Zero dimensions", vbInformation, "Selection rectangle"
      TURNSELECTOFF
   Else
      Set_LOHI_XY = True

' TEST
'LabTest = "   HIX,HIY=" & Str$(HIX) & ", " & Str$(HIY)

   End If
End Function

Private Sub DisplayEffects(Index As Integer)
' Only called from [BlurIm] to [BorderIm]
Dim BMIH As BITMAPINFOHEADER
   
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = ImageWidth(ImageNum)
      .biHeight = ImageHeight(ImageNum)
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With
   SetDIBits picSmall(ImageNum).hdc, picSmall(ImageNum).Image, _
      0, ImageHeight(ImageNum), picDATAREG(0, 0, 0), BMIH, 0
   
   picSmall(ImageNum).Picture = picSmall(ImageNum).Image
   
   DrawGrid
   optTools(Index).Value = False
   ShowTheColorCount ImageNum
   ToolsCaption
End Sub

'#### picPANEL Drawing ####

Private Sub picPANEL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Long
Dim r As RECT
   
   If LabMAC <> "Images" Then
      RestoreImage
      LabMAC = ""
      'aMouseDown = False
      Exit Sub
   End If

   aMouseDown = True
   
   If Button = vbLeftButton Then
      TheButton = vbLeftButton
      DrawCul = LColor
   ElseIf Button = vbRightButton Then
      TheButton = vbRightButton
      DrawCul = RColor
   Else
      TheButton = vbRightButton
      DrawCul = RColor
   End If
   
   If CurrentGen(ImageNum) = 0 Then
      Select Case Tools ' Exclude Effects apart from Border
      Case [Dots] To [Fill], [BoxCenShade] To [BoxHVCenShade], [Text], _
           [BorderIm], [Tools2Checker] To [Tools2VLines]
         BackUp ImageNum
      End Select
   End If
   
   If Tools <> [Dropper] And Tools <> [SelectON] And Tools <> -1 Then
      aSelectDrawn = False
   End If
   
   Select Case Tools
   Case [Dots]
      xStart = X
      yStart = Y
      DRAW 1
   Case [Lines]
      Line1(0).Visible = True
      Line1(1).Visible = True
      xStart = X
      yStart = Y
      xend = X
      yend = Y
      For k = 0 To 1
         Line1(k).x1 = xStart
         Line1(k).y1 = yStart
         Line1(k).x2 = xend
         Line1(k).y2 = yend
      Next k
   Case [Boxes], [BoxesSolid], [BoxesFilled], _
        [BoxCenShade], [BoxVertShade], [BoxHorzShade], _
        [BoxDiagTLBRShade], [BoxDiagBLTRShade], _
        [BoxVertCenShade], [BoxHorzCenShade], [BoxHVCenShade]
      xStart = X
      yStart = Y
      xend = X
      yend = Y
      
      StartOnLine
      
      For k = 0 To 1
         Box1(k).Visible = True
         Box1(k).Move xStart, yStart, 0, 0
      Next k

   Case [Ellipse], [EllipseSolid], [EllipseFilled], _
        [EllipseCenShade], [EllipseVertShade], [EllipseHorzShade]
      xStart = X
      yStart = Y
      xend = X
      yend = Y
      
      StartOnLine
      
      For k = 0 To 1
         Box1(k).Move xStart, yStart, 0, 0
         Box1(k).Visible = True
         Ellipse1(k).Move xStart, yStart, 0, 0
         Ellipse1(k).Visible = True
      Next k
   
   
   Case [Fill]
   Case [Dropper]
   
   Case [SelectON]
      If aEffects Then
         CancelSelection
         picPANEL.MouseIcon = LoadResPicture("PEN", vbResCursor)
         picPANEL.MousePointer = vbCustom
         Screen.MousePointer = vbDefault
         ' Avoid click through
         GetWindowRect Form1.picToolbar.hwnd, r
         SetCursorPos r.Left + 10, r.Top + 10
         
         Tools = -1 '  Now None to Avoid extra dot at 0,0
         'optTools(0).Value = True ' Dot button on
         ToolsCaption
         If CurrentGen(ImageNum) = 0 Then
            aSelectDrawn = True
            Exit Sub
         Else
            BackUp ImageNum
         End If
      Else
         Box1(0).Visible = True
         Box1(1).Visible = True
         xStart = X
         yStart = Y
         xend = X
         yend = Y
         
         StartOnLine
         
         For k = 0 To 1
            Box1(k).Move xStart, yStart, 0, 0
         Next k
         aSelectDrawn = True
      End If

   Case [MoveSEL]
   Case [MoveCOPY]
   
   Case [Tools2Blur] To [Tools2LessRed]
      shpCirc.Visible = True
      shpCirc.Left = X - shpCirc.Width + 4
      shpCirc.Top = Y - shpCirc.Height + 4
   
   End Select
End Sub

Private Sub picPANEL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
Dim r As Byte, g As Byte, B As Byte
Dim xSrc As Single, ySrc As Single
   
   LabXY = "X =" & Str$(X \ GridMult) & "    Y =" & Str$(Y \ GridMult)
         
   If LabMAC <> "Images" Then
      Cul = picPANEL.Point(X, Y)
      If Cul >= 0 Then
         LabDropper.BackColor = Cul
         LngToRGB Cul, r, g, B
         LabRGB(0) = r
         LabRGB(1) = g
         LabRGB(2) = B
      End If
      'cmdORGCul.SetFocus
      Exit Sub
   End If
   
   Select Case Tools
   Case [Tools2Blur] To [Tools2LessRed]
      shpCirc.Visible = True
      shpCirc.Left = X - shpCirc.Width + 4
      shpCirc.Top = Y - shpCirc.Width + 4
   End Select
   
   If aMouseDown Then
      Select Case Tools
      Case [Dots]
         xStart = X
         yStart = Y
         DRAW 1
      Case [Lines]
         xend = X
         yend = Y
         ' Draw shape line
         Line1(0).x2 = xend
         Line1(0).y2 = yend
         Line1(1).x2 = xend
         Line1(1).y2 = yend
      Case [Boxes], [BoxesSolid], [BoxesFilled], _
           [BoxCenShade], [BoxVertShade], [BoxHorzShade], _
           [BoxDiagTLBRShade], [BoxDiagBLTRShade], _
           [BoxVertCenShade], [BoxHorzCenShade], [BoxHVCenShade]
         xend = X
         yend = Y
         DrawBox1 Form1, picPANEL, X, Y
      Case [Ellipse], [EllipseSolid], [EllipseFilled], _
           [EllipseCenShade], [EllipseVertShade], [EllipseHorzShade]
         xend = X
         yend = Y
         DrawBox1 Form1, picPANEL, X, Y
         DrawEllipse1 X, Y
      Case [Fill]
      Case [Dropper]
      Case [SelectON]
         Box1(0).Visible = True
         Box1(1).Visible = True
         xend = X
         yend = Y
         
         DrawBox1 Form1, picPANEL, X, Y
      
      Case [MoveSEL], [MoveCOPY]
         If Button = vbLeftButton Then
            MoveBoxSEL X, Y
            ' Transfer to picPANEL
            xSrc = Box1(0).Left \ GridMult
            ySrc = Box1(0).Top \ GridMult
            '''''''''''''''''''''''
            Shape1.Visible = True
            Shape1.Move xSrc * GridMult, ySrc * GridMult, Box1(0).Width, Box1(0).Height
            'TransferSmallToLarge picSmall(ImageNum), picPANEL
            ' Possible but grid temp disappears. Need another Box.
            '''''''''''''''''''''''
            DrawGrid  ' Does TransferSmallToLarge picSmall(ImageNum), picPANEL

            TransferpicSEL picSEL(ImageNum), picPANEL, xSrc * GridMult, ySrc * GridMult
         End If

      Case [Tools2Blur] To [Tools2LessRed]
         shpCirc.Visible = True
         shpCirc.Left = X - shpCirc.Width + 4
         shpCirc.Top = Y - shpCirc.Width + 4
      End Select
      
      Select Case Tools
      Case [Tools2Blur] 'To [Tools2VLines]
         BlurTool picSmall(ImageNum), X, Y
      Case [Tools2Grey]
         GreyTool picSmall(ImageNum), X, Y
      Case [Tools2Invert]
         InvertTool picSmall(ImageNum), X, Y
      Case [Tools2Bright], [Tools2Dark]
         BrightDarkTool picSmall(ImageNum), X, Y
         
      Case [Tools2Checker]
         CheckerBoard picSmall(ImageNum), X, Y, DrawCul
      Case [Tools2HLines]
         HorzLines picSmall(ImageNum), X, Y, DrawCul
      Case [Tools2VLines]
         VertLines picSmall(ImageNum), X, Y, DrawCul
      
      Case [Tools2Random]
         RandomTool picSmall(ImageNum), X, Y
         
      Case [Tools2MoreBlue] To [Tools2LessRed]
         BrightDarkTool picSmall(ImageNum), X, Y
      End Select
   End If
   
   If Tools = [Dropper] Then
      If aDropper Then
         Cul = picPANEL.Point(X, Y)
         If Cul >= 0 Then
            LabDropper.BackColor = Cul
            LngToRGB Cul, r, g, B
            LabRGB(0) = r
            LabRGB(1) = g
            LabRGB(2) = B
         End If
      End If
   End If
End Sub

Private Sub picPANEL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Long
Dim Cul As Long
Dim CulFill As Long
Dim xSrc As Single, ySrc As Single
Dim sw As Single, sh As Single
Dim ix As Long, iy As Long
Dim ABYTE As Byte

   
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
   
   aMouseDown = False
   
   Select Case Tools
   Case [Dots]
      DrawGrid
   Case [Lines]
      Line1(0).Visible = False
      Line1(1).Visible = False
      xend = X
      yend = Y
      DRAW
   Case [Boxes], [BoxesSolid], [BoxesFilled], _
        [Ellipse], [EllipseSolid], [EllipseFilled], _
        [BoxCenShade], [BoxVertShade], [BoxHorzShade], _
        [EllipseCenShade], [EllipseVertShade], [EllipseHorzShade], _
        [BoxDiagTLBRShade], [BoxDiagBLTRShade], _
        [BoxVertCenShade], [BoxHorzCenShade], [BoxHVCenShade]
      
      Box1(0).Visible = False
      Box1(1).Visible = False
      Ellipse1(0).Visible = False
      Ellipse1(1).Visible = False
      xend = X
      yend = Y
      DRAW
   Case [Fill]
      Cul = picSmall(ImageNum).Point(X \ GridMult, Y \ GridMult)
      If Cul >= 0 Then
         If TheButton = vbLeftButton Then
            CulFill = LColor
         ElseIf TheButton = vbRightButton Then
            CulFill = RColor
         Else
            CulFill = RColor
         End If
         
         picSmall(ImageNum).FillColor = CulFill
         picSmall(ImageNum).FillStyle = vbFSSolid
         ExtFloodFill picSmall(ImageNum).hdc, X \ GridMult, Y \ GridMult, Cul, FLOODFILLSURFACE
         picSmall(ImageNum).Picture = picSmall(ImageNum).Image
         picSmall(ImageNum).FillStyle = vbFSTransparent  'Default (Transparent)
         
         If ImBPP(ImageNum) = 32 Then
            For iy = 0 To ImageHeight(ImageNum) - 1
            For ix = 0 To ImageWidth(ImageNum) - 1
               Cul = picSmall(ImageNum).Point(ix, iy)
               If Cul = TColorBGR Then
                  ABYTE = 0
                  SetAlphas ix, iy, ABYTE
               ElseIf Cul = CulFill Then
                  ABYTE = 255
                  SetAlphas ix, iy, ABYTE
                  CulTo32bppArrays Cul, ix, iy, ABYTE
               End If
            Next ix
            Next iy
         End If
         
         DrawGrid
      End If
   Case [SelectON]
       optTools([MoveSEL]).Enabled = True
       optTools([MoveCOPY]).Enabled = True
       Label4(2).Enabled = True
       For k = 19 To 22 ' Reflectors
         optTools2(k).Enabled = True
       Next k
       k = k
   Case [MoveSEL], [MoveCOPY]
      If Button = vbRightButton Then  ' Fix moved image
         ' Transfer picSEL to picSmall()
         sw = Box1(0).Width \ GridMult
         sh = Box1(0).Height \ GridMult
         xSrc = Box1(0).Left \ GridMult
         ySrc = Box1(0).Top \ GridMult
         
         BitBlt picSmall(ImageNum).hdc, xSrc, ySrc, sw, sh, _
            picSEL(ImageNum).hdc, 0, 0, vbSrcCopy
         picSmall(ImageNum).Picture = picSmall(ImageNum).Image
         
         
         If ImBPP(ImageNum) = 32 Then
            For iy = ySrc To ySrc + sh - 1
            For ix = xSrc To xSrc + sw - 1
               Cul = picSmall(ImageNum).Point(ix, iy)
               If Cul = TColorBGR Then ABYTE = 0 Else ABYTE = 255
               SetAlphas ix, iy, ABYTE
               CulTo32bppArrays Cul, ix, iy, ABYTE
            Next ix
            Next iy
         End If
         
         Shape1.Visible = False
         
         picSEL(ImageNum).Visible = False
         DrawGrid
         
         ' Cancel Selection settings
         CancelSelection
         
         Tools = [Dots]
         optTools(0).Value = True
         ToolsCaption
         
         picPANEL.MouseIcon = LoadResPicture("PEN", vbResCursor)
         picPANEL.MousePointer = vbCustom
         Screen.MousePointer = vbDefault
         
         If CurrentGen(ImageNum) = 0 Then Exit Sub
   
      End If
   End Select
   
   If Tools = [Dropper] Then
      If aDropper Then
         LabErase.Visible = False
         LineErase.Visible = False
         Cul = picPANEL.Point(X, Y)
         If Cul >= 0 Then
            If Button = vbLeftButton Then
               LColor = Cul
               LabColor(0).BackColor = Cul
               If LColor = TColorBGR Then
                  LabErase.Caption = "LColor Erases"
                  LabErase.Visible = True
                  LineErase.Visible = True
               End If
            ElseIf Button = vbRightButton Then
               RColor = Cul
               LabColor(1).BackColor = Cul
               If RColor = TColorBGR Then
                  LabErase.Caption = "RColor Erases"
                  LabErase.Visible = True
                  LineErase.Visible = True
               End If
            Else
               RColor = Cul
               LabColor(1).BackColor = Cul
            End If
         End If
         ' Check if Dropper has picked up Erase color
         If LColor = TColorBGR And RColor = TColorBGR Then
            LabErase.Caption = "LColor && RColor Erases"
            LabErase.Visible = True
            LineErase.Visible = True
         End If
      Else
         LabColor(0).BackColor = LColor
      End If
   End If
   
   Select Case Tools
   
   Case [Dots]
      'DrawGrid  ' Here for dots as 64x64 grid is a bit slow
                ' when TPixels marked
   Case [Dropper]
      aDropper = False
      Tools = [Dots]
      optTools(0).Value = True
      ToolsCaption
      picPANEL.MouseIcon = LoadResPicture("PEN", vbResCursor)
      picPANEL.MousePointer = vbCustom
      Exit Sub
   End Select
   
   ' Color count
   ShowTheColorCount ImageNum

   If Tools <> -1 Then
      Select Case Tools
      Case Is < [Text], [BoxCenShade] To [BoxHVCenShade] ', Is >= [Tools2Blur]
         If aSelectDrawn Then
            aSelectDrawn = False
         Else
            BackUp ImageNum
         End If
      Case Is >= [Tools2Blur]
         If CurrentGen(ImageNum) = 0 Then
            Exit Sub
         Else
            BackUp ImageNum
         End If
      End Select
   End If
   LabSpec = "(Bpp =" & Str$(ImBPP(ImageNum)) & ")"
End Sub

Private Sub DRAW(Optional DoGrid As Long = 0)

' All these Tools come to DRAW  from picPANEL_MouseDown for Dots
' and picPANEL_MouseUp for the rest :
' [Dots], [Lines] [Boxes], [BoxesSolid], [BoxesFilled], _
' [Ellipse] , [EllipseSolid], [EllipseFilled], _
' [BoxCenShade], [BoxVertShade], [BoxHorzShade], _
' [EllipseCenShade], [EllipseVertShade], [EllipseHorzShade], _
' [BoxDiagTLBRShade], [BoxDiagBLTRShade], _
' [BoxVertCenShade], [BoxHorzCenShade],[BoxHVCenShade]

Dim k As Long
Dim ix As Long, iy As Long
Dim xxs As Long, xxe As Long
Dim yys As Long, yye As Long

'Dim xxs As Single, xxe As Single
'Dim yys As Single, yye As Single

Dim bxxs As Long, bxxe As Long
Dim byys As Long, byye As Long

' For Circle
Dim xc As Single, yc As Single
Dim zrad As Single, zradx As Single, zrady As Single, zasp As Single
Dim Cul As Long

' For shading
Dim zIStep As Single
Dim ixa As Long, iya As Long
Dim ixb As Long, iyb As Long
Dim xk As Single, DY As Single
Dim yk As Single, dx As Single
' For diagonal shading
Dim zdx As Single, zdy As Single
Dim zx As Single, zy As Single
   
Dim ABYTE As Byte
   
' xStart = X
' Ystart = Y
   If DrawCul < 0 Then Exit Sub
   
   xxs = xStart \ GridMult
   If xxs < 0 Then xxs = 0
   If xxs > ImageWidth(ImageNum) - 1 Then xxs = ImageWidth(ImageNum) - 1
   xxe = xend \ GridMult
   If xxe < 0 Then xxe = 0
   If xxe > ImageWidth(ImageNum) - 1 Then xxe = ImageWidth(ImageNum) - 1
   
   yys = yStart \ GridMult
   If yys < 0 Then yys = 0
   If yys > ImageHeight(ImageNum) - 1 Then yys = ImageHeight(ImageNum) - 1
   yye = yend \ GridMult
   If yye < 0 Then yye = 0
   If yye > ImageHeight(ImageNum) - 1 Then yye = ImageHeight(ImageNum) - 1
   
   bxxs = Box1(0).Left \ GridMult
   bxxe = bxxs + Box1(0).Width \ GridMult - 1
   byys = Box1(0).Top \ GridMult
   byye = byys + Box1(0).Height \ GridMult - 1
   
   '============================================================

   Select Case Tools
   Case [Dots]
      xleft = (xxs) * GridMult + 1
      xright = xleft + GridMult - 2
      ytop = (yys) * GridMult + 1
      ybelow = ytop + GridMult - 2
      
      picPANEL.Line (xleft, ytop)-(xright, ybelow), DrawCul, BF
      SetPixelV picSmall(ImageNum).hdc, xxs, yys, DrawCul
      
      picSmall(ImageNum).Picture = picSmall(ImageNum).Image
      
      If ImBPP(ImageNum) = 32 Then
         ABYTE = 255
         If GetPixel(picSmall(ImageNum).hdc, xxs, yys) = TColorBGR Then
            ABYTE = 0
         End If
         SetAlphas CLng(xxs), CLng(yys), ABYTE
         CulTo32bppArrays DrawCul, CLng(xxs), CLng(yys), ABYTE
      End If
   Case [Lines]
      picSmall(ImageNum).Line (xxs, yys)-(xxe, yye), DrawCul
      picSmall(ImageNum).PSet (xxe, yye), DrawCul  ' because of Bug in Line function
   Case [Boxes]
      picSmall(ImageNum).Line (bxxs, byys)-(bxxe, byye), DrawCul, B
   Case [BoxesSolid]
      picSmall(ImageNum).Line (bxxs, byys)-(bxxe, byye), DrawCul, BF
   Case [BoxesFilled]
      picSmall(ImageNum).Line (bxxs, byys)-(bxxe, byye), DrawCul, B
      xc = (bxxs + bxxe) \ 2
      yc = (byys + byye) \ 2
      Cul = picSmall(ImageNum).Point(xc, yc)
      If TheButton = vbLeftButton Then
         picSmall(ImageNum).FillColor = RColor
      Else
         picSmall(ImageNum).FillColor = LColor
      End If
      picSmall(ImageNum).FillStyle = vbFSSolid
      'ExtFloodFill picSmall(ImageNum).hDC, xc, yc, Cul, FLOODFILLSURFACE
      ExtFloodFill picSmall(ImageNum).hdc, xc, yc, DrawCul, FLOODFILLBORDER
      picSmall(ImageNum).Picture = picSmall(ImageNum).Image
      picSmall(ImageNum).FillStyle = vbFSTransparent  'Default (Transparent)
   '============================================================
   
   Case [BoxCenShade]
   
      AdjustBoxPoints bxxs, byys, bxxe, byye   ' s TL, e BR
      picSmall(ImageNum).Line (bxxs, byys)-(bxxe, byye), DrawCul, B
      NSteps = bxxe - bxxs
      If byye - byys < NSteps Then NSteps = byye - byys
      zIStep = Abs(bxxe - bxxs) / (2 * NSteps)
      If zIStep > Abs(byye - byys) / (2 * NSteps) Then
         zIStep = Abs(byye - byys) / (2 * NSteps)
      End If
      GetColorSteps
      For k = 0 To NSteps
         ixa = bxxs + k * zIStep
         iya = byys + k * zIStep
         ixb = bxxe - k * zIStep
         iyb = byye - k * zIStep
         picSmall(ImageNum).Line (ixa, iya)-(ixb, iyb), RGB(zRL, zGL, zBL), B
         GetNextColors
      Next k
      
   Case [BoxVertShade]
      AdjustBoxPoints bxxs, byys, bxxe, byye   ' s TL, e BR
      picSmall(ImageNum).Line (bxxs, byys)-(bxxs, byye + 1), DrawCul
      NSteps = bxxe - bxxs
      GetColorSteps
      ixa = bxxs
      For k = 0 To NSteps
         picSmall(ImageNum).Line (ixa, byys)-(ixa, byye + 1), RGB(zRL, zGL, zBL)
         ixa = ixa + 1
         GetNextColors
      Next k

   Case [BoxHorzShade]
      AdjustBoxPoints bxxs, byys, bxxe, byye   ' s TL, e BR
      picSmall(ImageNum).Line (bxxs, byys)-(bxxe + 1, byys), DrawCul
      NSteps = byye - byys
      GetColorSteps
      iya = byys
      For k = 0 To NSteps
         picSmall(ImageNum).Line (bxxs, iya)-(bxxe + 1, iya), RGB(zRL, zGL, zBL)
         iya = iya + 1
         GetNextColors
      Next k
      
   '------------------------------------------------------------------------
   Case [Ellipse], [EllipseSolid], [EllipseFilled], _
        [EllipseCenShade], [EllipseVertShade], [EllipseHorzShade]
      ' Calc for all ellipse params
      zradx = Abs(bxxe - bxxs) / 2
      zrady = Abs(byye - byys) / 2
      If zradx = 0 Then
         zrad = zrady
         zasp = 10
      ElseIf zradx >= zrady Then
         zrad = zradx
         zasp = zrady / zradx
      Else  'zradx < zrady
         zrad = zrady
         zasp = zrady / zradx
      End If
      xc = (bxxe + bxxs) / 2
      yc = (byye + byys) / 2
      
   '============================================================
      Select Case Tools
         Case [Ellipse]
            If zrad <= 1 Then
               picSmall(ImageNum).PSet (xxs, yys), DrawCul
            Else
               picSmall(ImageNum).Circle (xc, yc), zrad, DrawCul, , , zasp
            End If
         Case [EllipseSolid]
            If zrad <= 1 Then
               picSmall(ImageNum).PSet (xxs, yys), DrawCul
            Else
               picSmall(ImageNum).FillStyle = 0      ' Solid
               picSmall(ImageNum).FillColor = DrawCul
               picSmall(ImageNum).Circle (xc, yc), zrad, DrawCul, , , zasp
               picSmall(ImageNum).FillStyle = 1      ' Default (Transparent)
            End If
            
         Case [EllipseFilled] 'with right or left color
            If zrad <= 1 Then
               picSmall(ImageNum).PSet (xxs, yys), DrawCul
            Else
               picSmall(ImageNum).Circle (xc, yc), zrad, DrawCul, , , zasp
               Cul = picSmall(ImageNum).Point(xc, yc)
               If TheButton = vbLeftButton Then
                  picSmall(ImageNum).FillColor = RColor
               Else
                  picSmall(ImageNum).FillColor = LColor
               End If
               picSmall(ImageNum).FillStyle = vbFSSolid
               'ExtFloodFill picSmall(ImageNum).hDC, xc, yc, Cul, FLOODFILLSURFACE
               ExtFloodFill picSmall(ImageNum).hdc, xc, yc, DrawCul, FLOODFILLBORDER
      
               picSmall(ImageNum).Picture = picSmall(ImageNum).Image
               picSmall(ImageNum).FillStyle = vbFSTransparent  'Default (Transparent)
            End If
   '============================================================
      
         Case [EllipseCenShade]
            If zrad <= 1 Then
               picSmall(ImageNum).PSet (xxs, yys), DrawCul
            Else
               If zradx = 0 Then zradx = 1
               If zrady = 0 Then zrady = 1
               If zradx > zrady Then
                  NSteps = 5 * (2 * zradx)
               Else
                  NSteps = 5 * (2 * zrady)
               End If
               GetColorSteps
               If zrad = 0 Then zrad = 1
               picSmall(ImageNum).DrawWidth = 2
               For xk = zrad To 0 Step -0.1
                  Cul = RGB(zRL, zGL, zBL)
                  If Cul = TColorBGR Then
                     Cul = Cul + 1
                  End If
                  picSmall(ImageNum).Circle (xc, yc), xk, Cul, , , zasp
                  GetNextColors
               Next xk
               picSmall(ImageNum).DrawWidth = 1
            End If
         
         Case [EllipseVertShade]
            If zrad <= 1 Then
               picSmall(ImageNum).PSet (xxs, yys), DrawCul
            Else
               If zradx = 0 Then zradx = 1
               If zrady = 0 Then zrady = 1
               NSteps = 10 * (2 * zradx)
               GetColorSteps
               For xk = xc - zradx To xc + zradx Step 0.1
                  DY = zrady * Sqr(1 - (xc - xk) ^ 2 / zradx ^ 2)
                  picSmall(ImageNum).Line (xk, yc - DY)-(xk, yc + DY + 1), RGB(zRL, zGL, zBL)
                  GetNextColors
               Next xk
            End If
         
         Case [EllipseHorzShade]
            If zrad <= 1 Then
               picSmall(ImageNum).PSet (xxs, yys), DrawCul
            Else
               If zradx = 0 Then zradx = 1
               If zrady = 0 Then zrady = 1
               NSteps = 10 * (2 * zrady)
               GetColorSteps
               For yk = yc - zrady To yc + zrady Step 0.1
                  dx = zradx * Sqr(1 - (yc - yk) ^ 2 / zrady ^ 2)
                  picSmall(ImageNum).Line (xc - dx, yk)-(xc + dx + 1, yk), RGB(zRL, zGL, zBL)
                  GetNextColors
               Next yk
            End If
      End Select
   '============================================================
      
   Case [BoxDiagTLBRShade], [BoxDiagBLTRShade]
      AdjustBoxPoints bxxs, byys, bxxe, byye   ' s TL, e BR
      dx = (bxxe - bxxs)
      DY = (byye - byys)
      NSteps = dx
      GetColorStepsReversed
      If dx = DY Then
         zdx = 1
         zdy = 1
      ElseIf dx > DY Then
         zdx = 1
         zdy = DY / (dx - 1)
      Else  ' dy > dx
         NSteps = DY
         GetColorStepsReversed
         zdy = 1
         zdx = dx / (DY - 1)
      End If
      If Tools = [BoxDiagTLBRShade] Then  ' \\
         For k = 0 To NSteps
            zy = zdy * k
            zx = zdx * k
            If bxxs + zx <= bxxe Then
               picSmall(ImageNum).Line (bxxs + zx, byys)-(bxxe, byye - zy), RGB(zRL, zGL, zBL)
               If byye - zy >= byys Then
                  picSmall(ImageNum).PSet (bxxe, byye - zy), RGB(zRL, zGL, zBL)
               End If
            End If
            If bxxe - zx >= bxxs Then
               picSmall(ImageNum).Line (bxxe - zx, byye)-(bxxs, byys + zy), RGB(zRL, zGL, zBL)
               If byys + zy <= byye Then
                  picSmall(ImageNum).PSet (bxxs, byys + zy), RGB(zRL, zGL, zBL)
               End If
            End If
            GetNextColors
         Next k
      Else  ' [BoxDiagBLTRShade]  //
         For k = 0 To NSteps
            zy = zdy * k
            zx = zdx * k
            If bxxe - zx >= bxxs Then
               picSmall(ImageNum).Line (bxxe - zx, byys)-(bxxs, byye - zy), RGB(zRL, zGL, zBL)
            End If
            If byye - zy >= byys Then
               picSmall(ImageNum).PSet (bxxs, byye - zy), RGB(zRL, zGL, zBL)
            End If
            If byys + zy <= byye Then
               picSmall(ImageNum).Line (bxxe, byys + zy)-(bxxs + zx, byye), RGB(zRL, zGL, zBL)
            End If
            If bxxs + zx <= bxxe Then
               picSmall(ImageNum).PSet (bxxs + zx, byye), RGB(zRL, zGL, zBL)
            End If
            GetNextColors
         Next k
      End If
   
   Case [BoxVertCenShade] To [BoxHVCenShade]
      AdjustBoxPoints bxxs, byys, bxxe, byye   ' s TL, e BR
      If Tools = [BoxVertCenShade] Then  ' ||
         picSmall(ImageNum).Line (bxxs, byys)-(bxxs, byye + 1), DrawCul
         NSteps = (bxxe - bxxs) \ 2
         If NSteps = 0 Then NSteps = 1
         GetColorSteps
         ixa = bxxs
         For k = 0 To NSteps
            picSmall(ImageNum).Line (ixa, byys)-(ixa, byye + 1), RGB(zRL, zGL, zBL)
            picSmall(ImageNum).Line (bxxe - k, byys)-(bxxe - k, byye + 1), RGB(zRL, zGL, zBL)
            ixa = ixa + 1
            GetNextColors
         Next k
      ElseIf Tools = [BoxHorzCenShade] Then
         picSmall(ImageNum).Line (bxxs, byys)-(bxxe + 1, byys), DrawCul
         NSteps = (byye - byys) \ 2
         If NSteps = 0 Then NSteps = 1
         GetColorSteps
         iya = byys
         For k = 0 To NSteps
            picSmall(ImageNum).Line (bxxs, iya)-(bxxe + 1, iya), RGB(zRL, zGL, zBL)
            picSmall(ImageNum).Line (bxxs, byye - k)-(bxxe + 1, byye - k), RGB(zRL, zGL, zBL)
            iya = iya + 1
            GetNextColors
         Next k
      Else  ' [BoxHVCenShade] different patterns when image not square
         Dim NStepsX As Long, NStepsY As Long, NStepsMax As Long
         
         picSmall(ImageNum).Line (bxxs, byys)-(bxxe + 1, byys), DrawCul
         NStepsY = (byye - byys) \ 2
         NStepsX = (bxxe - bxxs) \ 2
         NStepsMax = NStepsX
         If NStepsY > NStepsX Then NStepsMax = NStepsY
         
         If NStepsMax = 0 Then NStepsMax = 1
         NSteps = NStepsMax
         GetColorSteps
         iya = byys
         ixa = bxxs
         For k = 0 To NStepsMax
            If k <= NStepsY Then
               picSmall(ImageNum).Line (bxxs, iya)-(bxxe + 1, iya), RGB(zRL, zGL, zBL)
               picSmall(ImageNum).Line (bxxs, byye - k)-(bxxe + 1, byye - k), RGB(zRL, zGL, zBL)
               iya = iya + 1
            End If
            If k <= NStepsX Then
               picSmall(ImageNum).Line (ixa, byys)-(ixa, byye + 1), RGB(zRL, zGL, zBL)
               picSmall(ImageNum).Line (bxxe - k, byys)-(bxxe - k, byye + 1), RGB(zRL, zGL, zBL)
               ixa = ixa + 1
            End If
            GetNextColors
         Next k
      End If
   
   Case [Dropper]
   End Select
      
   '=========================================================================
   ' 32 BPP
   If Tools <> [Dots] Then
      If ImBPP(ImageNum) = 32 Then
         ' Adjust Alphas for all Tools apart from Dots & Fill
         Select Case Tools
         Case [Lines], _
              [Boxes], [BoxesSolid], _
              [Ellipse], [EllipseSolid]
            ' Scan the whole of picSmall(ImageNum) &
            ' where color = TColorBGR set Alpha#() to 0
            ' where color = DrawCul set Alpha#() to 255
            For iy = 0 To ImageHeight(ImageNum) - 1
            For ix = 0 To ImageWidth(ImageNum) - 1
               Cul = picSmall(ImageNum).Point(ix, iy)
               If Cul = TColorBGR Then
                  ABYTE = 0
                  SetAlphas ix, iy, ABYTE
               ElseIf Cul = DrawCul Then
                  ABYTE = 255
                  CulTo32bppArrays Cul, ix, iy, ABYTE
               End If
            Next ix
            Next iy
         
         Case [BoxesFilled], [EllipseFilled]
            ' Scan the whole of picSmall(ImageNum) &
            ' where color = TColorBGR set Alpha to 0
            ' where color = DrawCul set Alpha to 255
            ' where color = FillCul set Alpha to 255
            For iy = 0 To ImageHeight(ImageNum) - 1
            For ix = 0 To ImageWidth(ImageNum) - 1
               Cul = picSmall(ImageNum).Point(ix, iy)
               If Cul = TColorBGR Then
                  ABYTE = 0
                  SetAlphas ix, iy, ABYTE
               ElseIf Cul = DrawCul Or Cul = picSmall(ImageNum).FillColor Then
                  ABYTE = 255
                  CulTo32bppArrays Cul, ix, iy, ABYTE
               End If
            Next ix
            Next iy
         
         Case [BoxCenShade] To [BoxHorzShade], _
              [BoxDiagTLBRShade] To [BoxHVCenShade]  ' All Shaded Boxes
            ' Range of colors, L to R Color
            For iy = byys To byye
            For ix = bxxs To bxxe
               Cul = picSmall(ImageNum).Point(ix, iy)
               ABYTE = 255
               CulTo32bppArrays Cul, ix, iy, ABYTE
            Next ix
            Next iy
         
         Case [EllipseCenShade] To [EllipseHorzShade]
         ' Not a box
            For iy = byys To byye
            For ix = bxxs To bxxe
               Cul = picSmall(ImageNum).Point(ix, iy)
               If Cul <> TColorBGR Then
                  ABYTE = 255
                  CulTo32bppArrays Cul, ix, iy, ABYTE
               End If
            Next ix
            Next iy

         End Select
   '============================================================
      End If
   End If
   If DoGrid = 0 Then DrawGrid
End Sub


'####  For shading ####

Private Sub AdjustBoxPoints(x1 As Long, y1 As Long, x2 As Long, y2 As Long)
' Make x1,y1 TopLeft - x2,y2 bottom right
   Dim xys As Integer
   If x1 > x2 Then
      xys = x1
      x1 = x2
      x2 = xys
   End If
   If y1 > y2 Then
      xys = y1
      y1 = y2
      y2 = xys
   End If
   ' To avoid division by 0
   If x2 = x1 Then x2 = x1 + 1
   If y2 = y1 Then y2 = y1 + 1
End Sub

Public Sub GetColorSteps()
' All Public
' For Shading
   zRL = (LColor And &HFF&)
   zGL = (LColor And &HFF00&) / &H100&
   zBL = (LColor And &HFF0000) / &H10000
   zRR = (RColor And &HFF&)
   zGR = (RColor And &HFF00&) / &H100&
   zBR = (RColor And &HFF0000) / &H10000
   zRIncr = (zRR - zRL) / NSteps
   zGIncr = (zGR - zGL) / NSteps
   zBIncr = (zBR - zBL) / NSteps
End Sub

Public Sub GetColorStepsReversed()
' All Public
' For Shading
   zRL = (RColor And &HFF&)
   zGL = (RColor And &HFF00&) / &H100&
   zBL = (RColor And &HFF0000) / &H10000
   zRR = (LColor And &HFF&)
   zGR = (LColor And &HFF00&) / &H100&
   zBR = (LColor And &HFF0000) / &H10000
   zRIncr = (zRR - zRL) / NSteps
   zGIncr = (zGR - zGL) / NSteps
   zBIncr = (zBR - zBL) / NSteps
End Sub


Public Sub GetNextColors()
' All Public
' For shading
   zRL = zRL + zRIncr
   If zRL < 0 Then zRL = 0
   If zRL > 255 Then zRL = 255
   zGL = zGL + zGIncr
   If zGL < 0 Then zGL = 0
   If zGL > 255 Then zGL = 255
   zBL = zBL + zBIncr
   If zBL < 0 Then zBL = 0
   If zBL > 255 Then zBL = 255
End Sub

'##############################

Private Sub DrawEllipse1(X As Single, Y As Single)
' Draw shape ellipse
' From picPANEL_MouseMove
' xStart & yStart from picPANEL_MouseDown
Dim sw As Long
Dim sh As Long
Dim xs As Long, ys As Long
Dim xe As Long, ye As Long
   
   xs = xStart
   ys = yStart
   xe = X
   ye = Y
      
   If xe > picPANEL.Width - 1 Then
      xe = picPANEL.Width
   End If
      
   If xe < xs Then
      If xe < 0 Then xe = 0
      X = xe
      xe = xs
      xs = X
      xs = (xs \ GridMult) * GridMult
   Else
      xe = (xe \ GridMult) * GridMult
   End If

   If ye > picPANEL.Height - 1 Then
      ye = picPANEL.Height
   End If
   If ye < ys Then
      If ye < 0 Then ye = 0
      Y = ye
      ye = ys
      ys = Y
      ys = (ys \ GridMult) * GridMult
   Else
      ye = (ye \ GridMult) * GridMult
   End If
   sw = Abs(xe - xs)
   sh = Abs(ye - ys)
      
   Ellipse1(0).Move xs, ys, sw, sh
   Ellipse1(1).Move xs, ys, sw, sh
End Sub

Private Sub MoveBoxSEL(X As Single, Y As Single)
' Move shape box from picPANEL Mouse_Move
   Box1(0).Left = X - Box1(0).Width \ 2
   Box1(0).Top = Y - Box1(0).Height \ 2
   
   ' Keep select on picPANEL  - bit messy
   If Box1(0).Left < 0 Then
      Box1(0).Left = 0
   End If
   If Box1(0).Top < 0 Then
      Box1(0).Top = 0
   End If
   If Box1(0).Left + Box1(0).Width > picPANEL.Width Then
      Box1(0).Left = picPANEL.Width - Box1(0).Width '- 1
   End If
   If Box1(0).Top + Box1(0).Height > picPANEL.Height Then
      Box1(0).Top = picPANEL.Height - Box1(0).Height '- 1
   End If

   Box1(1).Left = Box1(0).Left
   Box1(1).Top = Box1(0).Top
End Sub

'#### Opening Images ####

Private Sub mnuFile_Click()
   If aSelect Then TURNSELECTOFF
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
End Sub

Private Sub mnuOpen_Click()
   If aSelect Then TURNSELECTOFF
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
End Sub

Private Sub mnuOpenintoImage_Click(Index As Integer)
Dim Title$, Filt$, InDir$
Dim r As RECT
   
   Screen.MousePointer = vbDefault
   
   If aSelect Then TURNSELECTOFF

   ReDim ICOCURBPP(0 To 2)
   
   ' To prevent click-thru
   Select Case Tools
   Case [Tools2Blur] To [ReflectBottom]   '33 - 55
      optTools2(Tools - 33).Value = False
      shpCirc.Visible = False
      Tools = -1
      LabTool = "None"
   Case [BoxCenShade] To [BoxHVCenShade]
      optTools(Tools).Value = False
      shpCirc.Visible = False
      Tools = -1
      LabTool = "None"
   Case -1
   Case Else
      optTools(Tools).Value = False
      Tools = -1
      LabTool = "None"
      picPANEL.MouseIcon = LoadResPicture("PEN", vbResCursor)
      picPANEL.MousePointer = vbCustom
   End Select
   
   TempImageNum = Index ' 0,1,2
   If Jumper = 0 Then  ' ie not Drag/Drop
      
      If Index = 3 Then Exit Sub ' Info only (Will fill from Image 1 if multi-con)
      Title$ = "Load image file <= 64 x 64  to Image =" & Str$(Index + 1)
      Filt$ = "Pics bmp,gif,jpg,ico,cur|*.bmp;*.gif;*.jpg;*.ico;*.cur"
      FileSpec$ = ""
      InDir$ = CPath$ 'AppPathSpec$
      Set CommonDialog1 = New cOSDialog
   
      CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd
      Set CommonDialog1 = Nothing
      
   End If
   
   ' Drag/Drop comes here
   Jumper = 0
   
   ' Avoid click through
   GetWindowRect Form1.picSmallFrame.hwnd, r
   SetCursorPos r.Left + 130, r.Top + 90

   ProcessInFile 'FileSpec$

End Sub

Private Sub ProcessInFile() 'FileSpec$
' From mnuOpenIntoImage
'      mnuReloadImage
'      mnuRecentFiles
Dim Ext$
Dim fnum As Integer
Dim Message$
' For icon & cursor
Dim BStream() As Byte
Dim IcoStream() As Byte
Dim BMPWidth As Integer
Dim BMPHeight As Integer
Dim k As Long
Dim i As Long
Dim NB As Byte, NG As Byte, NR As Byte, NA As Byte
Dim NB0 As Byte, NG0 As Byte, NR0 As Byte
Dim TB As Byte, TG As Byte, TR As Byte
Dim Palpha As Single
Dim ResValue As Long
Dim ix As Long, iy As Long
Dim a$, resp$
' For large picture's aspect ratio
Dim BUW As Long
Dim BUH As Long
Dim zaspectBU As Single
Dim neww As Long
Dim newh As Long

Dim TheNAM$, ThePATH$

   'mnuOpenInmtoImage does
   'TempImageNum = Index ' 0,1,2

   On Error GoTo OPENERROR
  
   If Len(FileSpec$) = 0 Then
      Exit Sub
   Else
      
      ReDim ICOCURBPP(0 To 2)
         
      Ext$ = UCase$(Right$(FileSpec$, 3))
      
      CheckFile FileSpec$, Message$
      
      ' If ICO returns
      '    NumIcons bpp in IcoStream(36)
      '  if BMP bpp = BMPbpp  (IntegerMarker(2))
      
      If Len(Message$) <> 0 Then
         MsgBox Message$, vbCritical, "Opening " & Ext$ & " file"
         Exit Sub
      End If
   
      'If CurrentGen(ImageNum) = 0 Then BackUp ImageNum
      
      CPath$ = FileSpec$
      If Ext$ <> "ICO" And Ext$ <> "CUR" Then 'BMP, GIF, JPG
         
         BMPbpp = 0
         If Ext$ = "BMP" Then   ' Get bpp for BMP
            fnum = FreeFile
            Open FileSpec$ For Binary As #fnum
            i = LOF(fnum)
            ReDim BStream(0 To i - 1)
            Get #fnum, , BStream
            Close #fnum
            'k = UBound(BStream(), 1)
            BMPbpp = BStream(28)
            ImBPP(TempImageNum) = BMPbpp
            Erase BStream()
         End If
         
         If BMPbpp <> 32 Then  ' BMP (bpp <32), (BMPbpp=0 if GIF or JPG)
            picSmallBU.Picture = LoadPicture
            picSmallBU.Picture = LoadPicture(FileSpec$)
            BUW = picSmallBU.Width
            BUH = picSmallBU.Height
            If BUW > MaxWidth Or BUH > MaxHeight Then
               a$ = "Image size =" & Str$(BUW) & " x" & Str$(BUH) & " > 64 x 64" & vbCrLf & vbCrLf
               a$ = a$ & "Resize to fit Image Number" & Str$(TempImageNum + 1) & "?    " & vbCrLf
               If AspectNumber = 6 Then
                  a$ = a$ & "(Keeps aspect ratio)"
               Else
                  a$ = a$ & "(Ignores aspect ratio)"
               End If
               resp$ = MsgBox(a$, vbQuestion + vbYesNo, "Opening " & Ext$ & " file")
               If resp$ = vbYes Then
                  ' Compress into picTemp(#), set picSmallBU to new size, Blit picTemp to picSmallBU
                  '------------------------------
                  If AspectNumber = 6 Then   ' Keep aspect ratio
                     zaspectBU = BUW / BUH
                     If BUW >= BUH Then  ' zASpectBU >= 1
                        neww = ImageWidth(TempImageNum)
                        newh = ImageWidth(TempImageNum) / zaspectBU
                     Else  ' BUW < BUH   ' zAspectBU < 1
                        newh = ImageHeight(TempImageNum)
                        neww = ImageHeight(TempImageNum) * zaspectBU
                     End If
                     
                     picSmall(TempImageNum).Picture = LoadPicture
                     picSmall(TempImageNum).Width = neww
                     picSmall(TempImageNum).Height = newh
                     ImageWidth(TempImageNum) = neww
                     ImageHeight(TempImageNum) = newh
                  End If
                  '------------------------------
                  With picTemp(TempImageNum)
                     .Width = picSmall(TempImageNum).Width
                     .Height = picSmall(TempImageNum).Height
                  End With
                  ' Compress picSmallBU into picTemp
                  SetStretchBltMode picTemp(TempImageNum).hdc, HALFTONE
                  StretchBlt picTemp(TempImageNum).hdc, 0, 0, ImageWidth(TempImageNum), ImageHeight(TempImageNum), _
                     picSmallBU.hdc, 0, 0, picSmallBU.Width, picSmallBU.Height, vbSrcCopy
                  picTemp(TempImageNum).Picture = picTemp(TempImageNum).Image
                  
                  picSmallBU.Picture = LoadPicture
                  With picSmallBU
                     .Width = picSmall(TempImageNum).Width
                     .Height = picSmall(TempImageNum).Height
                  End With
                  BitBlt picSmallBU.hdc, 0, 0, ImageWidth(TempImageNum), ImageHeight(TempImageNum), _
                     picTemp(TempImageNum).hdc, 0, 0, vbSrcCopy
                  picSmallBU.Picture = picSmallBU.Image
                  ImageNum = TempImageNum
                  
                  'mnuAlphaEdit(ImageNum).Enabled = True
               Else
                  picSmallBU.Picture = LoadPicture
                  picSmallBU.Width = 8
                  picSmallBU.Height = 8
                  ImageNum = TempImageNum
                  
                  'mnuAlphaEdit(ImageNum).Enabled = True
                  
                  Exit Sub  ' Done with BMP,GIF & JPG
               End If
            End If
         End If
      
      End If
      
      ' ICO,CUR or 32bpp BMP here
         
      HotX(TempImageNum) = 0
      HotY(TempImageNum) = 0
         
         
      DisableButtons
         
      ' Have picture in picSmallBU
      Response$ = ""
      
      If Ext$ = "ICO" And NumIcons > 1 Then ' Multiple CUR files not done
         
         frmSelect.Show vbModal
         
         
         Select Case UCase$(Response$)
         Case "1", "2", "3"      ' Single if 1 then done else parse
            ResValue = Val(Response$)
            If ResValue > NumIcons Then
               MsgBox "Invalid entry", vbInformation, "Loading ICO File"
               Unload frmSelect
               FileSpec$ = ""
               EnableButtons
               Exit Sub
            End If
         
            fnum = FreeFile
            Open FileSpec$ For Binary As #fnum
            ReDim BStream(0 To LOF(fnum) - 1)
            Get #fnum, , BStream
            Close #fnum
            
            Select Case UCase$(Response$)
            Case "1"
               If Not ETools2ctICON(0, BStream(), IcoStream()) Then Unload frmSelect: GoTo OPENERROR
            Case "2"
               If Not ETools2ctICON(1, BStream(), IcoStream()) Then Unload frmSelect: GoTo OPENERROR
            Case "3"
               If Not ETools2ctICON(2, BStream(), IcoStream()) Then Unload frmSelect: GoTo OPENERROR
            End Select
            
            picSmallBU.Picture = LoadPicture
            
            ' Get image to picSmallBU
            Set picSmallBU.Picture = PictureFromByteStream(IcoStream)
            picSmallBU.Picture = picSmallBU.Image
         
            BUW = picSmallBU.Width
            BUH = picSmallBU.Height
            
            
            If Ico36bpp = 32 Then
            '"1" & 0  0
            '"2" & 0  0
            '"3" & 0  0
            '"1" & 1  1
               Select Case TempImageNum
               Case 0
                  ReDim DATACUL0(0 To 3, 0 To picSmallBU.Width - 1, 0 To picSmallBU.Height - 1)
                  FILL3D DATACUL0(), DATACULSRC()
               Case 1
                  ReDim DATACUL1(0 To 3, 0 To picSmallBU.Width - 1, 0 To picSmallBU.Height - 1)
                  FILL3D DATACUL1(), DATACULSRC()
               Case 2
                  ReDim DATACUL2(0 To 3, 0 To picSmallBU.Width - 1, 0 To picSmallBU.Height - 1)
                  FILL3D DATACUL2(), DATACULSRC()
               End Select
            End If
            
            ImageNum = TempImageNum   ' One Transfer picSmallBU to picSmall(ImageNum)
            
            ICOCURBPP(ImageNum) = Ico36bpp 'IcoStream(36)
            ImBPP(ImageNum) = Ico36bpp
            If ImBPP(ImageNum) = 32 And aAlphaRestricted Then
               BMPbpp = 24
               ImBPP(ImageNum) = 24
               ICOCURBPP(ImageNum) = 24
            End If
            If Not TransferImage(Ext$) Then
               MsgBox "Transfer error", vbCritical, "TransferImage"
               Unload frmSelect
               GoTo OPENERROR
            End If
            
            BackUp ImageNum
            
         Case "A"  ' ALL To parse
            ResValue = 0
            
            fnum = FreeFile
            Open FileSpec$ For Binary As #fnum
            ReDim BStream(0 To LOF(fnum) - 1)
            Get #fnum, , BStream
            Close #fnum
            
            For k = 0 To NumIcons - 1
               If ETools2ctICON(k, BStream(), IcoStream()) Then
                  picSmallBU.Picture = LoadPicture
                  Set picSmallBU.Picture = PictureFromByteStream(IcoStream)  ' Image in picSmallBU
                  picSmallBU.Picture = picSmallBU.Image
                  
                  BUW = picSmallBU.Width
                  BUH = picSmallBU.Height
                  
                  If Ico36bpp = 32 Then
                     Select Case k
                     Case 0
                        ReDim DATACUL0(0 To 3, 0 To picSmallBU.Width - 1, 0 To picSmallBU.Height - 1)
                        FILL3D DATACUL0(), DATACULSRC()
                     Case 1
                        ReDim DATACUL1(0 To 3, 0 To picSmallBU.Width - 1, 0 To picSmallBU.Height - 1)
                        FILL3D DATACUL1(), DATACULSRC()
                     Case 2
                        ReDim DATACUL2(0 To 3, 0 To picSmallBU.Width - 1, 0 To picSmallBU.Height - 1)
                        FILL3D DATACUL2(), DATACULSRC()
                     End Select
                  End If
                  ImageNum = k   'TempImageNum
                  ICOCURBPP(ImageNum) = Ico36bpp 'IcoStream(36)
                  ImBPP(ImageNum) = Ico36bpp
                  If ImBPP(ImageNum) And aAlphaRestricted Then
                     BMPbpp = 24
                     ImBPP(ImageNum) = 24
                     ICOCURBPP(ImageNum) = 24
                  End If
                  
                  If Not TransferImage(Ext$) Then  ' Transfers picSmallBU to picSmall(k)
                     MsgBox "Transfer error", vbCritical, "TransferImage"
                     Unload frmSelect
                     GoTo OPENERROR
                  End If
                  
                  KILLSAVS k
                  BackUp k ' From picSmall(k)
                  cmdUNDO(k).Enabled = False

               Else
                  Unload frmSelect
                  GoTo OPENERROR
               End If
               If k = 2 Then Exit For
            Next k
         Case Else
            MsgBox "Cancelled", vbInformation, "Loading ICO File"
            FileSpec$ = ""
            EnableButtons
            cmdSelect_MouseUp 0, 1, 0, 0, 0
            Exit Sub
         End Select
      
      Else   ' ICO or CUR or 32bpp BMP
         
         If Ext$ = "ICO" Or Ext$ = "CUR" Then
         
            fnum = FreeFile
            Open FileSpec$ For Binary As #fnum
            i = LOF(fnum)
            ReDim BStream(0 To i - 1)
            Get #fnum, , BStream
            Close #fnum
            'k = UBound(BStream)
            
            BStream(2) = 1
            HotX(TempImageNum) = 0
            HotY(TempImageNum) = 0
            If Ext$ = "CUR" Then
               HotX(TempImageNum) = BStream(10)
               HotY(TempImageNum) = BStream(12)
            End If
            
            If Not ETools2ctICON(0, BStream(), IcoStream()) Then
               Unload frmSelect
               ImageNum = TempImageNum
               GoTo OPENERROR
            End If
'               From ETools2ctICON: 'Original colors if bpp = 32
'               DATACULSRC(0, ix, iy) = NB
'               DATACULSRC(1, ix, iy) = NG
'               DATACULSRC(2, ix, iy) = NR
'               DATACULSRC(3, ix, iy) = NA
'               IconStream has modified RGB and Alpha = 255
            
            ' ETools2ctICON returns Ico36bpp from IconStream(36)
            ' & IcoStream().  BStream() intact.
            If Ico36bpp = 32 Then
               ' IcoStream(0)(1),,= 66 77 BM,, ie returned as a
               ' 32bpp bmp with alpha effect done against TColorBGR
               ' background
               
               ' These strictly not nec as picSmallBU autosizes
               ' but a check
               picSmallBU.Width = IcoStream(18)    ' W
               picSmallBU.Height = IcoStream(22)   ' H
            
               BUW = picSmallBU.Width
               BUH = picSmallBU.Height
               
               If Ext$ = "CUR" Then
                  If HotX(TempImageNum) < 0 Or HotX(TempImageNum) > IcoStream(18) - 1 Then
                     MsgBox "Cursor HotX wrong", vbCritical, "Opening Cursor"
                     ImageNum = TempImageNum
                     GoTo OPENERROR2
                  End If
                  If HotY(TempImageNum) < 0 Or HotY(TempImageNum) > IcoStream(22) - 1 Then
                     MsgBox "Cursor HotY wrong", vbCritical, "Opening Cursor"
                     ImageNum = TempImageNum
                     GoTo OPENERROR2
                  End If
               End If
            Else  ' Ico Cur bpp<32
               picSmallBU.Width = IcoStream(26)    ' W
               picSmallBU.Height = IcoStream(30) / 2 ' H
               If Ext$ = "CUR" Then
                  If HotX(TempImageNum) < 0 Or HotX(TempImageNum) > IcoStream(26) - 1 Then
                     MsgBox "Cursor HotX wrong", vbCritical, "Opening Cursor"
                     ImageNum = TempImageNum
                     GoTo OPENERROR2
                  End If
                  If HotY(TempImageNum) < 0 Or HotY(TempImageNum) > IcoStream(30) / 2 - 1 Then
                     MsgBox "Cursor HotY wrong", vbCritical, "Opening Cursor"
                     ImageNum = TempImageNum
                     GoTo OPENERROR2
                  End If
               End If
            End If
            
            picSmallBU.Picture = LoadPicture
            Set picSmallBU.Picture = PictureFromByteStream(IcoStream)
            picSmallBU.Picture = picSmallBU.Image
            
            BUW = picSmallBU.Width
            BUH = picSmallBU.Height
            
            
            If Ico36bpp = 32 Then   ' Fill DATACUL#() with Original colors
               Select Case TempImageNum
               Case 0
                  ReDim DATACUL0(0 To 3, 0 To BUW - 1, 0 To BUH - 1)
                  FILL3D DATACUL0(), DATACULSRC()
               Case 1
                  ReDim DATACUL1(0 To 3, 0 To BUW - 1, 0 To BUH - 1)
                  FILL3D DATACUL1(), DATACULSRC()
               Case 2
                  ReDim DATACUL2(0 To 3, 0 To BUW - 1, 0 To BUH - 1)
                   FILL3D DATACUL2(), DATACULSRC()
               End Select
            End If
            
            ICOCURBPP(TempImageNum) = Ico36bpp 'IcoStream(36)
            ImBPP(TempImageNum) = Ico36bpp
            
            If ImBPP(TempImageNum) = 32 And aAlphaRestricted Then
               BMPbpp = 24
               ImBPP(TempImageNum) = 24
               ICOCURBPP(TempImageNum) = 24
            End If
            
         ElseIf Ext$ = "BMP" And BMPbpp = 32 Then
         
         
            ImBPP(TempImageNum) = 32
            fnum = FreeFile
            Open FileSpec$ For Binary As #fnum
            i = LOF(fnum)
            ReDim BStream(0 To i - 1)
            Get #fnum, , BStream
            Close #fnum
            
            BMPWidth = BStream(18) + 256& * BStream(19)
            BMPHeight = BStream(22) + 256& * BStream(23)
            
            Select Case TempImageNum
            Case 0: ReDim DATACUL0(0 To 3, 0 To BMPWidth - 1, 0 To BMPHeight - 1)
            Case 1: ReDim DATACUL1(0 To 3, 0 To BMPWidth - 1, 0 To BMPHeight - 1)
            Case 2: ReDim DATACUL2(0 To 3, 0 To BMPWidth - 1, 0 To BMPHeight - 1)
            End Select
            
            ix = 0
            iy = 0
            
            ' Modify BGR  for alpha
            ' Set background color to transparent color
            LngToRGB TColorBGR, TR, TG, TB
            
            For k = 54 To UBound(BStream) Step 4
               ' To show alpha bytes
'               NA = CByte(BStream(k + 3))
'               BStream(k) = NA
'               BStream(k + 1) = NA
'               BStream(k + 2) = NA
               
               NB0 = CByte(BStream(k))
               NG0 = CByte(BStream(k + 1))
               NR0 = CByte(BStream(k + 2))
               NA = CByte(BStream(k + 3))
               If aAlphaRestricted Then
                  ' Only lets colors through where NA=255
                  If NA <> 255 Then NA = 0
               End If
               Select Case TempImageNum
               Case 0
                  DATACUL0(0, ix, iy) = NB0
                  DATACUL0(1, ix, iy) = NG0
                  DATACUL0(2, ix, iy) = NR0
                  DATACUL0(3, ix, iy) = NA
               Case 1
                  DATACUL1(0, ix, iy) = NB0
                  DATACUL1(1, ix, iy) = NG0
                  DATACUL1(2, ix, iy) = NR0
                  DATACUL1(3, ix, iy) = NA
               Case 2
                  DATACUL2(0, ix, iy) = NB0
                  DATACUL2(1, ix, iy) = NG0
                  DATACUL2(2, ix, iy) = NR0
                  DATACUL2(3, ix, iy) = NA
               End Select
               
               ix = ix + 1
               If ix > BMPWidth - 1 Then
                  ix = 0
                  iy = iy + 1
               End If
               
               Palpha = (NA / 255)
               NB = TB * (1 - Palpha) + NB0 * Palpha
               NG = TG * (1 - Palpha) + NG0 * Palpha
               NR = TR * (1 - Palpha) + NR0 * Palpha
               
               ' Adjust for tiny Ealpha where NB=TB when NA<>0
               If NA <> 0 Then
                  If NB = TB Then
                  If NG = TG Then
                  If NR = TR Then
                     If NB0 > TB Then
                        NB = TB + 1
                     Else
                        NB = TB - 1
                     End If
                  End If
                  End If
                  End If
               End If
               
               BStream(k) = CByte(NB)
               BStream(k + 1) = CByte(NG)
               BStream(k + 2) = CByte(NR)
               BStream(k + 3) = 255
            Next k

            
            BMPWidth = BStream(18) + 256& * BStream(19)
            BMPHeight = BStream(22) + 256& * BStream(23)
            
            picSmallBU.Picture = LoadPicture
            Set picSmallBU.Picture = PictureFromByteStream(BStream)
            picSmallBU.Picture = picSmallBU.Image
            
            ImageNum = TempImageNum   ' One Transfer picSmallBU to picSmall(ImageNum)
            If aAlphaRestricted Then
               BMPbpp = 24
               ImBPP(ImageNum) = 24
               'mnuAlphaEdit(ImageNum + 8).Enabled = True
            End If
         
         End If ' ElseIf Ext$ = "BMP" And BMPbpp = 32 Then
         
         ' BStream() now modified so original stream lost
         ImageNum = TempImageNum   ' One Transfer picSmallBU to picSmall(ImageNum)
         
         If Not TransferImage(Ext$) Then
            MsgBox "Transfer error", vbCritical, "TransferImage"
            GoTo OPENERROR2
            
         End If
         
         If Ext$ = "GIF" Or Ext$ = "JPG" Then
            ImBPP(ImageNum) = 0
         End If
         
         If ImBPP(ImageNum) <> 32 Then
            'mnuAlphaEdit(ImageNum + 8).Enabled = True
            ShowTheColorCount (ImageNum)
         End If
         
         KILLSAVS ImageNum
         BackUp ImageNum ' From picSmall(k)
         cmdUNDO(ImageNum).Enabled = False  '<<<<<<<<
      End If
'=====================================================
      ImFileSpec$(ImageNum) = FileSpec$
      
      ' RECENT FILES ENTRY
      ' NumRecentFiles used by Get & Write INI
      mnuRecentFiles(0).Visible = True   ' Break
      
      'mnuRecentFiles(1 to 8)
      'FileArray$(1 to 8)
      
'      ' Check if FileSpec$ already in list ie in FileArray$(k)
      For k = 1 To FileCount
         If FileSpec$ = FileArray$(k) Then Exit For
      Next k
      If k > FileCount Then  ' Not in list, add it in
                  
         If NumRecentFiles < MaxFileCount Then
            NumRecentFiles = NumRecentFiles + 1
            FileCount = FileCount + 1
         End If
         For k = 7 To 1 Step -1
            If Len(FileArray$(k)) > 0 Then
               FileArray$(k + 1) = FileArray$(k)
            End If
         Next k
         FileArray$(1) = FileSpec$

         For k = 1 To 8
            If Len(FileArray$(k)) > 0 Then
               TheNAM$ = GetFileName(FileArray$(k))
               ThePATH$ = GetPath(FileArray$(k))
               a$ = TheNAM$ & ": " & ThePATH$
               mnuRecentFiles(k).Caption = LTrim$(Str$(k)) & ". " & ShortenString$(a$, 42)
               'mnuRecentFiles(k).Caption = LTrim$(Str$(k)) & ". " & ShortenString$(FileArray$(k), 40)
               mnuRecentFiles(k).Visible = True
            End If
         Next k
      End If

   End If   ' If Len(FileSpec$) = 0 Then
   On Error GoTo 0
   EnableButtons
   Exit Sub
'=========
OPENERROR:
   MsgBox "Failed to open. File error or Vista icon or a size > 64 pixels. " _
           & vbCrLf & "or File or Folder name changed. Could try Extractor menu." _
           & vbCrLf & "bpp=" & Str$(ICOBPP(0)) _
           & vbCrLf & FileSpec$ & _
           "  ", vbCritical, "Opening file"
   'mnuAlphaEdit(ImageNum + 4).Enabled = False

OPENERROR2:
   FileSpec$ = ""
   'Err.Clear
   picSmallBU.Picture = LoadPicture
   picSmall(ImageNum).Picture = LoadPicture
   With picSmallBU
      .Width = 16
      .Height = 16
   End With
   With picSmall(ImageNum)
      .Width = 16
      .Height = 16
   End With
   ImageWidth(ImageNum) = 16
   ImageHeight(ImageNum) = 16
   scrWidth.Value = ImageWidth(ImageNum)
   scrHeight.Value = 65 - ImageHeight(ImageNum)
   scrHeight_Scroll
   scrWidth_Scroll
   DrawGrid
   With picORG(ImageNum)
      .Width = picSmallBU.Width
      .Height = picSmallBU.Height
   End With
   BitBlt picORG(ImageNum).hdc, 0, 0, picSmallBU.Width, picSmallBU.Height, _
          picSmallBU.hdc, 0, 0, vbSrcCopy
   picORG(ImageNum).Picture = picORG(ImageNum).Image
   EnableButtons
   Exit Sub
End Sub

Private Sub DisableButtons()
Dim k As Long
   
   For k = 0 To optTools.Count - 1    ' Disable all Tools
      optTools(k).Enabled = False
   Next k
End Sub

Private Sub EnableButtons()
Dim k As Long
   
   For k = 0 To optTools.Count - 1
      optTools(k).Enabled = True
   Next k
   ' Move/CopySelect disabled
   optTools([MoveSEL]).Enabled = False
   optTools([MoveCOPY]).Enabled = False
End Sub

Private Function TransferImage(Ext$) As Boolean
Dim Capt$
   
   On Error GoTo TransferError
   ' From mnuOpenIntoImage  picSmallBU - > picSmall
   With picSmall(ImageNum)
      .Width = picSmallBU.Width
      .Height = picSmallBU.Height
   End With
   
   BitBlt picSmall(ImageNum).hdc, 0, 0, picSmallBU.Width, picSmallBU.Height, _
          picSmallBU.hdc, 0, 0, vbSrcCopy
   ' Can change TColorBGR !!??
   
   picSmall(ImageNum).Picture = picSmall(ImageNum).Image
   
   ImageHeight(ImageNum) = picSmall(ImageNum).Height
   ImageWidth(ImageNum) = picSmall(ImageNum).Width
   
   scrWidth.Value = ImageWidth(ImageNum)
   scrHeight.Value = 65 - ImageHeight(ImageNum)
   scrHeight_Scroll
   scrWidth_Scroll
   ShowTheColorCount ImageNum
   cmdSelect_MouseUp CInt(ImageNum), 1, 0, 0, 0
   DrawGrid
   
   ' Transfer picSmallBU - > picORG for Reload original
   With picORG(ImageNum)
      .Width = picSmallBU.Width
      .Height = picSmallBU.Height
   End With
   BitBlt picORG(ImageNum).hdc, 0, 0, picSmallBU.Width, picSmallBU.Height, _
          picSmallBU.hdc, 0, 0, vbSrcCopy
   picORG(ImageNum).Picture = picORG(ImageNum).Image
   
   Capt$ = GetFileName(FileSpec$) & " (" & FileLen(FileSpec$) & " B)"
   If Ext$ = "ICO" Or Ext$ = "CUR" Then
      Capt$ = Capt$ & " ( bpp =" & Str$(ICOCURBPP(ImageNum)) & ")"
   ElseIf Ext$ = "BMP" Then
      Capt$ = Capt$ & " ( bpp =" & Str$(BMPbpp) & ")"
   ElseIf Ext$ = "JPG" Then
      Capt$ = Capt$ & " ( bpp = 24)"
   ElseIf Ext$ = "GIF" Then
      Capt$ = Capt$ & " ( bpp = 8)"
   End If
   
   Capt$ = "  " & Capt$
   NameSpec$(ImageNum) = Capt$
   ImFileSpec$(ImageNum) = FileSpec$

   LabSpec = Capt$ & "  Path: " & ShortenString$(GetPath(ImFileSpec$(ImageNum)), 64) '74)
   SetAlphaEditMenu

   KILLSAVS ImageNum
   
   cmdUNDO(ImageNum).Enabled = False

   LabCurGen(ImageNum) = CurrentGen(ImageNum)
   LabMaxGen(ImageNum) = MaxGen(ImageNum)

   mnuReloadImage(ImageNum).Enabled = True
   cmdClear(ImageNum).Enabled = True
   cmdReload(ImageNum).Enabled = True
   TransferImage = True
   Exit Function
'=======
TransferError:
   Exit Function
End Function

'#### Saving BMP ####

Private Sub mnuSaveBMPImage_Click(Index As Integer)
Dim r As RECT
Dim Title$, Filt$, InDir$
Dim FIndex As Long
Dim Cul As Long
Dim CulC As Long
Dim Color1 As Long, Color2 As Long
Dim ix As Long, iy As Long
Dim resp As Long

   If Index = 4 Then Exit Sub
   
   If Index <> 3 Then   ' ie not saving drawing panel
      If IsBlank(Index) Then
         resp = MsgBox("Blank image" & Str$(Index + 1) & " ?!" & vbCrLf & "Continue?", vbQuestion + vbYesNo, "Saving BMP")
         If resp = vbNo Then Exit Sub
      End If
   End If

   If Index < 3 Then
      cmdSelect_MouseUp Index, 1, 0, 0, 0 ' Sets ImageNum = Index
   End If
   
   ' Could ask if original colors acceptable
   If aQueryOC And ImBPP(ImageNum) = 32 Then
      resp = MsgBox("Do you want to check the Original Colors before saving?  ", vbQuestion + vbYesNo, "Saving 32bpp BMP")
      If resp = vbYes Then
         MsgBox "Goto Alpha menu & edit Original Colors for image" & Str$(ImageNum + 1) & " ", vbOKOnly, "Saving 32bpp BMP"
         Exit Sub
      End If
   End If
   
   If Index < 3 Then
      Title$ = "Save Image" & Str$(ImageNum + 1) & " As BMP file"
   Else
      Title$ = "Save Drawing panel As BMP file"
   End If
   Filt$ = "Windows Bitmap (*.bmp)|*.bmp"
   SaveSpec$ = ""
   InDir$ = SavePath$ 'AppPathSpec$
   Set CommonDialog1 = New cOSDialog
   CommonDialog1.ShowSave SaveSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
   Set CommonDialog1 = Nothing
   
   ' Avoid click through
   GetWindowRect Form1.picSmallFrame.hwnd, r
   SetCursorPos r.Left + 130, r.Top + 90
   
   If Len(SaveSpec$) = 0 Then
      Exit Sub
   Else
      SavePath$ = SaveSpec$
      FixExtension SaveSpec$, ".bmp"
      picSmall(ImageNum).Picture = picSmall(ImageNum).Image
      
      If Index = 3 Then
         ' Save gridded image  as 24bpp
         SavePicture picPANEL.Image, SaveSpec$  ' 24 bpp pic for now
         Exit Sub
      End If
      
      svOptimizeNumber = OptimizeNumber
      If ImBPP(ImageNum) = 32 Then OptimizeNumber = 0
      
      If OptimizeNumber = 1 Then  ' Not optimized saving
         ' Save image ImageNum as 24bpp
         SavePicture picSmall(ImageNum).Image, SaveSpec$  ' 24 bpp pic for now
      Else   ' Optimize
            CulC = GetColorCount(picSmall(ImageNum))
            If CulC <= 256 And ImBPP(ImageNum) <> 32 Then
               
               GetTheBitsBGR ImageNum, picSmall(ImageNum)
               ' IE
               ' Gets picDATAREG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
               ' from picSmall(ImageNum)
               If aDIBError Then Exit Sub
               
               Select Case CulC
               Case Is <= 2
                  ' Find 2 colors
                  Color1 = RGB(picDATAREG(2, 0, 0), picDATAREG(1, 0, 0), picDATAREG(0, 0, 0))
                  Color2 = Color1
                  For iy = 0 To ImageHeight(ImageNum) - 1
                  For ix = 0 To ImageWidth(ImageNum) - 1
                     Cul = RGB(picDATAREG(2, ix, iy), picDATAREG(1, ix, iy), picDATAREG(0, ix, iy))
                     If Cul <> Color2 Then
                        Color2 = Cul
                        Exit For
                     End If
                  Next ix
                  If ix < ImageWidth(ImageNum) Then Exit For
                  Next iy
                  ' Have Color1, Color2
                  If Not ModSave2BMP.SaveBMP2(SaveSpec$, picDATAREG(), ImageWidth(ImageNum), ImageHeight(ImageNum), Color1, Color2) Then
                     MsgBox "Error in saving 2 color bmp", vbInformation, "Saving 1bpp BMP"
                     OptimizeNumber = svOptimizeNumber
                     Exit Sub
                  End If
               
               Case Is <= 16
                  GetPalette picDATAREG() ' Returns gifPAL(0 To 255) (0 To 15)
                  If Not ModSave16BMP.SaveBMP16(SaveSpec$, picDATAREG(), ImageWidth(ImageNum), ImageHeight(ImageNum), gifPAL()) Then
                     MsgBox "Error in saving 16 color bmp", vbInformation, "Saving 4bpp BMP"
                     OptimizeNumber = svOptimizeNumber
                     Exit Sub
                  End If
               
               Case Is <= 256
                  GetPalette picDATAREG() ' Returns gifPAL(0 To 255)
                  If Not ModSave256BMP.SaveBMP256(SaveSpec$, picDATAREG(), ImageWidth(ImageNum), ImageHeight(ImageNum), gifPAL()) Then
                     MsgBox "Error in saving 256 color bmp", vbInformation, "Saving 8bpp BMP"
                     OptimizeNumber = svOptimizeNumber
                     Exit Sub
                  End If
               End Select
            Else  'CulC > 256 Or ImBPP(ImageNum) = 32 Then
               If ImBPP(ImageNum) = 24 Then
                  ' Save image ImageNum as 24bpp
                  SavePicture picSmall(ImageNum).Image, SaveSpec$  ' 24 bpp pic
                  
               ElseIf ImBPP(ImageNum) = 32 Then
                  
                  'GetOriginalColors ImageNum  ' Copies DATACUL#() to picDATAREG()
                  
                  ' Save DATACUL#() instead
                  If Not Save32bppBMP(SaveSpec$, ImageWidth(ImageNum), ImageHeight(ImageNum)) Then
                     MsgBox "Error in saving 32bpp bmp", vbInformation, "Saving 32bpp BMP"
                     OptimizeNumber = svOptimizeNumber
                     Exit Sub
                  End If
               Else   ' Error
                  OptimizeNumber = svOptimizeNumber
                  MsgBox "Error saving BMP in mnuSaveBMPImage", vbCritical, "Saving BMP"
               End If
            End If   ' If CulC <= 256 And ImBPP(ImageNum) <> 32 Then
      End If
      OptimizeNumber = svOptimizeNumber
   End If
End Sub

'#### Save CUR ####

Private Sub mnuSaveCURImage_Click(Index As Integer)

' Exactly the same as icon apart from
' Offset  2(Loc3)  2 instead of 1 Integer
' Offset 10(Loc12) HotX instead of 1 nump (but 1 Not reliable for ico)
' Offset 12(Loc13) HotY instead of 8 bpp  (but 8 Not reliable for ico)
Dim r As RECT
Dim Title$, Filt$, InDir$
Dim FIndex As Long
Dim resp As Long

   If Index = 3 Then Exit Sub
   
   If IsBlank(Index) Then
      resp = MsgBox("Blank image" & Str$(Index + 1) & " ?!" & vbCrLf & "Continue?", vbQuestion + vbYesNo, "Saving CUR")
      If resp = vbNo Then Exit Sub
   End If
   
   
   cmdSelect_MouseUp Index, 1, 0, 0, 0   ' Does ImageNum = Index
   
   If aQueryOC And ImBPP(ImageNum) = 32 Then
      resp = MsgBox("Do you want to check the Original Colors before saving?  ", vbQuestion + vbYesNo, "Saving 32bpp CUR")
      If resp = vbYes Then
         MsgBox "Goto Alpha menu & edit Original Colors for image" & Str$(ImageNum + 1) & " ", vbOKOnly, "Saving 32bpp CUR"
         Exit Sub
      End If
   End If


   Title$ = "Save Image" & Str$(Index + 1) & " As CUR file"
   Filt$ = "CURSOR (*.cur)|*.cur"
   SaveSpec$ = ""
   InDir$ = SavePath$ 'AppPathSpec$
   Set CommonDialog1 = New cOSDialog
   CommonDialog1.ShowSave SaveSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
   Set CommonDialog1 = Nothing
   
   ' Avoid click through
   GetWindowRect Form1.picSmallFrame.hwnd, r
   SetCursorPos r.Left + 130, r.Top + 90
   
   If Len(SaveSpec$) = 0 Then
      Exit Sub
   Else
      SavePath$ = SaveSpec$
      SaveIndex = 0
      svOptimizeNumber = OptimizeNumber
      If ImBPP(Index) = 32 Then OptimizeNumber = 0
      
      SaveSingleICOorCUR SaveSpec$, 2    ' 2 = CUR
      OptimizeNumber = svOptimizeNumber
   End If

End Sub

'#### Save ICO ####

Private Sub mnuSaveICOImage_Click(Index As Integer)
Dim r As RECT
Dim Title$, Filt$, InDir$
Dim FIndex As Long
Dim kup As Long
Dim Im As Integer
Dim resp As Long

   If Index = 5 Then Exit Sub
   
   If Index < 3 Then
      Im = Index
   Else
      Im = 0
   End If
      
   If IsBlank(Im) Then
      resp = MsgBox("Blank image" & Str$(Im + 1) & " ?!" & vbCrLf & "Continue?", vbQuestion + vbYesNo, "Saving ICO")
      If resp = vbNo Then Exit Sub
   End If
   
   Im = Index
   
   If Index > 2 Then Im = 0
   cmdSelect_MouseUp Im, 1, 0, 0, 0   ' Does ImageNum = Im
   
   If aQueryOC And ImBPP(ImageNum) = 32 Then
      resp = MsgBox("Do you want to check the Original Colors before saving?  ", vbQuestion + vbYesNo, "Saving 32bpp ICO(s)")
      If resp = vbYes Then
         MsgBox "Goto Alpha menu & edit Original Colors for image" & Str$(ImageNum + 1) & " ", vbOKOnly, "Saving 32bpp ICO(s)"
         Exit Sub
      End If
   End If
   
   If Index < 3 Then
      Title$ = "Save Image" & Str$(Index + 1) & " As ICO file"
   ElseIf Index = 3 Then
      ' 0 &  1
      Title$ = "Save Images" & Str$(Index - 2) & " &" & Str$(Index - 1) & " As ICO file"
   ElseIf Index = 4 Then
      ' 0 & 1 & 2
      Title$ = "Save Images" & Str$(Index - 3) & "," & Str$(Index - 2) & " &" & Str$(Index - 1) & " As ICO file"
   End If
   
   Filt$ = "ICON (*.ico)|*.ico"
   SaveSpec$ = ""
   InDir$ = SavePath$ 'AppPathSpec$
   Set CommonDialog1 = New cOSDialog
   CommonDialog1.ShowSave SaveSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
   Set CommonDialog1 = Nothing
   
   ' Avoid click through
   GetWindowRect Form1.picSmallFrame.hwnd, r
   SetCursorPos r.Left + 130, r.Top + 90
   
   If Len(SaveSpec$) = 0 Then
      Exit Sub
   Else
      SavePath$ = SaveSpec$
      
      If FileExists(SaveSpec$) Then
         Kill SaveSpec$
      End If
      
      SaveIndex = Index
      
      svOptimizeNumber = OptimizeNumber
      If ImBPP(0) = 32 Or ImBPP(1) = 32 Or _
         ImBPP(2) = 32 Then
         OptimizeNumber = 0
      End If
      
      Select Case SaveIndex
      Case 0, 1, 2   ' ImageNums
         SaveSingleICOorCUR SaveSpec$, 1   ' 1 = ICO
      Case 3   ' ImageNums 0 & 1
         kup = 1
         SaveMultiICOs SaveSpec$, kup  ' For k = 0 & 1
      Case 4   ' ImageNums 0, 1 & 2
         kup = 2
         SaveMultiICOs SaveSpec$, kup  ' For k = 0, 1 & 2
      End Select
      OptimizeNumber = svOptimizeNumber
   End If
End Sub

Private Sub SaveSingleICOorCUR(SSpec$, ICOCUR As Integer)
' ICOCUR 1 Icon, 2 Cursor
Dim k As Long
Dim fnum As Long
Dim NColors As Long
      
      If FileExists(SSpec$) Then
         Kill SaveSpec$
      End If
      
      fnum = FreeFile
      Open SSpec$ For Binary As fnum
      
      ' 8bpp  Not optimized
      With IcoCurHdr
         .ires = 0
         .ityp = ICOCUR    '1 ICO, 2 CUR
         .inum = 1    ' Single image
      End With
      Put #fnum, , IcoCurHdr
      ' Optimize can change these
      XORW(ImageNum) = (ImageWidth(ImageNum) + 3) And &HFFFFFFFC ' To 4 byte boundary
      ANDW(ImageNum) = ((ImageWidth(ImageNum) + 7) \ 8 + 3) And &HFFFFFFFC ' To 4 byte boundary !!
      XORSize(ImageNum) = XORW(ImageNum) * ImageHeight(ImageNum)
      ANDSize(ImageNum) = ANDW(ImageNum) * ImageHeight(ImageNum)
      PaletteSize(ImageNum) = 1024
      IcoCurBMIH(ImageNum).ibpp = 8
      
      If OptimizeNumber = 0 Then  ' Optimized saving
         If ImBPP(ImageNum) <> 32 Then
            NColors = GetColorCount(picSmall(ImageNum))
            Select Case NColors
            Case Is <= 2   ' 1 bpp
               XORW(ImageNum) = ((ImageWidth(ImageNum) + 7) \ 8 + 3) And &HFFFFFFFC
               PaletteSize(ImageNum) = 8
               IcoCurBMIH(ImageNum).ibpp = 1
            Case Is <= 16  ' 4 bpp
               XORW(ImageNum) = ((ImageWidth(ImageNum) + 1) \ 2 + 3) And &HFFFFFFFC
               PaletteSize(ImageNum) = 64
               IcoCurBMIH(ImageNum).ibpp = 4
            Case Is <= 256 ' 8 bpp
               XORW(ImageNum) = (ImageWidth(ImageNum) + 3) And &HFFFFFFFC ' To 4 byte boundary
               PaletteSize(ImageNum) = 1024
               IcoCurBMIH(ImageNum).ibpp = 8
            Case Else   ' 24 bpp
               k = (3 * ImageWidth(ImageNum) + 3) And &HFFFFFFFC
               k = k - 3 * ImageWidth(ImageNum)
               XORW(ImageNum) = 3 * ImageWidth(ImageNum) + k
               PaletteSize(ImageNum) = 0
               IcoCurBMIH(ImageNum).ibpp = 24
            End Select
            ANDW(ImageNum) = ((ImageWidth(ImageNum) + 7) \ 8 + 3) And &HFFFFFFFC ' To 4 byte boundary
            XORSize(ImageNum) = XORW(ImageNum) * ImageHeight(ImageNum)
            ANDSize(ImageNum) = ANDW(ImageNum) * ImageHeight(ImageNum)
         Else  ' IMBPP(ImageNum)=32
            XORW(ImageNum) = 4 * ImageWidth(ImageNum)
            PaletteSize(ImageNum) = 0
            IcoCurBMIH(ImageNum).ibpp = 32
            ANDW(ImageNum) = ((ImageWidth(ImageNum) + 7) \ 8 + 3) And &HFFFFFFFC ' To 4 byte boundary
            XORSize(ImageNum) = XORW(ImageNum) * ImageHeight(ImageNum)           ' eg W=48 ANDA=6 then 8
            ANDSize(ImageNum) = ANDW(ImageNum) * ImageHeight(ImageNum)           ' 48*48*8 = 384
      
         End If
      End If
      
      With IcoCurInfo(ImageNum)
         .bWidth = CByte(ImageWidth(ImageNum))
         .bHeight = CByte(ImageHeight(ImageNum))
         .Bres1 = 0
         .Bres2 = 0
         If ICOCUR = 1 Then
            .iHotX = 1    '  for icon
            .iHotY = IcoCurBMIH(ImageNum).ibpp    '  for icon
         Else
            .iHotX = HotX(ImageNum)    '  for cursor
            .iHotY = HotY(ImageNum)    '  for cursor
         End If
         .LSize = 40 + PaletteSize(ImageNum) + XORSize(ImageNum) + ANDSize(ImageNum)
         .Loffset = 22  ' Offset to image for cursor
      End With
      
      With IcoCurBMIH(ImageNum)
         .LBIH = 40
         .LWidth = ImageWidth(ImageNum)
         .LHeight = 2 * ImageHeight(ImageNum)
         .inump = 1
         '.ibpp = 8
         ' +6 long zeros
      End With
      Put #fnum, , IcoCurInfo(ImageNum)
      Put #fnum, , IcoCurBMIH(ImageNum)
      
      ProcessOutFile ImageNum '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
      
      If IcoCurBMIH(ImageNum).ibpp < 24 Then
         Put #fnum, , gifPAL()
      End If
      Put #fnum, , XORARR()
      Put #fnum, , ANDARR()
      Close #fnum
      
      
      If ICOCUR = 2 Then
         ' Set cursor scroll-bars
         scrHotX.Value = HotX(ImageNum)
         LabHotXY(0) = "X=" & HotX(ImageNum)
         scrHotY.Value = HotY(ImageNum)
         LabHotXY(1) = "Y=" & HotY(ImageNum)
      End If
End Sub

Private Sub SaveMultiICOs(SSpec$, kup As Long)
' kup = 1 for ImagesNums 0 & 1
' kup = 2 for ImagesNums 0, 1 & 2
Dim k As Long, kk As Long
Dim fnum As Long
Dim NColors As Long
Dim svImageNum As Long
      
      svImageNum = ImageNum
               
      If FileExists(SSpec$) Then
         Kill SSpec$
      End If
      fnum = FreeFile
      Open SSpec$ For Binary As fnum
      
      For k = 0 To kup
         XORW(k) = (ImageWidth(k) + 3) And &HFFFFFFFC ' To 4 byte boundary
         ANDW(k) = ((ImageWidth(k) + 7) \ 8 + 3) And &HFFFFFFFC ' To 4 byte boundary
         XORSize(k) = XORW(k) * ImageHeight(k)
         ANDSize(k) = ANDW(k) * ImageHeight(k)
         PaletteSize(k) = 1024
         IcoCurBMIH(k).ibpp = 8
      Next k
      
      If OptimizeNumber = 0 Then   ' Optimized saving
         For k = 0 To kup
            If ImBPP(k) <> 32 Then
               ImageNum = k
               NColors = GetColorCount(picSmall(k))
               Select Case NColors
               Case Is <= 2   ' 1 bpp
                  XORW(k) = ((ImageWidth(k) + 7) \ 8 + 3) And &HFFFFFFFC
                  PaletteSize(k) = 8
                  IcoCurBMIH(k).ibpp = 1
               Case Is <= 16  ' 4 bpp
                  XORW(k) = ((ImageWidth(k) + 1) \ 2 + 3) And &HFFFFFFFC
                  PaletteSize(k) = 64
                  IcoCurBMIH(k).ibpp = 4
               Case Is <= 256 ' 8 bpp
                  XORW(k) = (ImageWidth(k) + 3) And &HFFFFFFFC ' To 4 byte boundary
                  PaletteSize(k) = 1024
                  IcoCurBMIH(k).ibpp = 8
               Case Else   ' 24 bpp
                  kk = (3 * ImageWidth(k) + 3) And &HFFFFFFFC
                  kk = kk - 3 * ImageWidth(k)
                  XORW(k) = 3 * ImageWidth(k) + kk
                  PaletteSize(k) = 0
                  IcoCurBMIH(k).ibpp = 24
               End Select
               ANDW(k) = ((ImageWidth(k) + 7) \ 8 + 3) And &HFFFFFFFC ' To 4 byte boundary
               XORSize(k) = XORW(k) * ImageHeight(k)
               ANDSize(k) = ANDW(k) * ImageHeight(k)
            Else ' BPP = 32
               XORW(k) = 4 * ImageWidth(k)
               PaletteSize(k) = 0
               IcoCurBMIH(k).ibpp = 32
               ANDW(k) = ((ImageWidth(k) + 7) \ 8 + 3) And &HFFFFFFFC ' To 4 byte boundary
               XORSize(k) = XORW(k) * ImageHeight(k)
               ANDSize(k) = ANDW(k) * ImageHeight(k)
            End If
         Next k
      End If
      
      With IcoCurHdr
         .ires = 0
         .ityp = 1
         .inum = kup + 1  ' Number of images for icon (2 or 3)
      End With
      Put #fnum, , IcoCurHdr

      For k = 0 To kup
         With IcoCurInfo(k)
            .bWidth = CByte(ImageWidth(k))
            .bHeight = CByte(ImageHeight(k))
            .Bres1 = 0
            .Bres2 = 0
            .iHotX = 1    ' NPlanes  ,iHotX for cursor
            .iHotY = IcoCurBMIH(k).ibpp   ' bpp      ,iHotY for cursor
            .LSize = 40 + PaletteSize(k) + XORSize(k) + ANDSize(k)
            If kup = 1 Then   '' 2 Icons
               Select Case k
               Case 0: .Loffset = 38
               Case 1: .Loffset = 38 + IcoCurInfo(0).LSize
               End Select
            Else
               Select Case k  '' 3 Icons
               Case 0: .Loffset = 54
               Case 1: .Loffset = 54 + IcoCurInfo(0).LSize
               Case 2: .Loffset = 54 + IcoCurInfo(0).LSize + IcoCurInfo(1).LSize
               End Select
            End If
         End With
         Put #fnum, , IcoCurInfo(k)
      Next k
         
'Public Type IcoCurBIH
'   LBIH As Long         ' size of BMIH (40)
'   LWidth As Long       ' width
'   LHeight As Long      ' 2 * height
'   inump As Integer     ' num planes = 1
'   ibpp As Integer      ' bpp (1,4,8,24)
'   L1 As Long           ' 0
'   L2 As Long           ' 0
'   L3 As Long           ' 0
'   L4 As Long           ' 0
'   L5 As Long           ' 0
'   L6 As Long           ' 0
'End Type                ' Length 40 B
'Public IcoCurBMIH(2) As IcoCurBIH
         
      For k = 0 To kup
         With IcoCurBMIH(k)
            .LBIH = 40
            .LWidth = ImageWidth(k)
            .LHeight = 2 * ImageHeight(k)
            .inump = 1
            .ibpp = IcoCurBMIH(k).ibpp
            ' +6 long zeros
         End With
         Put #fnum, , IcoCurBMIH(k)
            
         ProcessOutFile k
         
         If IcoCurBMIH(k).ibpp < 24 Then
            Put #fnum, , gifPAL()
         End If
         Put #fnum, , XORARR()
         Put #fnum, , ANDARR()
         
      Next k
      Close #fnum
      
      ImageNum = svImageNum
End Sub


Private Sub ProcessOutFile(ImNum As Long)
' Called by mnuSaveCURImage_Click
'           mnuSaveICOImage_Click
'           SaveSingleICOorCUR
'           SaveMultiICOs

' ImNUm = ImageNum
Dim ix As Long, iy As Long
Dim Cul As Long
Dim BMIH As BITMAPINFOHEADER
Dim sh As Long, sw As Long
Dim NColors As Long
Dim POW As Long, BPOW As Byte
Dim tempXORARR() As Byte
Dim n As Long
Dim k1 As Long, k2 As Long
Dim svImageNum

Dim ABYTE As Byte
Dim CorgR As Byte, CorgG As Byte, CorgB As Byte

   ' Transfer pixel colors to picDATAREG() from picSmall
   ' picDATAREG(0 To 3, 0 To ImageWidth(ImNum) - 1, 0 To ImageHeight(ImNum) - 1)
   '===========================================================
   GetTheBitsBGR ImNum, picSmall(ImNum)
   '===========================================================
   If aDIBError Then
      svOptimizeNumber = OptimizeNumber
      Exit Sub
   End If
   
   'LngToRGB TColorBGR, SR, SG, SB
   'TColorRGB = RGB(SB, SG, SR)
   
   'Build_picSmallDATA_picMaskData Form1.picSmall() 'Public picSmallDATA() Public picMaskDATA() from Form1.picSmall()
   
   sw = picSmall(ImNum).Width
   sh = picSmall(ImNum).Height
   ReDim picMaskDATA(0 To sw - 1, 0 To sh - 1)
   ReDim picSmallDATA(0 To sw - 1, 0 To sh - 1)
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = sw
      .biHeight = sh
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With
   If GetDIBits(Form1.hdc, picSmall(ImNum).Image, 0, sh, picSmallDATA(0, 0), BMIH, 0) = 0 Then
      MsgBox "DIB ERROR", vbCritical, "ProcessOutFile"
   End If

   If aDIBError Then
      svOptimizeNumber = OptimizeNumber
      Exit Sub
   End If
   
   ' Build temp mask ie where color = transparent color
   ' Makes a new Mask and Mask set =1 when Alpha = 0
   ' When Alpha=0 the color = TColorRGB
   
   For iy = 0 To sh - 1
   For ix = 0 To sw - 1
      picMaskDATA(ix, iy) = 0
      Cul = picSmallDATA(ix, iy) And &HFFFFFF  ' NB 4byte nums can be -ve in VB
      If Cul >= 0 Then
         If Cul = TColorRGB Then
            picMaskDATA(ix, iy) = 1
         End If
      End If
   Next ix
   Next iy
   
   Get_Pal_Indexes picDATAREG(), 256, True ' ReDims gifPAL(0 To 255) & Returns gifPAL() & bA8() indexes
'  Output: Public gifPAL(0 to 255) Longs                                          ' 1024 RGB palette
'          Public bA8(0 To UBound(picDATAREG(), 2), 0 To UBound(picDATAREG(), 3)) ' Indexes(bytes) to palette
'                              width                        height
   svImageNum = ImageNum
   
   ' 256 colors  8bpp
   'XORW = (ImageWidth(ImNum) + 3) And &HFFFFFFFC ' To 4 byte boundary
   ReDim XORARR(0 To XORW(ImNum) - 1, 0 To ImageHeight(ImNum) - 1) ' byte array
   ReDim tempXORARR(0 To ImageWidth(ImNum) - 1, 0 To ImageHeight(ImNum) - 1) ' byte array
   
   ' Transfer bA8() indexes to tempXORARR()
   For iy = 0 To ImageHeight(ImNum) - 1
   For ix = 0 To ImageWidth(ImNum) - 1
      tempXORARR(ix, iy) = bA8(ix, iy)
   Next ix
   Next iy
   
   ' Same for <= 8bpp
   ReDim ANDARR(0 To ANDW(ImNum) - 1, 0 To ImageHeight(ImNum) - 1) ' 0's in Byte array
'TEST
'GoTo SkipANDMask
   
   ' Build AND Mask
   ' Bits in ANDARR():0 block(Show XORARR() color),
   '                  1 transparent(Show background color)
   '     so only need 1's where color was = TColorBGR in XORARR()
   '     ie where picMaskDATA(ix, iy) = 1
   ' Find index where gifPAL(ImNum) = TColorBGR
   'For iy = ImageHeight(ImNum) - 1 To 0 Step -1
   For iy = 0 To ImageHeight(ImNum) - 1
   For ix = 0 To ImageWidth(ImNum) - 1
      Cul = picMaskDATA(ix, iy) And &HFFFFFF  ' NB 4byte nums can be -ve in VB
      ' Possible output reconciliation here see end of Form
      If Cul = 1 Then
         ' set bit in ANDARR bit ix row iy AndARR(?,iy)
         ' ? = byte pos = ix\8  ANDARR(ix\8,iy)
         ' start bit of ix\8 = (8 * ix\8)
         ' bit to set in ix\8 = 7 - (ix - (8 * (ix\8)))
         ' orbyte= 2 ^ bitset
         ' ANDARR(ix\8,iy) = ANDARR(ix\8,iy) Or orbyte
         POW = 7 - (ix - (8 * (ix \ 8)))  ' when ix=0  POW=7 & BPOW = 2^7 = 128 10000000
                                          ' when ix=14 POW=1 & BPOW = 2^1 = 2   00000010
                                          ' when ix=15 POW=0 & BPOW = 2^0 = 1   00000001
         BPOW = CByte(2 ^ POW)
         ANDARR(ix \ 8, iy) = ANDARR(ix \ 8, iy) Or BPOW
      End If
   Next ix
   Next iy
   
'SkipANDMask:
   
   If OptimizeNumber = 0 Then ' Optimized saving
      
      ImageNum = ImNum
      NColors = GetColorCount(picSmall(ImNum))
      
      If ImBPP(ImNum) < 32 Then
      
         Select Case NColors
         
         Case Is <= 2   ' Build Image 1bpp ' Recast XORARR() Indexes as bits. Redim Preseve gifPAL(0 to 1)
            ' tempXORARR() - > XORARR()  tempXORARR() contains 0 or 1
            ReDim Preserve gifPAL(0 To 1)
            ' Set any TColorBGR to 000 in gifPAL
            For k1 = 0 To 1
               If gifPAL(k1) = TColorRGB Then
                  gifPAL(k1) = 0
               End If
            Next k1
            
            ReDim XORARR(0 To (XORW(ImNum) - 1), 0 To ImageHeight(ImNum) - 1)  ' bit array
            For iy = ImageHeight(ImNum) - 1 To 0 Step -1
            For ix = 0 To ImageWidth(ImNum) - 1
               If tempXORARR(ix, iy) = 1 Then   ' Similar logic to ANDARR
                  POW = 7 - (ix - (8 * (ix \ 8)))
                  BPOW = CByte(2 ^ POW)
                  XORARR(ix \ 8, iy) = XORARR(ix \ 8, iy) Or BPOW
               End If
            Next ix
            Next iy
            IcoCurBMIH(ImNum).ibpp = 1
         
         Case Is <= 16  ' Build Image 4bpp ' Recast XORARR() as nybbles Redim Preseve gifPAL(0 to 63)
            ' tempXORARR() - > XORARR()  tempXORARR() contains 0 to 15
            ReDim Preserve gifPAL(0 To 15)
            ' Set any TColorBGR to 000 in gifPAL
            For k1 = 0 To 15
               If gifPAL(k1) = TColorRGB Then
                  gifPAL(k1) = 0
               End If
            Next k1
            
            ReDim XORARR(0 To (XORW(ImNum) - 1), 0 To ImageHeight(ImNum) - 1)  ' nybble array
            For iy = ImageHeight(ImNum) - 1 To 0 Step -1
            n = 0
            For ix = 0 To ImageWidth(ImNum) - 1 Step 2
               ' k1 = 0-15  k2= 0-15
               k1 = 16 * tempXORARR(ix, iy)
               If ix + 1 > ImageWidth(ImNum) - 1 Then
                  k2 = 0
               Else
                  k2 = tempXORARR(ix + 1, iy)
               End If
               XORARR(n, iy) = CByte(k1) + CByte(k2)
               n = n + 1
            Next ix
            Next iy
            IcoCurBMIH(ImNum).ibpp = 4
         
         Case Is <= 256 ' 256 colors  8bpp  Done!
            ' Set any TColorBGR to 000 in gifPAL
            For k1 = 0 To 255
               If gifPAL(k1) = TColorRGB Then
                  gifPAL(k1) = 0
               End If
            Next k1
            ReDim XORARR(0 To XORW(ImNum) - 1, 0 To ImageHeight(ImNum) - 1) ' byte array
            ' Transfer tempXORARR() to XORARR()
            For iy = 0 To UBound(tempXORARR(), 2)
            For ix = 0 To UBound(tempXORARR(), 1)
               XORARR(ix, iy) = tempXORARR(ix, iy)
            Next ix
            Next iy
            
            IcoCurBMIH(ImNum).ibpp = 8
            
         Case Else      ' No palette 24bpp  Make RGB ??
            ReDim XORARR(0 To XORW(ImNum) - 1, 0 To ImageHeight(ImNum) - 1)  ' byte array
            ' Scan picDATAREG(0 to 3, ImageWidth(ImNum), ImageHeight(ImNum)) & get RGB
            For iy = ImageHeight(ImNum) - 1 To 0 Step -1
            For ix = 0 To ImageWidth(ImNum) - 1
               Cul = RGB(picDATAREG(2, ix, iy), picDATAREG(1, ix, iy), picDATAREG(0, ix, iy))
               If Cul <> TColorBGR Then   ' So where color was TColorBGR now left = 0 in XORARR().
                  XORARR(3 * ix, iy) = picDATAREG(0, ix, iy)
                  XORARR(3 * ix + 1, iy) = picDATAREG(1, ix, iy)
                  XORARR(3 * ix + 2, iy) = picDATAREG(2, ix, iy)
               Else
                  k1 = Cul
               End If
            Next ix
            Next iy
            IcoCurBMIH(ImNum).ibpp = 24
         End Select
   
      Else ' IMBPP(ImageNum)=32
            
         ReDim XORARR(0 To XORW(ImNum) - 1, 0 To ImageHeight(ImNum) - 1)  ' byte array
         
         For iy = ImageHeight(ImNum) - 1 To 0 Step -1
         For ix = 0 To ImageWidth(ImNum) - 1
               Select Case ImNum
               Case 0
                  CorgB = DATACUL0(0, ix, iy)
                  CorgG = DATACUL0(1, ix, iy)
                  CorgR = DATACUL0(2, ix, iy)
                  ABYTE = DATACUL0(3, ix, iy)
               Case 1
                  CorgB = DATACUL1(0, ix, iy)
                  CorgG = DATACUL1(1, ix, iy)
                  CorgR = DATACUL1(2, ix, iy)
                  ABYTE = DATACUL1(3, ix, iy)
               Case 2
                  CorgB = DATACUL2(0, ix, iy)
                  CorgG = DATACUL2(1, ix, iy)
                  CorgR = DATACUL2(2, ix, iy)
                  ABYTE = DATACUL2(3, ix, iy)
               End Select
               XORARR(4 * ix, iy) = CorgB
               XORARR(4 * ix + 1, iy) = CorgG
               XORARR(4 * ix + 2, iy) = CorgR
               XORARR(4 * ix + 3, iy) = ABYTE
         Next ix
         Next iy
         IcoCurBMIH(ImNum).ibpp = 32
   
      End If
   
   Else  ' No optimization, 8bpp Done!
      ' Set any TColorBGR to 000 in gifPAL
      For k1 = 0 To 255
         If gifPAL(k1) = TColorRGB Then
            gifPAL(k1) = 0
         End If
      Next k1
      ReDim XORARR(0 To XORW(ImNum) - 1, 0 To ImageHeight(ImNum) - 1) ' byte array
      ' Transfer tempXORARR() to XORARR()
      For iy = 0 To UBound(tempXORARR(), 2)
      For ix = 0 To UBound(tempXORARR(), 1)
         XORARR(ix, iy) = tempXORARR(ix, iy)
      Next ix
      Next iy
      
      IcoCurBMIH(ImNum).ibpp = 8
   End If
   
   ImageNum = svImageNum
End Sub


'#### Image Clear, Reload, Undo/Redo & UndoAll ####

Private Sub cmdClear_Click(Index As Integer)
' ImageNum = Clng(Index)
Dim resp As Long
   
   If aSelect Then TURNSELECTOFF
   
   If From_mnuNew Then
      resp = vbYes
   Else
      resp = MsgBox("Clear image" & Str$(Index + 1) & " to Transparent color" & vbCrLf & _
             "and clear any Undos/Redos." & vbCrLf & "Sure?", vbQuestion + vbYesNo, "Clear")
   End If
   
   If resp = vbYes Then
      Response$ = "1"
      picPANEL.BackColor = TColorBGR
      picPANEL.Picture = LoadPicture
      picPANEL.Picture = picPANEL.Image
      
      picSmall(Index).BackColor = TColorBGR
      picSmall(Index).Picture = LoadPicture
      picSmall(Index).Picture = picSmall(Index).Image
      
      cmdUNDO(Index).Enabled = False
      cmdRedo(Index).Enabled = False
      DrawGrid
      NameSpec$(Index) = ""
   
      ImageNum = CLng(Index)
      
      ImBPP(ImageNum) = 1
      ' Color count
      ShowTheColorCount ImageNum
      
      KILLSAVS ImageNum
   
      LabCurGen(ImageNum) = CurrentGen(ImageNum)
      LabMaxGen(ImageNum) = MaxGen(ImageNum)
      
      cmdSelect_MouseUp Index, 1, 0, 0, 0
   
      cmdUNDO(Index).Enabled = False
      cmdRedo(Index).Enabled = False
      cmdClear(Index).Enabled = False
      cmdReload(Index).Enabled = False
      mnuAlphaEdit(Index).Enabled = False
      mnuAlphaEdit(Index + 4).Enabled = False
      
      Select Case ImageNum
      Case 0:
         ReDim DATACUL0(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)  '?
         ImBPP(ImageNum) = 0
      Case 1:
         ReDim DATACUL1(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)  '?
         ImBPP(ImageNum) = 0
      Case 2:
         ReDim DATACUL2(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)  '?
         ImBPP(ImageNum) = 0
      End Select
   
   Else
      Response$ = ""
   End If
   picPANEL.SetFocus
End Sub

Private Sub cmdReload_Click(Index As Integer)
   mnuReloadImage_Click Index
End Sub

Private Sub mnuReloadImage_Click(Index As Integer)
Dim resp As Long
   
   If aSelect Then TURNSELECTOFF
   
   resp = MsgBox("Reload original opened image" & Str$(ImageNum + 1) & vbCrLf & _
          "and clear any Undos/Redos." & vbCrLf & "Sure?", vbQuestion + vbYesNo, "Reload")
   If resp = vbYes Then
      LabMAC = "Images"
      
      ProcessInFile
   
   End If
End Sub

Public Sub BackUp(ImNum As Long)
Dim fnum As Long
Dim k As Long
Dim OldName$, NewName$
   
   CurrentGen(ImNum) = CurrentGen(ImNum) + 1
   
   If CurrentGen(ImNum) > MaxGenAllowed(ImNum) Then
      ' n=0 -> n=MaxGenAllowed(ImNum)-1
      
      ' SAV(im)n.dat = SAV(im)n+1.dat
      ' Kill SAV(im)n+1.dat
      ' CurrentGen(ImNum) = CurrentGen(ImNum) - 1
      ' MaxGen(ImNum) = CurrentGen(ImNum)
      ' Kill first
      OldName$ = AppPathSpec$ & "SAV"
      OldName$ = OldName$ & Trim$(Str$(ImNum)) & Trim$(Str$(1)) & ".dat"
      If FileExists(OldName$) Then Kill OldName$
      ' Rename  2->1 3->2,,,  MaxGenAllowed(ImNum) -> MaxGenAllowed(ImNum)-1
      For k = 2 To MaxGenAllowed(ImNum)
         OldName$ = AppPathSpec$ & "SAV"
         OldName$ = OldName$ & Trim$(Str$(ImNum)) & Trim$(Str$(k)) & ".dat"
         NewName$ = AppPathSpec$ & "SAV"
         NewName$ = NewName$ & Trim$(Str$(ImNum)) & Trim$(Str$(k - 1)) & ".dat"
         Name OldName$ As NewName$
      Next k
      ' Kill last
      OldName$ = AppPathSpec$ & "SAV"
      OldName$ = OldName$ & Trim$(Str$(ImNum)) & Trim$(Str$(MaxGenAllowed(ImNum))) & ".dat"
      If FileExists(OldName$) Then Kill OldName$
      CurrentGen(ImNum) = CurrentGen(ImNum) - 1
      MaxGen(ImNum) = CurrentGen(ImNum)
    End If
   
   If CurrentGen(ImNum) > MaxGen(ImNum) Then
      MaxGen(ImNum) = CurrentGen(ImNum)
   End If
   
   LabCurGen(ImageNum) = CurrentGen(ImageNum)
   LabMaxGen(ImageNum) = MaxGen(ImageNum)

   SAVDAT(0) = ImageWidth(ImNum)
   SAVDAT(1) = ImageHeight(ImNum)
   SAVDAT(2) = ImBPP(ImNum)
   
   GetTheBitsLong picSmall(ImNum) ' Get picSmallDATA()
   
   BUSpec$ = AppPathSpec$ & "SAV"
   BUSpec$ = BUSpec$ & Trim$(Str$(ImNum)) & Trim$(Str$(CurrentGen(ImNum))) & ".dat"
   fnum = FreeFile
   If FileExists(BUSpec$) Then Kill BUSpec$
   
   Open BUSpec$ For Binary As fnum
   Put #fnum, , SAVDAT()
   Put #fnum, , picSmallDATA()
   If ImBPP(ImNum) = 32 Then
      Select Case ImNum
      Case 0
         Put #fnum, , DATACUL0()
      Case 1
         Put #fnum, , DATACUL1()
      Case 2
         Put #fnum, , DATACUL2()
      End Select
   End If
   Close #fnum
   
   If MaxGen(ImNum) > 0 Then
      cmdUNDO(ImNum).Enabled = True
   End If
   If CurrentGen(ImNum) = MaxGen(ImNum) Then
      cmdRedo(ImNum).Enabled = False
   Else
      cmdRedo(ImNum).Enabled = True
   End If
   cmdClear(ImNum).Enabled = True
   
   If Not Set_LOHI_XY() Then
     TURNSELECTOFF
     Exit Sub
   End If
   
End Sub

Private Sub cmdUNDO_Click(Index As Integer)
' Index = ImageNum

Dim BMIH As BITMAPINFOHEADER
Dim fnum As Long

   ImageNum = Index
   
   If aSelect Then
      CancelSelection
      picPANEL.MouseIcon = LoadResPicture("PEN", vbResCursor)
      picPANEL.MousePointer = vbCustom
      Tools = [Dots]
      optTools(0).Value = True
      ToolsCaption
   End If
   
   'BackUp does CurrentGen(ImageNum) = CurrentGen(ImageNum) + 1
   '            & Sets MagGen(ImageNum)
   If CurrentGen(ImageNum) - 1 < 1 Then CurrentGen(ImageNum) = 2
   CurrentGen(ImageNum) = CurrentGen(ImageNum) - 1
   
   cmdRedo(ImageNum).Enabled = True
   
   If CurrentGen(ImageNum) = 1 Then
      cmdUNDO(ImageNum).Enabled = False
   End If
   If CurrentGen(ImageNum) = MaxGen(ImageNum) Then
      cmdRedo(ImageNum).Enabled = False
   End If
   
   LabCurGen(ImageNum) = CurrentGen(ImageNum)
   LabMaxGen(ImageNum) = MaxGen(ImageNum)
   
   BUSpec$ = AppPathSpec$ & "SAV"
   BUSpec$ = BUSpec$ & Trim$(Str$(ImageNum)) & Trim$(Str$(CurrentGen(ImageNum))) & ".dat"
   fnum = FreeFile
   
   Open BUSpec$ For Binary As fnum
   Get #fnum, , SAVDAT()
   
   ImageWidth(ImageNum) = SAVDAT(0)
   ImageHeight(ImageNum) = SAVDAT(1)
   ImBPP(ImageNum) = SAVDAT(2)
   LabSpec = "(Bpp =" & Str$(ImBPP(ImageNum)) & ")"
   SetAlphaEditMenu

   ReDim picSmallDATA(0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   
   Get #fnum, , picSmallDATA()
   If ImBPP(ImageNum) = 32 Then
      Select Case ImageNum
      Case 0
         ReDim DATACUL0(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         Get #fnum, , DATACUL0()
      Case 1
         ReDim DATACUL1(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         Get #fnum, , DATACUL1()
      Case 2
         ReDim DATACUL2(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         Get #fnum, , DATACUL2()
      End Select
   End If
   Close #fnum
   
   picSmall(ImageNum).Height = ImageHeight(ImageNum)
   picSmall(ImageNum).Width = ImageWidth(ImageNum)
   
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = ImageWidth(ImageNum)
      .biHeight = ImageHeight(ImageNum)
      .biBitCount = 32
   End With
   
   If SetDIBits(picSmall(ImageNum).hdc, picSmall(ImageNum).Image, _
      0, ImageHeight(ImageNum), picSmallDATA(0, 0), BMIH, 0) = 0 Then
      MsgBox "SetDIBits Error", vbCritical, "optTools"
   End If
   picSmall(ImageNum).Picture = picSmall(ImageNum).Image
   
   scrWidth.Value = ImageWidth(ImageNum)
   scrHeight.Value = 65 - ImageHeight(ImageNum)
   
   If Not Set_LOHI_XY Then Exit Sub

   DrawGrid
   
   ' Color count
   ShowTheColorCount ImageNum
   picPANEL.SetFocus
End Sub

Private Sub cmdUndoALL_Click(Index As Integer)
Dim resp As Long
   
   ImageNum = Index
   
   If aSelect Then TURNSELECTOFF
   
   resp = MsgBox("Kill all Undos/Redos for image" & Str$(Index + 1) & "  " & _
   vbCrLf & " & fix current image.", vbQuestion + vbYesNo, "Undo all or Stretch")
   If resp = vbYes Then
      Response$ = "1"
      KILLSAVS CLng(Index)
      cmdUNDO(Index).Enabled = False
      cmdRedo(Index).Enabled = False
      BackUp CLng(Index)
      cmdRedo(ImageNum).Enabled = False
      cmdUNDO(ImageNum).Enabled = False
   Else
      Response$ = ""
   End If
   LabCurGen(ImageNum) = CurrentGen(ImageNum)
   LabMaxGen(ImageNum) = MaxGen(ImageNum)
   picPANEL.SetFocus
End Sub

Private Sub cmdRedo_Click(Index As Integer)
Dim BMIH As BITMAPINFOHEADER
Dim fnum As Long
   
   ImageNum = Index
   
   If aSelect Then
      CancelSelection
      picPANEL.MouseIcon = LoadResPicture("PEN", vbResCursor)
      picPANEL.MousePointer = vbCustom
      Tools = [Dots]
      optTools(0).Value = True
      ToolsCaption
   End If
   
   If CurrentGen(ImageNum) + 1 > MaxGen(ImageNum) Then
      Exit Sub
   End If
   CurrentGen(ImageNum) = CurrentGen(ImageNum) + 1
   
   cmdUNDO(ImageNum).Enabled = True
   
   If CurrentGen(ImageNum) = MaxGen(ImageNum) Then
      cmdRedo(ImageNum).Enabled = False
   End If
   
   LabCurGen(ImageNum) = CurrentGen(ImageNum)
   LabMaxGen(ImageNum) = MaxGen(ImageNum)
   
   'BackUp does CurrentGen(ImageNum) = CurrentGen(ImageNum) + 1
   '            & Sets MagGen(ImageNum)
   BUSpec$ = AppPathSpec$ & "SAV"
   BUSpec$ = BUSpec$ & Trim$(Str$(ImageNum)) & Trim$(Str$(CurrentGen(ImageNum))) & ".dat"
   fnum = FreeFile
   
   Open BUSpec$ For Binary As fnum
   Get #fnum, , SAVDAT()
   ImageWidth(ImageNum) = SAVDAT(0)
   ImageHeight(ImageNum) = SAVDAT(1)
   ImBPP(ImageNum) = SAVDAT(2)
   LabSpec = "(Bpp =" & Str$(ImBPP(ImageNum)) & ")"
   SetAlphaEditMenu
   ReDim picSmallDATA(0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   Get #fnum, , picSmallDATA()
   If ImBPP(ImageNum) = 32 Then
      Select Case ImageNum
      Case 0
         ReDim DATACUL0(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         Get #fnum, , DATACUL0()
      Case 1
         ReDim DATACUL1(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         Get #fnum, , DATACUL1()
      Case 2
         ReDim DATACUL2(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         Get #fnum, , DATACUL2()
      End Select
   End If
   Close #fnum
  
   picSmall(ImageNum).Height = ImageHeight(ImageNum)
   picSmall(ImageNum).Width = ImageWidth(ImageNum)
   
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = ImageWidth(ImageNum)
      .biHeight = ImageHeight(ImageNum)
      .biBitCount = 32
   End With
   
   If SetDIBits(picSmall(ImageNum).hdc, picSmall(ImageNum).Image, _
      0, ImageHeight(ImageNum), picSmallDATA(0, 0), BMIH, 0) = 0 Then
      MsgBox "SetDIBits Error", vbCritical, "optTools"
   End If
   picSmall(ImageNum).Picture = picSmall(ImageNum).Image
   
   scrWidth.Value = ImageWidth(ImageNum)
   scrHeight.Value = 65 - ImageHeight(ImageNum)
   
   If Not Set_LOHI_XY Then Exit Sub

   ' Color count
   ShowTheColorCount ImageNum
   DrawGrid

   picPANEL.SetFocus
End Sub

'#### Image Stretch/Shrink ####

Private Sub mnuStretchShrink_Click()
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
   If aSelect Then TURNSELECTOFF
End Sub

Private Sub mnuStretchImage_Click(Index As Integer)
   
   If aSelect Then TURNSELECTOFF
   
   Select Case Index    ' Menu         Src  Dest
   Case 0: Adapter 0, 1 ' 1 to 2  image 0 to 1
   Case 1: Adapter 0, 2 ' 1 to 3  image 0 to 2
   Case 2: Adapter 1, 0 ' 2 to 1  image 1 to 0
   Case 3: Adapter 2, 0 ' 3 to 1  image 2 to 0
   Case 4: Adapter 1, 2 ' 2 to 3  image 1 to 2
   Case 5: Adapter 2, 1 ' 3 to 2  image 2 to 1
   Case 6   ' HALFTONE or COLORONCOLOR  ToneNumber 3 0r 4
   End Select
End Sub

Private Sub Adapter(Src As Long, Dest As Long)
' Stretch/Shrink Src to Dest
' COLORONCOLOR = 3
' HALFTONE = 4
Dim BMIH As BITMAPINFOHEADER
Dim svImageNum As Long
Dim k1 As Long, k2 As Long
Dim ix As Long, iy As Long
Dim RR As Byte, RG As Byte, RB As Byte

   svImageNum = ImageNum
   ImageNum = Dest
   
   cmdUndoALL_Click CInt(Dest)  ' Puts out a MsgBox
   
   ImageNum = Src
   If Len(Response$) = 0 Then
      ImageNum = svImageNum
      Exit Sub
   End If
   

' Test Shrink image 3 into 1 where
' image 3 is a 48x48 ico and image 1 size is 16x16
'------------------------------------------------------------------------------------
'   ' POOR  FOR SHRINKING, WHY?
'   SetStretchBltMode picSmall(Dest), 3 ' 3 or 4 makes no diff
'   StretchBlt picSmall(Dest).hDC, 0, 0, picSmall(Dest).Width, picSmall(Dest).Height, _
'      picSmall(Src).hDC, 0, 0, picSmall(Src).Width, picSmall(Src).Height, vbSrcCopy
'      picSmall(Dest).Picture = picSmall(Dest).Image
'
'   GoTo XX
'------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------
  ' Get picSmallDATA(0 To picSmall(Src).Width-1, 0 To picSmall(Src).Height-1)
  ' from picSmall(Src)
   GetTheBitsLong picSmall(Src)
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = picSmall(Src).Width
      .biHeight = picSmall(Src).Height
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With
   
'   ' HALFTONE = 4
'   ' COLORONCOLOR = 3
'   ' Could do :-
'   ' since if Src Width > Dest width Dest height also likely to be smaller
'   If picSmall(Src).Width > picSmall(Dest).Width Then
'      SetStretchBltMode picSmall(Dest).hDC, 4
'   Else
'      SetStretchBltMode picSmall(Dest).hDC, 3
'   End If
'   ' 3 more accurate for simple images
'   ' when stretching but leaves out more
'   ' when shrinking
'   ' 4 better image in some ways
'   ' but several stretch/shrinks
'   ' leads to more deterioration in image
   If ImBPP(Src) <> 32 Then
      
      If ToneNumber = 3 Then ' Halftone
         SetStretchBltMode picSmall(Dest).hdc, 4
         SetStretchBltMode picORG(Dest).hdc, 4
      Else  ' Coloroncolor
         SetStretchBltMode picSmall(Dest).hdc, 3
         SetStretchBltMode picORG(Dest).hdc, 3
      End If
   
   Else  ' 32bpp
   
      ' HalfTone matches Resize3DByteArray better
      SetStretchBltMode picSmall(Dest).hdc, 4
      SetStretchBltMode picORG(Dest).hdc, 4
   
   End If
   
   StretchDIBits picSmall(Dest).hdc, 0, 0, picSmall(Dest).Width, picSmall(Dest).Height, _
      0, 0, picSmall(Src).Width, picSmall(Src).Height, picSmallDATA(0, 0), BMIH, 0, vbSrcCopy
   
   picSmall(Dest).Picture = picSmall(Dest).Image
   
   
   picORG(Dest).Width = picSmall(Dest).Width
   picORG(Dest).Height = picSmall(Dest).Height
   picORG(Dest).Picture = LoadPicture
   picORG(Dest).BackColor = TColorBGR

   StretchDIBits picORG(Dest).hdc, 0, 0, picORG(Dest).Width, picORG(Dest).Height, _
   0, 0, picSmall(Src).Width, picSmall(Src).Height, picSmallDATA(0, 0), BMIH, 0, vbSrcCopy
   
   picORG(Dest).Picture = picORG(Dest).Image

'------------------------------------------------------------------------------------
   ImageNum = Dest

   Select Case Src
   Case 0
      If ImBPP(0) = 32 Then
         Select Case Dest
         Case 1   ' 0 to 1
            ImBPP(1) = 32
            k1 = 0: k2 = 1
            ReDim DATACUL1(0 To 3, ImageWidth(k2) - 1, ImageHeight(k2) - 1)
            Resize3DByteArray DATACUL1(), DATACUL0() ' Dest, Src

         Case 2   ' 0 to 2
            ImBPP(2) = 32
            k1 = 0: k2 = 2
            ReDim DATACUL2(0 To 3, ImageWidth(k2) - 1, ImageHeight(k2) - 1)
            Resize3DByteArray DATACUL2(), DATACUL0() ' Dest, Src
         End Select
      Else  ' ImBPP(Src) <> 32bpp so if ImBPP(dest) = 32bpp need to set ImBPP(Dest)=ImBPP(Src)
         ImBPP(Dest) = ImBPP(Src)
      End If
   
   Case 1
      If ImBPP(1) = 32 Then
         Select Case Dest
         Case 0   ' 1 to 0
            ImBPP(0) = 32
            k1 = 1: k2 = 0
            ReDim DATACUL0(0 To 3, ImageWidth(k2) - 1, ImageHeight(k2) - 1)
            Resize3DByteArray DATACUL0(), DATACUL1() ' Dest, Src
         Case 2   ' 1 to 2
            ImBPP(2) = 32
            k1 = 1: k2 = 2
            ReDim DATACUL2(0 To 3, ImageWidth(k2) - 1, ImageHeight(k2) - 1)
            Resize3DByteArray DATACUL2(), DATACUL1() ' Dest, Src
         End Select
      Else  ' ImBPP(Src) <> 32bpp so if ImBPP(dest) = 32bpp need to set ImBPP(Dest)=ImBPP(Src)
         ImBPP(Dest) = ImBPP(Src)
      End If
   
   Case 2
      If ImBPP(2) = 32 Then
         Select Case Dest
         Case 0   ' 2 to 0
            ImBPP(0) = 32
            k1 = 2: k2 = 0
            ReDim DATACUL0(0 To 3, ImageWidth(k2) - 1, ImageHeight(k2) - 1)
            Resize3DByteArray DATACUL0(), DATACUL2() ' Dest, Src
         Case 1   ' 2 to 1
            ImBPP(1) = 32
            k1 = 2: k2 = 1
            ReDim DATACUL1(0 To 3, ImageWidth(k2) - 1, ImageHeight(k2) - 1)
            Resize3DByteArray DATACUL1(), DATACUL2() ' Dest, Src
         End Select
      Else  ' ImBPP(Src) <> 32bpp so if ImBPP(dest) = 32bpp need to set ImBPP(Dest)=ImBPP(Src)
         ImBPP(Dest) = ImBPP(Src)
      End If
   End Select
   
   ' cmdSelect_MouseUp CInt(Dest), 1, 0, 0, 0  ' Sets ImageNum = Dest & Does DrawGrid
   '  DrawGrid  ' Also does ' TransferSmallToLarge picSmall(ImageNum), picPANEL
                         ' & sets picSmallDATA(0,0)
   
   ' For Halftone:-
   
   ' 0  Blur ALL
   ' 1  Blur image only   NU
   ' 2  Blur Alpha only
   
   ' For non-32bpp
   '    0  Blur ALL
   '    Blur makes picDATAREG() from picSmall(ImageNum)
   '    & DisplayEffects CInt(ImageNum)  does
   '      picDATAREG(0, 0, 0) to picSmall(ImageNum)
   
   ' For 32bpp
   '    2  Blur Alpha only
   '    DATACUL#() is used to fill picDATAREG()
   ImageNum = Dest
   
   If ImBPP(ImageNum) = 32 Then
      If ToneNumber = 3 Then ' Halftone
         For k1 = 1 To 6
            Blur picSmall(ImageNum), 2
            DisplayEffects CInt(ImageNum)
         Next k1
      Else   ' Coloroncolor
         For k1 = 1 To 1
            Blur picSmall(ImageNum), 2
            DisplayEffects CInt(ImageNum)
         Next k1
      End If
      Reconcile picSmall(ImageNum), ImageNum  ' Ensure Alpha and Image aligned

      'Reconcile ImageNum
   Else  ' <> 32 bpp
      If ToneNumber = 3 Then ' Halftone
         For k1 = 1 To 1
            Blur picSmall(ImageNum), 0
            DisplayEffects CInt(ImageNum)
         Next k1
      End If
   End If

   SetAlphaEditMenu
   
   cmdSelect_MouseUp CInt(ImageNum), 1, 0, 0, 0
   
   ShowTheColorCount ImageNum
   LabSpec = "(bpp =" & Str$(ImBPP(Dest)) & ")"
   BackUp ImageNum
   NameSpec$(Dest) = ""
   
End Sub

'#### Image swapper ####

Private Sub mnuSwap_Click()
   If aSelect Then TURNSELECTOFF
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
End Sub

Private Sub mnuSwapImages_Click(Index As Integer)
Dim a$
Dim k1 As Integer
Dim k2 As Integer
Dim k As Long
Dim CurrImageNum As Long
Dim resp As Long

   CurrImageNum = ImageNum
   Select Case Index
   Case 0   ' Swap 1 & 2  Images 0 & 1
      k1 = 0
      k2 = 1
   Case 1   ' Swap 1 & 3  Images 0 & 2
      k1 = 0
      k2 = 2
   Case 2   ' Swap 2 & 3  Images 1 & 2
      k1 = 1
      k2 = 2
   End Select
      
   resp = MsgBox("Kill all Undos/Redos for images" & Str$(k1 + 1) & " &" & Str$(k2 + 1) & "  " _
                 & vbCrLf & " & fix them as current images.", _
                  vbQuestion + vbYesNo, "Swapping")
   
   
   If resp = vbYes Then
      KILLSAVS CLng(k1)
      cmdUNDO(k1).Enabled = False
      cmdRedo(k1).Enabled = False
      KILLSAVS CLng(k2)
      cmdUNDO(k2).Enabled = False
      cmdRedo(k2).Enabled = False
   Else
      Exit Sub
   End If
   
   
      ' Displayed images
      
      ' Image  k1 to picSmallBU
      ' Image  k2 To k1
      ' picSmallBU to Image k1
      
      '--------------------------------------------------------
      ' Image  k1 to picSmallBU
      picSmallBU.Picture = LoadPicture
      With picSmallBU
         .Width = picSmall(k1).Width
         .Height = picSmall(k1).Height
      End With
      picSmallBU.Picture = picSmallBU.Image
      BitBlt picSmallBU.hdc, 0, 0, picSmallBU.Width, picSmallBU.Height, _
             picSmall(k1).hdc, 0, 0, vbSrcCopy
      picSmallBU.Picture = picSmallBU.Image
      
      ' Image  k2 To k1
      With picSmall(k1)
         .Width = picSmall(k2).Width
         .Height = picSmall(k2).Height
      End With
      picSmall(k1).Picture = picSmall(k1).Image
      BitBlt picSmall(k1).hdc, 0, 0, picSmall(k1).Width, picSmall(k1).Height, _
             picSmall(k2).hdc, 0, 0, vbSrcCopy
      picSmall(k1).Picture = picSmall(k1).Image
      
      ' picSmallBU to Image k2
      picSmall(k2).Picture = LoadPicture
      With picSmall(k2)
         .Width = picSmallBU.Width
         .Height = picSmallBU.Height
      End With
      picSmall(k2).Picture = picSmall(k2).Image
      BitBlt picSmall(k2).hdc, 0, 0, picSmall(k2).Width, picSmall(k2).Height, _
             picSmallBU.hdc, 0, 0, vbSrcCopy
      picSmall(k2).Picture = picSmall(k2).Image
      
      '--------------------------------------------------------
      
      ' picORG images
      
      ' Image  prevk1 to picSmallBU
      ' Image  prevk2 To prevk1
      ' picSmallBU to Image prevk1
      
      ' Image  prevk1 to picSmallBU
      picSmallBU.Picture = LoadPicture
      With picSmallBU
         .Width = picORG(k1).Width
         .Height = picORG(k1).Height
      End With
      picSmallBU.Picture = picSmallBU.Image
      BitBlt picSmallBU.hdc, 0, 0, picSmallBU.Width, picSmallBU.Height, _
             picORG(k1).hdc, 0, 0, vbSrcCopy
      picSmallBU.Picture = picSmallBU.Image
      
      ' Image  prevk2 To prevk1
      With picORG(k1)
         .Width = picORG(k2).Width
         .Height = picORG(k2).Height
      End With
      picORG(k1).Picture = picORG(k1).Image
      BitBlt picORG(k1).hdc, 0, 0, picORG(k1).Width, picORG(k1).Height, _
             picORG(k2).hdc, 0, 0, vbSrcCopy
      picORG(k1).Picture = picORG(k1).Image
      
      ' picSmallBU to Image prevk2
      picORG(k2).Picture = LoadPicture
      With picORG(k2)
         .Width = picSmallBU.Width
         .Height = picSmallBU.Height
      End With
      picORG(k2).Picture = picORG(k2).Image
      BitBlt picORG(k2).hdc, 0, 0, picORG(k2).Width, picORG(k2).Height, _
             picSmallBU.hdc, 0, 0, vbSrcCopy
      picORG(k2).Picture = picORG(k2).Image
      '--------------------------------------------------------
      ' Swap 32bpp images
      If ImBPP(k1) = 32 Or ImBPP(k2) = 32 Then
         If k1 = 0 And k2 = 1 Then
            SwapPictures k1, k2, DATACUL0(), DATACUL1()
            k = ImBPP(1)
            ImBPP(1) = ImBPP(0)
            ImBPP(0) = k
         ElseIf k1 = 0 And k2 = 2 Then
            SwapPictures k1, k2, DATACUL0(), DATACUL2()
            k = ImBPP(2)
            ImBPP(2) = ImBPP(0)
            ImBPP(0) = k
         
         ElseIf k1 = 1 And k2 = 2 Then
            SwapPictures k1, k2, DATACUL1(), DATACUL2()
            k = ImBPP(2)
            ImBPP(2) = ImBPP(1)
            ImBPP(1) = k
         End If
      End If
      '--------------------------------------------------------
      
      ' Swap NameSpec$
      a$ = NameSpec$(k1)
      NameSpec$(k1) = NameSpec$(k2)
      NameSpec$(k2) = a$
      '--------------------------------------------------------
      ' Swap FileSpec$
      a$ = ImFileSpec$(k1)
      ImFileSpec$(k1) = ImFileSpec$(k2)
      ImFileSpec$(k2) = a$
      '--------------------------------------------------------
      ' Swap widths & heights
      k = ImageWidth(k1)
      ImageWidth(k1) = ImageWidth(k2)
      ImageWidth(k2) = k
      
      k = ImageHeight(k1)
      ImageHeight(k1) = ImageHeight(k2)
      ImageHeight(k2) = k
      '--------------------------------------------------------
      
      ' Set values
      ImageNum = k1
      scrWidth.Value = ImageWidth(k1)
      scrHeight.Value = 65 - ImageHeight(k1)
      ShowTheColorCount ImageNum

      ImageNum = k2
      scrWidth.Value = ImageWidth(k2)
      scrHeight.Value = 65 - ImageHeight(k2)
      ShowTheColorCount ImageNum
      '--------------------------------------------------------
      
      ' Transfer to picPANEL
      If CurrImageNum = k1 Then
         cmdSelect_MouseUp k1, 1, 0, 0, 0
         cmdSelect_MouseUp k2, 1, 0, 0, 0
      Else
         cmdSelect_MouseUp k2, 1, 0, 0, 0
         cmdSelect_MouseUp k1, 1, 0, 0, 0
      End If
      
      BackUp CLng(k1)
      BackUp CLng(k2)
End Sub


'#### Cursor test ####

Private Sub mnuCursor_Click()
   If aSelect Then TURNSELECTOFF
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
   
End Sub

Private Sub mnuTestCursor_Click(Index As Integer)
Dim r As RECT
Dim Title$, Filt$, InDir$
Dim fnum As Long
Dim Ext$
Dim CStream() As Byte
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

   If aSelect Then TURNSELECTOFF
   
   On Error GoTo CURSORERROR
   
   Select Case Index '= 0 Then
   Case 0   ' Test a cursor
      Title$ = "Open CUR (or ICO) file"
      Filt$ = "Pics cur,ico|*.cur;*.ico"
      CurSpec$ = ""
      InDir$ = CPath$ 'AppPathSpec$
      Set CommonDialog1 = New cOSDialog
   
      CommonDialog1.ShowOpen CurSpec$, Title$, Filt$, InDir$, "", Me.hwnd
      Set CommonDialog1 = Nothing
      
      ' Avoid click through
      GetWindowRect Form1.picSmallFrame.hwnd, r
      SetCursorPos r.Left + 130, r.Top + 90
      
      If Len(CurSpec$) = 0 Then
         TURNOFFCURSOR
         Exit Sub
      Else
         
         Ext$ = UCase$(Right$(FileSpec$, 3))

         If aTestCursor Then TURNOFFCURSOR
         ' Tests cursor on form & all picboxes
         ' apart from picPANEL ?
         ShowNewCursor Form1, picPANEL, CurSpec$
         MousePointer = 99
         TestCursNum = TestCursNum + 1
         
         'Get HotX/Y
         fnum = FreeFile
         Open CurSpec$ For Binary As #fnum
         CStream = Space$(LOF(fnum))
         Get #fnum, , CStream
         Close #fnum
         If Ext$ = "ICO" Then
            HotX(ImageNum) = 0
            HotY(ImageNum) = 0
         Else
            HotX(ImageNum) = CStream(10)
            HotY(ImageNum) = CStream(12)
         End If
         CStream = Space$(1)
         
         'aTestCursor = False
         'scrHotX.Value = HotX(ImageNum)
         'scrHotY.Value = HotY(ImageNum)
         'aTestCursor = True
         
         LabHotXY(0) = "X=" & HotX(ImageNum)
         LabHotXY(1) = "Y=" & HotY(ImageNum)
         LabHotXY(0).Refresh
         LabHotXY(1).Refresh
      End If
   Case 1   ' Cancel cursor
      TURNOFFCURSOR
   End Select
   
   Exit Sub
'=========
CURSORERROR:
   TURNOFFCURSOR
   MsgBox "Error reading cursor file", vbCritical, "Small GFX"
End Sub

Private Sub TURNOFFCURSOR()
   RestoreOldCursor Form1, picPANEL
   MousePointer = 0
   Screen.MousePointer = vbDefault
   TestCursNum = 0
   aTestCursor = False
End Sub


Private Sub picSmallFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   picSmallFrame.MousePointer = vbDefault
End Sub

Private Sub scrHotX_Change()
   Call scrHotX_Scroll
End Sub

Private Sub scrHotX_Scroll()
   If Not aTestCursor Then Exit Sub
   HotX(ImageNum) = scrHotX.Value
   LabHotXY(0) = "X=" & HotX(ImageNum)
   LabHotXY(0).Refresh
   FlashHotSpot
End Sub

Private Sub scrHotY_Change()
   Call scrHotY_Scroll
End Sub

Private Sub scrHotY_Scroll()
   If Not aTestCursor Then Exit Sub
   HotY(ImageNum) = scrHotY.Value
   LabHotXY(1) = "Y=" & HotY(ImageNum)
   LabHotXY(1).Refresh
   FlashHotSpot
End Sub

Private Sub FlashHotSpot()
' NB Only for 32x32 sizes.
Dim Cul As Long
Dim T As Single
Dim Delay As Single
Dim k As Long
   
   On Error Resume Next
   
   Cul = picSmall(ImageNum).Point(HotX(ImageNum), HotY(ImageNum))
   If Cul >= 0 Then
      For k = 1 To 3
         ' Show hot spot
         picPANEL.DrawMode = vbXorPen
         picPANEL.Line (HotX(ImageNum) * GridMult + 2, HotY(ImageNum) * GridMult + 2)- _
         (HotX(ImageNum) * GridMult + GridMult - 2, HotY(ImageNum) * GridMult + GridMult - 2), RGB(255, 128, 128), BF
         picPANEL.Refresh
         Delay = 0.05  ' Sec
         T = Timer
         Do While Timer < T + Delay
         Loop
         ' Clear hot spot
         picPANEL.Line (HotX(ImageNum) * GridMult + 2, HotY(ImageNum) * GridMult + 2)- _
         (HotX(ImageNum) * GridMult + GridMult - 2, HotY(ImageNum) * GridMult + GridMult - 2), RGB(255, 128, 128), BF
         picPANEL.Refresh
      Next k
      picPANEL.DrawMode = vbCopyPen
   End If
   
   picPANEL.SetFocus
End Sub

Private Sub cmdFlash_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Long
   
   For k = 1 To 3
      FlashHotSpot
   Next k
End Sub

'#### Image selector ####

Private Sub cmdSelect_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Public ImageNum As Long
' ImageNum = CLng(Index)
Dim k As Long
   
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

   
   If aSelect Then TURNSELECTOFF
   
   picPANEL.Picture = LoadPicture
   
   shpCirc.Visible = False
   
   For k = 0 To 2
      LabNum(k).Enabled = False
      cmdUNDO(k).Enabled = False
      cmdRedo(k).Enabled = False
      cmdClear(k).Enabled = True
      cmdReload(k).Enabled = False
   Next k
   
   If mnuReloadImage(Index).Enabled Then
      cmdReload(Index).Enabled = True
   End If
   
   LabNum(Index).Enabled = True
   LabNum(Index).Refresh
   
   ImageNum = CLng(Index)  ' Used throughout
   
   If CurrentGen(ImageNum) > 1 Then
      cmdUNDO(ImageNum).Enabled = True
   End If
   
   If CurrentGen(ImageNum) >= 1 And CurrentGen(ImageNum) < MaxGen(ImageNum) Then
      cmdRedo(ImageNum).Enabled = True
   End If
   
   LabCurGen(ImageNum) = CurrentGen(ImageNum)
   LabMaxGen(ImageNum) = MaxGen(ImageNum)
   
   LabImageNumber = Index + 1
   LabImageNumber.Refresh
   LabImCurNum = Index + 1
   LabImCurNum.Refresh
   
   acmdSelect = True   ' Not used yet
   scrWidth.Value = ImageWidth(Index)
   scrHeight.Value = 65 - ImageHeight(Index)
   If Not Set_LOHI_XY Then Exit Sub
   acmdSelect = False   ' Not used yet
   
   ' When called these each do
   ' TransferSmallToLarge picSmall(ImageNum), picPANEL
   ' & DrawGrid but not always called ?
   DrawGrid
   
   If Len(NameSpec$(ImageNum)) <> 0 Then
      LabSpec = NameSpec$(ImageNum) & "  Path: " & ShortenString$(GetPath(ImFileSpec$(ImageNum)), 64) '74)
   Else
      ShowTheColorCount ImageNum
      LabSpec = "(bpp =" & Str$(ImBPP(ImageNum)) & ")"
   End If
   
   SetAlphaEditMenu
   picPANEL.SetFocus
End Sub

Private Sub SetAlphaEditMenu()
   If ImBPP(0) = 32 Then
      mnuAlphaEdit(0).Enabled = True
      mnuAlphaEdit(4).Enabled = True
   Else
      mnuAlphaEdit(0).Enabled = False
      mnuAlphaEdit(4).Enabled = False
   End If
   If ImBPP(1) = 32 Then
      mnuAlphaEdit(1).Enabled = True
      mnuAlphaEdit(5).Enabled = True
      Else
      mnuAlphaEdit(1).Enabled = False
      mnuAlphaEdit(5).Enabled = False
      End If
   If ImBPP(2) = 32 Then
      mnuAlphaEdit(2).Enabled = True
      mnuAlphaEdit(6).Enabled = True
   Else
      mnuAlphaEdit(2).Enabled = False
      mnuAlphaEdit(6).Enabled = False
   End If
End Sub

'#### Colors ####

Public Function GetColorCount(PIC As PictureBox) As Long
' pic = picSmall(ImageNum)
   GetTheBitsLong PIC
   GetColorCount = ColorCount(picSmallDATA())
End Function


Public Sub ShowTheColorCount(ImNum As Long)    ' Color count
Dim CulC As Long
   CulC = GetColorCount(picSmall(ImNum))
   LabNumColors(ImNum) = "Colors =" & Str$(CulC)
   If ImBPP(ImNum) <> 32 Then
      Select Case CulC
      Case Is <= 2: ImBPP(ImNum) = 1
      Case Is <= 16: ImBPP(ImNum) = 4
      Case Is <= 256: ImBPP(ImNum) = 8
      Case Else
         ImBPP(ImNum) = 24
      End Select
   End If
End Sub

Private Sub cmdPAL_Click(Index As Integer)
Dim CF As CFDialog
Dim TheColor As Long

   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

   If Index < 5 Then PALIndex = Index
   Select Case Index
   Case 0: QBColors picPAL
      QBColors picVisColor, 2
   Case 1: ShortBandedPAL picPAL
      ShortBandedPAL picVisColor, 2
   Case 2: LongBandedPAL picPAL
      LongBandedPAL picVisColor, 2
   Case 3: GreyPAL picPAL
      GreyPAL picVisColor, 2
   Case 4: CenteredPAL picPAL
      CenteredPAL picVisColor, 2
   Case 5: ' System colors
     Set CF = New CFDialog
     If CF.VBChooseColor(TheColor, , , , Me.hwnd) Then
         LColor = TheColor
         LabColor(0).BackColor = LColor
     End If
     Set CF = Nothing
   End Select
   If aShow Then picPANEL.SetFocus
End Sub

Private Sub cmdSwapLR_Click()
Dim Cul As Long
   Cul = LColor
   LColor = RColor
   RColor = Cul
   LabColor(0).BackColor = LColor
   LabColor(1).BackColor = RColor
   Cul = LabColor(0).ForeColor
   LabColor(0).ForeColor = LabColor(1).ForeColor
   LabColor(1).ForeColor = Cul
   CheckForEraseColor
   picPANEL.SetFocus
End Sub

Private Sub picPAL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SR As Byte, SG As Byte, SB As Byte
Dim Cul As Long
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
   
   Cul = picPAL.Point(X, Y)
   If Cul >= 0 Then
      If Cul = TColorBGR Then Cul = TColorBGR + 1
      LabDropper.BackColor = Cul
      LngToRGB Cul, SR, SG, SB
      LabRGB(0) = SR
      LabRGB(1) = SG
      LabRGB(2) = SB
   End If
End Sub

Private Sub picPAL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
   
   LabErase.Visible = False
   LineErase.Visible = False
   
   Cul = picPAL.Point(X, Y)
   
   ' Avoid accidental erasing
   'If Cul = TColorBGR Then Cul = TColorBGR + 1
   
   If Cul >= 0 Then
      If Button = vbLeftButton Then
         LColor = Cul
      ElseIf Button = vbRightButton Then
         RColor = Cul
      End If
      LabColor(0).BackColor = LColor
      LabColor(1).BackColor = RColor
   End If
   
   CheckForEraseColor
End Sub

Private Sub picErase_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LabErase.Visible = True
   LineErase.Visible = True
If Button = vbLeftButton Then
   LColor = TColorBGR
   LabColor(0).BackColor = LColor
Else
   RColor = TColorBGR
   LabColor(1).BackColor = RColor
End If

   CheckForEraseColor
End Sub

Private Sub CheckForEraseColor()
   If LColor = TColorBGR Then
      LabErase.Caption = "L Erases"
      LabErase.Visible = True
      LineErase.Visible = True
   End If
   
   If RColor = TColorBGR Then
      LabErase.Caption = "R Erases"
      LabErase.Visible = True
      LineErase.Visible = True
   End If
   
   If LColor = TColorBGR And RColor = TColorBGR Then
      LabErase.Caption = "L && R Erases"
      LabErase.Visible = True
      LineErase.Visible = True
   End If
End Sub


'#### SHOW WITH DIFF VISIBILITY ####

Private Sub picVisColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Long
Dim sw As Long, sh As Long
   
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
   
   If Button = vbLeftButton Then
   
      VisMouse = True
   
      ' Get test TColorBGR
      VisColor = picVisColor.Point(X, Y)
      If VisColor < 0 Then VisColor = RGB(128, 128, 128) '0
      ' Save picSmall() in picTemp()
      For k = 0 To 2
         sw = picSmall(k).Width
         sh = picSmall(k).Height
         With picTemp(k)
            .Picture = LoadPicture
            .Width = sw
            .Height = sh
            BitBlt .hdc, 0, 0, sw, sh, picSmall(k).hdc, 0, 0, vbSrcCopy
            .Picture = .Image
         End With
      Next k
         
      For k = 0 To 2
         ShowWithTColor k, picSmall(k), picSmall(k), 0
      Next k
      aGrid = 0
      DrawGrid
   
   
      aGrid = 2
      chkGridOnOff.Value = vbChecked
      aGrid = 1
      chkGridOnOff.Caption = "Grid On"
   End If
End Sub

Private Sub picVisColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Long
Dim sw As Long, sh As Long
   
   ShowRGBS picVisColor.Point(X, Y)
   
   If Button = vbLeftButton Then
   If VisMouse Then
   
      ' Restore images
      For k = 0 To 2
         sw = picSmall(k).Width
         sh = picSmall(k).Height
         With picSmall(k)
            .Picture = LoadPicture
            .Width = sw
            .Height = sh
            BitBlt .hdc, 0, 0, sw, sh, picTemp(k).hdc, 0, 0, vbSrcCopy
            .Picture = .Image
         End With
      Next k
      picSmallFrame.BackColor = &H808080 'TColorBGR
      
      
      ' Get test TColorBGR
      VisColor = picVisColor.Point(X, Y)
      If VisColor < 0 Then VisColor = RGB(194, 195, 197) '0
      For k = 0 To 2
         ShowWithTColor k, picSmall(k), picSmall(k), 0
      Next k
      aGrid = 0
      DrawGrid
      aGrid = 2
      chkGridOnOff.Value = vbChecked
      aGrid = 1
      chkGridOnOff.Caption = "Grid On"
   
   End If
   End If
End Sub

Private Sub picVisColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Long
Dim sw As Long, sh As Long
   
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
   
   If Button = vbLeftButton Then
      ' Restore picSmall() from picTemp()
      For k = 0 To 2
         sw = picSmall(k).Width
         sh = picSmall(k).Height
         With picSmall(k)
            .Picture = LoadPicture
            .Width = sw
            .Height = sh
            BitBlt .hdc, 0, 0, sw, sh, picTemp(k).hdc, 0, 0, vbSrcCopy
            .Picture = .Image
         End With
      Next k
      picSmallFrame.BackColor = &H808080 'TColorBGR
      VisColor = TColorBGR
      aGrid = 1
      chkGridOnOff.Caption = "Grid On"
      DrawGrid
      VisMouse = False
   End If
End Sub

' Show RGBs

Private Sub picGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ShowRGBS (picGrid.Point(X, Y))
End Sub


Private Sub LabColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   ShowRGBS (LabColor(Index).BackColor)
End Sub

Private Sub ShowRGBS(Cul As Long)
Dim SR As Byte, SG As Byte, SB As Byte
   If Cul >= 0 Then
      LabDropper.BackColor = Cul
      LngToRGB Cul, SR, SG, SB
      LabRGB(0) = SR
      LabRGB(1) = SG
      LabRGB(2) = SB
   End If
End Sub

'#### Menu Preferences ####

Private Sub mnuPrefs_Click()
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
End Sub

Private Sub mnuPreferences_Click(Index As Integer)
   Select Case Index
   Case 0, 1   ' OptimizeNumber = 0 Optimize, 1 Non-optimize
      mnuPreferences(0).Checked = False
      mnuPreferences(1).Checked = False
      OptimizeNumber = Index
      If Index = 0 Then
         mnuSaveBMPImage(4).Caption = "(Optimized saving)"
         mnuSaveICOImage(5).Caption = "(Optimized saving)"
         mnuSaveCURImage(3).Caption = "(Optimized saving)"
      Else
         mnuSaveBMPImage(4).Caption = "(Not optimized saving)"
         mnuSaveICOImage(5).Caption = "(Not optimized saving)"
         mnuSaveCURImage(3).Caption = "(Not optimized saving)"
      End If
      mnuPreferences(Index).Checked = True
   Case 2   ' Break
   Case 3, 4   ' ToneNumber = 3 HALFTONE, 4 COLORONCOLOR
      mnuPreferences(3).Checked = False
      mnuPreferences(4).Checked = False
      ToneNumber = Index
      If Index = 3 Then
         mnuStretchImage(6).Caption = "(HALFTONE)"
         mnuCLIP(7).Caption = "(HALFTONE)"
      Else
         mnuStretchImage(6).Caption = "(COLORONCOLOR)"
         mnuCLIP(7).Caption = "(COLORONCOLOR)"
      End If
      mnuPreferences(Index).Checked = True
   Case 5   ' Break
   Case 6, 7   ' AspectNumber = 6 Keep, 7 Ignore aspect ratio
      mnuPreferences(6).Checked = False
      mnuPreferences(7).Checked = False
      AspectNumber = Index
      If Index = 6 Then
         mnuCLIP(8).Caption = "(Keep aspect ratio)"
         mnuCaptureToImage(12).Caption = "(Keep aspect ratio)"
      Else
         mnuCLIP(8).Caption = "(Ignore aspect ratio)"
         mnuCaptureToImage(12).Caption = "(Ignore aspect ratio)"
      End If
      mnuPreferences(Index).Checked = True
   Case 8   ' Break
   Case 9, 10, 11
      ' MarkPixels =
      '  9 Mark transparent pixel-Square
      ' 10 Mark transparent pixel-Diagonal
      ' 11 Unmark transparent pixels
   
      mnuPreferences(9).Checked = False   ' Square    103
      mnuPreferences(10).Checked = False  ' Diagonal  102
      mnuPreferences(11).Checked = False  ' Unmarked  101
      MarkPixels = Index
      mnuPreferences(Index).Checked = True
      Select Case Index
      Case 9
         imTMarker.Picture = LoadResPicture(103, vbResBitmap)
      Case 10
         imTMarker.Picture = LoadResPicture(102, vbResBitmap)
      Case 11
         imTMarker.Picture = LoadResPicture(101, vbResBitmap)
      End Select
      mnuPreferences(Index).Checked = True
      DrawGrid
   Case 12   ' Break
   Case 13, 14 ' Extractorbackups  13 All, 14 Just last
      mnuPreferences(13).Checked = False
      mnuPreferences(14).Checked = False
      ExtractorBackups = Index
      mnuPreferences(Index).Checked = True
   Case 15  ' Break
   Case 16  ' Alpha On/Off for 32bpp Extracted images
      If mnuPreferences(16).Checked Then
         mnuPreferences(16).Checked = False
         aAlphaRestricted = False  ' Alpha ON
         mnuOpenIntoImage(4).Visible = False
      Else
         mnuPreferences(16).Checked = True
         aAlphaRestricted = True   ' Alpha OFF
         mnuOpenIntoImage(4).Visible = True
      End If
   Case 17   ' Break
   Case 18  ' Query original colors on saving 32bpp
      If mnuPreferences(18).Checked Then
            mnuPreferences(18).Checked = False
            aQueryOC = False  ' Query ON
      Else
         mnuPreferences(18).Checked = True
         aQueryOC = True   ' Query OFF
      End If
   
   End Select
End Sub

'#### Scrollers & Flippers ####

Private Sub cmdLRUD_Click(Index As Integer)
' Scrolls & Flips
Dim ix As Long, iy As Long
Dim BMIH As BITMAPINFOHEADER
   
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

   If aMoveSEL Then
      TURNSELECTOFF
      aSelect = True
      Exit Sub
   End If
   
   If Not Set_LOHI_XY Then Exit Sub
   
   If CurrentGen(ImageNum) = 0 Then Exit Sub
   
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = ImageWidth(ImageNum)
      .biHeight = ImageHeight(ImageNum)
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With

   ' Transfer pixel colors to picSmallDATA() from picSmall
   '===========================================================
   GetTheBitsLong picSmall(ImageNum)
   '===========================================================
    
   If aDIBError Then Exit Sub
   
   ReDim picDummy(0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   
   Dim TempL As Long
   ReDim TempA(0 To ImageWidth(ImageNum) - 1) As Long
 
   picDummy() = picSmallDATA()
   
   Select Case Index
   Case 0   ' Move picSmall Left
      If ImageWidth(ImageNum) = 1 Then Exit Sub
      For iy = LOY To HIY
         TempL = picSmallDATA(LOX, iy)   ' Left column
         CopyMemory picSmallDATA(LOX, iy), picSmallDATA(LOX + 1, iy), 4 * (HIX - LOX)
         picSmallDATA(HIX, iy) = TempL   ' To right column
         If ImBPP(ImageNum) = 32 Then ScrollLeft32 ImageNum, iy
      Next iy
   Case 1   ' Move picSmall Right
      If ImageWidth(ImageNum) = 1 Then Exit Sub
      For iy = LOY To HIY
         TempL = picSmallDATA(HIX, iy) ' Right column
         For ix = HIX - 1 To LOX Step -1
            picSmallDATA(ix + 1, iy) = picSmallDATA(ix, iy)
         Next ix
         picSmallDATA(LOX, iy) = TempL  ' To left column
         If ImBPP(ImageNum) = 32 Then ScrollRight32 ImageNum, iy
      Next iy
   
   Case 2   ' Move picSmall Up
      If ImageHeight(ImageNum) = 1 Then Exit Sub
      CopyMemory TempA(0), picSmallDATA(LOX, HIY), 4 * (HIX - LOX + 1) ' Top row
      For iy = HIY To LOY + 1 Step -1
         CopyMemory picSmallDATA(LOX, iy), picSmallDATA(LOX, iy - 1), 4 * (HIX - LOX + 1)
      Next iy
      CopyMemory picSmallDATA(LOX, LOY), TempA(0), 4 * (HIX - LOX + 1) ' To Bottom row
      If ImBPP(ImageNum) = 32 Then ScrollUp32 ImageNum
   Case 3   ' Move picSmall Down
      If ImageHeight(ImageNum) = 1 Then Exit Sub
      CopyMemory TempA(0), picSmallDATA(LOX, LOY), 4 * (HIX - LOX + 1) ' Bottom row
      For iy = LOY To HIY - 1
         CopyMemory picSmallDATA(LOX, iy), picSmallDATA(LOX, iy + 1), 4 * (HIX - LOX + 1)
      Next iy
      CopyMemory picSmallDATA(LOX, HIY), TempA(0), 4 * (HIX - LOX + 1)   ' To Top row
      If ImBPP(ImageNum) = 32 Then ScrollDown32 ImageNum
   
   Case 4   ' Horz Flip
      If ImageWidth(ImageNum) = 1 Then Exit Sub
      For iy = LOY To HIY
      For ix = LOX To HIX
         picSmallDATA(HIX + LOX - ix, iy) = picDummy(ix, iy)
      Next ix
      Next iy
      If ImBPP(ImageNum) = 32 Then HorzFlip ImageNum
   Case 5   ' Vert Flip
      If ImageHeight(ImageNum) = 1 Then Exit Sub
      For ix = LOX To HIX
      For iy = LOY To HIY
         picSmallDATA(ix, HIY + LOY - iy) = picDummy(ix, iy)
      Next iy
      Next ix
      If ImBPP(ImageNum) = 32 Then VertFlip ImageNum
   End Select
   
   ' Transfer
   SetDIBits picSmall(ImageNum).hdc, picSmall(ImageNum).Image, _
      0, ImageHeight(ImageNum), picSmallDATA(0, 0), BMIH, 0
   picSmall(ImageNum).Picture = picSmall(ImageNum).Image
   
   Erase Temp(), picSmallDATA(), picDummy()
   
   If aSelect Then
      aEffects = True
      Tools = [SelectON]
   Else
       BackUp ImageNum
   End If
   

   DrawGrid


   picORG(ImageNum).Width = ImageWidth(ImageNum)
   picORG(ImageNum).Height = ImageHeight(ImageNum)
   picORG(ImageNum).Picture = LoadPicture
   
   BitBlt picORG(ImageNum).hdc, 0, 0, ImageWidth(ImageNum), ImageHeight(ImageNum), _
         picSmall(ImageNum).hdc, 0, 0, vbSrcCopy


End Sub

'#### Scrollbars Width & Height ####

Private Sub cmdBackUp_Click()
   BackUp ImageNum
End Sub


Private Sub scrHeight_Change()
   If aSelect Then TURNSELECTOFF
   Call scrHeight_Scroll
End Sub

Private Sub scrHeight_Scroll()
Dim ix As Long, iy As Long
Dim iyy As Long
   If Not aScroll Then Exit Sub
   
   ImageHeight(ImageNum) = 65 - scrHeight.Value ' Used throughout
   LabHeight = "Height =" & Str$(ImageHeight(ImageNum))
   
   If picSmall(ImageNum).BackColor <> TColorBGR Then
      picSmall(ImageNum).BackColor = TColorBGR
   End If
   picSmall(ImageNum).Height = ImageHeight(ImageNum)
   picSmall(ImageNum).Picture = picSmall(ImageNum).Image
   picSmall(ImageNum).Width = ImageWidth(ImageNum)
   picSmall(ImageNum).Picture = picSmall(ImageNum).Image
   
   If ImBPP(ImageNum) = 32 Then
   
   ReDim DATACULSRC(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   
   iyy = ImageHeight(ImageNum) - 1
   
   Select Case ImageNum
   Case 0
      If ImBPP(0) = 32 Then
         For iy = UBound(DATACUL0(), 3) To 0 Step -1  'iy
            If iyy >= 0 Then
               For ix = 0 To UBound(DATACUL0(), 2)
                  If ix <= ImageWidth(ImageNum) - 1 Then
                     DATACULSRC(0, ix, iyy) = DATACUL0(0, ix, iy)
                     DATACULSRC(1, ix, iyy) = DATACUL0(1, ix, iy)
                     DATACULSRC(2, ix, iyy) = DATACUL0(2, ix, iy)
                     DATACULSRC(3, ix, iyy) = DATACUL0(3, ix, iy)
                  End If
               Next ix
            End If
            iyy = iyy - 1
         Next iy
         ReDim DATACUL0(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         FILL3D DATACUL0(), DATACULSRC()
      End If
   Case 1
      If ImBPP(1) = 32 Then
         For iy = UBound(DATACUL1, 3) To 0 Step -1
            If iyy >= 0 Then
               For ix = 0 To UBound(DATACUL1, 2)
                  If ix <= ImageWidth(ImageNum) - 1 Then
                     DATACULSRC(0, ix, iyy) = DATACUL1(0, ix, iy)
                     DATACULSRC(1, ix, iyy) = DATACUL1(1, ix, iy)
                     DATACULSRC(2, ix, iyy) = DATACUL1(2, ix, iy)
                     DATACULSRC(3, ix, iyy) = DATACUL1(3, ix, iy)
                  End If
               Next ix
            End If
            iyy = iyy - 1
         Next iy
         ReDim DATACUL1(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         FILL3D DATACUL1(), DATACULSRC()
      End If
   Case 2
      If ImBPP(2) = 32 Then
         For iy = UBound(DATACUL2, 3) To 0 Step -1
            If iyy >= 0 Then
               For ix = 0 To UBound(DATACUL2, 2)
                  If ix <= ImageWidth(ImageNum) - 1 Then
                     DATACULSRC(0, ix, iyy) = DATACUL2(0, ix, iy)
                     DATACULSRC(1, ix, iyy) = DATACUL2(1, ix, iy)
                     DATACULSRC(2, ix, iyy) = DATACUL2(2, ix, iy)
                     DATACULSRC(3, ix, iyy) = DATACUL2(3, ix, iy)
                  End If
               Next ix
            End If
            iyy = iyy - 1
         Next iy
         ReDim DATACUL2(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         FILL3D DATACUL2(), DATACULSRC()
      End If
   End Select
   
   End If
   
   LabWH(ImageNum) = "W x H =" & Str$(ImageWidth(ImageNum)) & " x" & Str$(ImageHeight(ImageNum))
   
   If ImageHeight(ImageNum) <= 32 And ImageWidth(ImageNum) <= 32 Then
      GridMult = 12
   ElseIf ImageHeight(ImageNum) <= 38 And ImageWidth(ImageNum) <= 38 Then
      GridMult = 10
   ElseIf ImageHeight(ImageNum) <= 48 And ImageWidth(ImageNum) <= 48 Then
      GridMult = 8
   ElseIf ImageHeight(ImageNum) <= 64 And ImageWidth(ImageNum) <= 64 Then
      GridMult = 6
   
   ' > 64 is too big for image views
'   ElseIf ImageHeight(ImageNum) <= 96 And ImageWidth(ImageNum) <= 96 Then
'      GridMult = 4   '64x64 ~ 250x250 box might be useable as 2 boxes
'                     ' for image and alpha bytes display
'   Else
'      GridMult = 3   '64x64 ~ 192x192 box might be useable as 3 boxes
'                     ' for image, mask and alpha bytes display
'      GridMult = 3   '48x48 ~ 144x144 box might be useable as 3 boxes
'                     ' for image, mask and alpha bytes display
   
   End If
   
   picPANEL.Height = ImageHeight(ImageNum) * GridMult
   picPANEL.Width = ImageWidth(ImageNum) * GridMult
   picPANEL.Left = (picC.Width - picPANEL.Width) \ 2 - 1
   picPANEL.Top = (picC.Height - picPANEL.Height) \ 2 - 1
   
   picPANEL.Picture = picPANEL.Image
   
   DrawGrid   ' also does TransferSmallToLarge picSmall(ImageNum), picPANEL
   
   ' Test for cursor size
   If (ImageHeight(ImageNum) = 32 And ImageWidth(ImageNum) = 32) Or _
      (ImageHeight(ImageNum) = 48 And ImageWidth(ImageNum) = 48) Or _
      (ImageHeight(ImageNum) = 64 And ImageWidth(ImageNum) = 64) Then
      
      Label5(2).Enabled = True
      Picture5.Enabled = True
      cmdFlash.Enabled = True
      aTestCursor = True
   Else
      Label5(2).Enabled = False
      Picture5.Enabled = False
      cmdFlash.Enabled = False
      aTestCursor = False
   End If
End Sub

Private Sub scrWidth_Change()
   If aSelect Then TURNSELECTOFF
   Call scrWidth_Scroll
End Sub

Private Sub scrWidth_Scroll()
Dim ix As Long, iy As Long
   If Not aScroll Then Exit Sub
   
   ImageWidth(ImageNum) = scrWidth.Value   ' Used throughout
   LabWidth = "Width  =" & Str$(ImageWidth(ImageNum))
   
   If picSmall(ImageNum).BackColor <> TColorBGR Then
      picSmall(ImageNum).BackColor = TColorBGR
   End If
   picSmall(ImageNum).Width = ImageWidth(ImageNum)
   picSmall(ImageNum).Picture = picSmall(ImageNum).Image
   picSmall(ImageNum).Height = ImageHeight(ImageNum)
   picSmall(ImageNum).Picture = picSmall(ImageNum).Image
   
   If ImBPP(ImageNum) = 32 Then
   
      ReDim DATACULSRC(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
      
      Select Case ImageNum
      Case 0
         For iy = 0 To UBound(DATACUL0(), 3)  'iy
            If iy <= ImageHeight(ImageNum) - 1 Then
               For ix = 0 To UBound(DATACUL0(), 2)   ' ix
                  If ix <= ImageWidth(ImageNum) - 1 Then
                     DATACULSRC(0, ix, iy) = DATACUL0(0, ix, iy)
                     DATACULSRC(1, ix, iy) = DATACUL0(1, ix, iy)
                     DATACULSRC(2, ix, iy) = DATACUL0(2, ix, iy)
                     DATACULSRC(3, ix, iy) = DATACUL0(3, ix, iy)
                  End If
               Next ix
            End If
         Next iy
         ReDim DATACUL0(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         FILL3D DATACUL0(), DATACULSRC()
         
      Case 1
         For iy = 0 To UBound(DATACUL1(), 3)  ' iy
            If iy <= ImageHeight(ImageNum) - 1 Then
               For ix = 0 To UBound(DATACUL1(), 2)  ' ix
                  If ix <= ImageWidth(ImageNum) - 1 Then
                     DATACULSRC(0, ix, iy) = DATACUL1(0, ix, iy)
                     DATACULSRC(1, ix, iy) = DATACUL1(1, ix, iy)
                     DATACULSRC(2, ix, iy) = DATACUL1(2, ix, iy)
                     DATACULSRC(3, ix, iy) = DATACUL1(3, ix, iy)
                  End If
               Next ix
            End If
         Next iy
         ReDim DATACUL1(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         FILL3D DATACUL1(), DATACULSRC()
         
      Case 2
         For iy = 0 To UBound(DATACUL2, 3)
            If iy <= ImageHeight(ImageNum) - 1 Then
               For ix = 0 To UBound(DATACUL2, 2)
                  If ix <= ImageWidth(ImageNum) - 1 Then
                     'AlphaSRC(ix, iy) = Alpha2(ix, iy)
                     DATACULSRC(0, ix, iy) = DATACUL2(0, ix, iy)
                     DATACULSRC(1, ix, iy) = DATACUL2(1, ix, iy)
                     DATACULSRC(2, ix, iy) = DATACUL2(2, ix, iy)
                     DATACULSRC(3, ix, iy) = DATACUL2(3, ix, iy)
                  End If
               Next ix
            End If
         Next iy
         ReDim DATACUL2(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
         FILL3D DATACUL2(), DATACULSRC()
         
      End Select
   
   End If
   
   LabWH(ImageNum) = "W x H =" & Str$(ImageWidth(ImageNum)) & " x" & Str$(ImageHeight(ImageNum))
   If ImageHeight(ImageNum) <= 32 And ImageWidth(ImageNum) <= 32 Then
      GridMult = 12
   ElseIf ImageHeight(ImageNum) <= 38 And ImageWidth(ImageNum) <= 38 Then
      GridMult = 10
   ElseIf ImageHeight(ImageNum) <= 48 And ImageWidth(ImageNum) <= 48 Then
      GridMult = 8
   ElseIf ImageHeight(ImageNum) <= 64 And ImageWidth(ImageNum) <= 64 Then
      GridMult = 6
   
   ' Too big for image views
'   ElseIf ImageHeight(ImageNum) <= 96 And ImageWidth(ImageNum) <= 96 Then
'      GridMult = 4
'   Else
'      GridMult = 3
   End If
   
   picPANEL.Height = ImageHeight(ImageNum) * GridMult
   picPANEL.Width = ImageWidth(ImageNum) * GridMult
   picPANEL.Top = (picC.Height - picPANEL.Height) \ 2 - 1
   picPANEL.Left = (picC.Width - picPANEL.Width) \ 2 - 1
   picPANEL.Picture = picPANEL.Image
   
   DrawGrid   ' also does TransferSmallToLarge picSmall(ImageNum), picPANEL
   
   ' Test for cursor size
   If (ImageHeight(ImageNum) = 32 And ImageWidth(ImageNum) = 32) Or _
      (ImageHeight(ImageNum) = 48 And ImageWidth(ImageNum) = 48) Or _
      (ImageHeight(ImageNum) = 64 And ImageWidth(ImageNum) = 64) Then
      Label5(2).Enabled = True
      Picture5.Enabled = True
      cmdFlash.Enabled = True
      aTestCursor = True
   Else
      Label5(2).Enabled = False
      Picture5.Enabled = False
      cmdFlash.Enabled = False
      aTestCursor = False
   End If
End Sub

'#### GRID ####

Private Sub chkGridOnOff_Click()
   If aGrid = 2 Then Exit Sub
   aGrid = chkGridOnOff.Value
   If aGrid = 1 Then
      chkGridOnOff.Caption = "Grid On"
   Else
      chkGridOnOff.Caption = "Grid Off"
   End If
   picPANEL.SetFocus
   DrawGrid
End Sub

Public Sub DrawGrid()
'Public GridMult As Long
'Public GridLinesCul As Long
Dim k As Long
Dim ix As Long, iy As Long
Dim mix As Long, miy As Long
Dim Cul As Long
   
   picPANEL.Picture = LoadPicture
   TransferSmallToLarge picSmall(ImageNum), picPANEL  ' Also sets picSmallDATA(0,0)
   
   'For permanently showing masks add picboxes picMask(0 to 2)
   'ShowMask picSmall(ImageNum), picMask(ImageNum)
   
   If aGrid = 1 Then
      ' Horz lines
      For k = 0 To GridMult * ImageHeight(ImageNum) Step GridMult
         picPANEL.Line (0, k)-(picPANEL.Width, k), GridLinesCul
      Next k
      ' Vert lines
      For k = 0 To GridMult * ImageWidth(ImageNum) Step GridMult
         picPANEL.Line (k, 0)-(k, picPANEL.Height), GridLinesCul
      Next k
   
      ' Have picSmallDATA(0 To ImageWidth(ImageNum)-1, 0 To Imageheight(ImageNum)-1))
      ' from TransferSmallToLarge
      ' MarkPixels =
      '  9 Mark transparent pixel-Square
      ' 10 Mark transparent pixel-Diagonal
      ' 11 Unmark transparent pixels
      If MarkPixels = 9 Or MarkPixels = 10 Then
         'LngToRGB TColorBGR, SR, SG, SB
         'TColorRGB = RGB(SB, SG, SR)
         For iy = 0 To ImageHeight(ImageNum) - 1
         miy = iy * GridMult + 1
         For ix = 0 To ImageWidth(ImageNum) - 1
            Cul = picSmallDATA(ix, ImageHeight(ImageNum) - iy - 1) And &HFFFFFF
            If Cul > 0 Then
            If Cul = TColorRGB Then
               mix = ix * GridMult + 1
               '/
               If MarkPixels = 10 Then
                  picPANEL.Line (mix, miy)-(mix + GridMult - 1, miy + GridMult - 1), vbWhite
               ElseIf MarkPixels = 9 Then  ' Square
                  picPANEL.Line (mix, miy)-(mix + GridMult - 2, miy + GridMult - 2), vbWhite, B
               End If
               ' \
               ' picPANEL.Line (mix, miy + GridMult - 2)-(ix * GridMult + GridMult, iy * GridMult), vbWhite
            End If
            End If
         Next ix
         Next iy
      End If
   End If
   
   ' Draw eTools2 lines on picC
   Dim cx1 As Single, cx2 As Single
   Dim cy1 As Single, cy2 As Single
      picC.Cls
      ' Centre lines
      picC.Line (0, picC.ScaleHeight \ 2)-(picC.ScaleWidth, picC.ScaleHeight \ 2), vbRed
      picC.Line ((picC.ScaleWidth - 1) \ 2, 0)-((picC.ScaleWidth - 1) \ 2, picC.ScaleHeight), vbRed
      ' Right & bottom eTools2 lines
      cx1 = picPANEL.Left
      cx2 = cx1 + picPANEL.Width
      cy1 = picPANEL.Top
      cy2 = cy1 + picPANEL.Height
      picC.Line (cx2, cy1)-(cx2, cy2), GridLinesCul
      picC.Line (cx1, cy2)-(cx2 + 1, cy2), GridLinesCul
      
      ' Draw around current picSmall
      cx1 = picSmall(ImageNum).Left - 1
      cx2 = cx1 + picSmall(ImageNum).Width + 1
      cy1 = picSmall(ImageNum).Top - 1
      cy2 = cy1 + picSmall(ImageNum).Height + 1
      picSmallFrame.Cls
      picSmallFrame.DrawStyle = 2
      picSmallFrame.Line (cx1 - 1, cy1 - 1)-(cx2 + 1, cy2 + 1), &H8080F8, B
   
   If (ImageWidth(ImageNum) = 32 And ImageHeight(ImageNum) = 32) Or _
      (ImageWidth(ImageNum) = 48 And ImageHeight(ImageNum) = 48) Or _
      (ImageWidth(ImageNum) = 64 And ImageHeight(ImageNum) = 64) Then
      mnuSaveCUR.Enabled = True
      LabHotXY(0) = "X=" & HotX(ImageNum)
      LabHotXY(1) = "Y=" & HotY(ImageNum)
      aTestCursor = False
      scrHotX.Value = HotX(ImageNum)
      scrHotY.Value = HotY(ImageNum)
      aTestCursor = True
      For k = 0 To 2
         If ImageWidth(k) = 32 And ImageHeight(k) = 32 Or _
            ImageWidth(k) = 48 And ImageHeight(k) = 48 Or _
            ImageWidth(k) = 64 And ImageHeight(k) = 64 Then
            mnuSaveCURImage(k).Enabled = True
         Else
            mnuSaveCURImage(k).Enabled = False
         End If
      Next k
   Else
      mnuSaveCUR.Enabled = False
   End If

End Sub


Private Sub picGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
   
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If
   
   If aSelect Then
      TURNSELECTOFF
   End If
   
   Cul = picGrid.Point(X, Y)
   If Cul >= 0 Then
      GridLinesCul = Cul
      DrawGrid
   End If
End Sub


'#### SHOW MASKS, ALPHA & ORIGINAL COLORS ####

Private Sub cmdMask_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Public Sub ShowMask(picS As PictureBox, picM As PictureBox)
Dim k As Long
Dim sw As Long, sh As Long
   
   If LabMAC = "Mask" Or LabMAC = "Alphas" Or LabMAC = "Original colors" Then
      cmdMask_MouseUp 1, 0, 0, 0
      Exit Sub
   End If
   
   LabMAC = "Mask"
   'For k = 0 To 2   ' To show all panes
   k = ImageNum
      sw = picSmall(k).Width
      sh = picSmall(k).Height
      With picTemp(k)
         .Picture = LoadPicture
         .Width = sw
         .Height = sh
         BitBlt .hdc, 0, 0, sw, sh, picSmall(k).hdc, 0, 0, vbSrcCopy
         .Picture = .Image
      End With
   'Next k
      
   'For k = 0 To 2
      ShowMask picSmall(k), picSmall(k)
   'Next k
   aGrid = 0
   DrawGrid
   aGrid = 2
   chkGridOnOff.Value = vbChecked
   aGrid = 1
   chkGridOnOff.Caption = "Grid On"
   picPANEL.SetFocus
End Sub

Private Sub cmdMask_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RestoreImage
End Sub

Private Sub cmdAlpha_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Show Alphas
Dim k As Long
Dim sw As Long, sh As Long
   
   If LabMAC = "Masks" Or LabMAC = "Alphas" Or LabMAC = "Original colors" Then
      cmdMask_MouseUp 1, 0, 0, 0
      Exit Sub
   End If

   LabMAC = "Alphas"
   'For k = 0 To 2
   k = ImageNum
      sw = picSmall(k).Width
      sh = picSmall(k).Height
      With picTemp(k)
         .Picture = LoadPicture
         .Width = sw
         .Height = sh
         BitBlt .hdc, 0, 0, sw, sh, picSmall(k).hdc, 0, 0, vbSrcCopy
         .Picture = .Image
      End With
   'Next k
      
   'If ImBPP(0) <> 32 And ImBPP(1) <> 32 And ImBPP(2) <> 32 Then  ' If all panes
   If ImBPP(k) <> 32 Then
     LabMAC = "No Alphas"
   Else
      'For k = 0 To 2
         If ImBPP(k) = 32 Then
            ShowAlpha k, picSmall(k), picSmall(k)
         End If
      'Next k
   End If
   aGrid = 0
   DrawGrid
   aGrid = 2
   chkGridOnOff.Value = vbChecked
   aGrid = 1
   chkGridOnOff.Caption = "Grid On"
   If aShow Then picPANEL.SetFocus
End Sub

Private Sub cmdAlpha_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RestoreImage
End Sub

Private Sub cmdORGCul_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Show original colors
Dim k As Long
Dim sw As Long, sh As Long

   If LabMAC = "Original colors" Then
      cmdMask_MouseUp 1, 0, 0, 0
      Exit Sub
   End If
   
   LabMAC = "Original colors"
   k = ImageNum
   ' Save picSmall() in picTemp()
   'For k = 0 To 2
      sw = picSmall(k).Width
      sh = picSmall(k).Height
      With picTemp(k)
         .Picture = LoadPicture
         .Width = sw
         .Height = sh
         BitBlt .hdc, 0, 0, sw, sh, picSmall(k).hdc, 0, 0, vbSrcCopy
         .Picture = .Image
      End With
   'Next k

   'For k = 0 To 2
      ShowWithTColor k, picSmall(k), picSmall(k), 1
   'Next k
   aGrid = 0
   DrawGrid
   aGrid = 2
   chkGridOnOff.Value = vbChecked
   aGrid = 1
   chkGridOnOff.Caption = "Grid On"
   'cmdORGCul.SetFocus
   If aShow Then picPANEL.SetFocus
End Sub

Private Sub cmdORGCul_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RestoreImage
End Sub

Private Sub RestoreImage()
' Restore original image
Dim k As Long
Dim sw As Long, sh As Long
   LabMAC = "Images"
   ' Restore picSmall() from picTemp()
   'For k = 0 To 2
   k = ImageNum
      sw = picSmall(k).Width
      sh = picSmall(k).Height
      With picSmall(k)
         .Picture = LoadPicture
         .Width = sw
         .Height = sh
         BitBlt .hdc, 0, 0, sw, sh, picTemp(k).hdc, 0, 0, vbSrcCopy
         .Picture = .Image
      End With
   'Next k
   picSmallFrame.BackColor = &H808080 'TColorBGR
   DrawGrid
End Sub

Private Function IsBlank(ImNum As Integer) As Boolean
Dim ix As Long, iy As Long
Dim Cul As Long
   IsBlank = True
   For iy = 0 To ImageHeight(ImNum) - 1
   For ix = 0 To ImageWidth(ImNum) - 1
      Cul = picSmall(ImNum).Point(ix, iy)
      If picSmall(ImNum).Point(ix, iy) <> TColorBGR Then
         IsBlank = False
         Exit For
      End If
   Next ix
   Next iy
End Function


'#### Locate controls ####

Private Sub LocateCtrls_Init()
Dim k As Long
Dim SR As Byte, SG As Byte, SB As Byte

   For k = 1 To 12 '8 [Line] to [Text]
      With optTools(k)
         ' Left, Top, Width, Height
         .Move optTools(k - 1).Left + optTools(0).Width + 3, optTools(0).Top, _
               optTools(0).Width, optTools(0).Height
      End With
   Next k
   
   For k = [BoxVertShade] To [BoxHVCenShade]  ' 24 - 32
      With optTools(k)
         .Move optTools(k - 1).Left + optTools(k).Width + 3, optTools(k - 1).Top, _
               optTools(k - 1).Width, optTools(k - 1).Height
      End With
   Next k
   
   With Line2(0)
   .x1 = optTools(8).Left + optTools(8).Width + 3
   .x2 = .x1
   .y1 = 0
   .y2 = picToolbar.Height
   End With
   
   With Line2(0)
      Line2(1).x1 = .x1 + 1
      Line2(1).x2 = Line2(1).x1
      Line2(1).y1 = .y1
      Line2(1).y2 = .y2
   End With
   
   For k = 1 To cmdPAL.Count - 1
      With cmdPAL(k)
         .Top = cmdPAL(0).Top
         .Left = cmdPAL(k - 1).Left + 25
      End With
   Next k
   
   TColorBGR = RGB(194, 195, 197)   ' 32bit unlikely color for TColorBGR
   'TColorBGR = RGB(198, 207, 214)  ' 16bit & 32 bit unlikely color
                                    ' OK on some PCs but
                                    ' does not work on all PC OSs
                                    ' different dithering methods.
   LngToRGB TColorBGR, SR, SG, SB
   TColorRGB = RGB(SB, SG, SR)
   
   For k = 0 To 2
      With picSmall(k)
         .Width = 48 + 2
         .Height = 48 + 2
         .BackColor = TColorBGR
      End With
      LabWH(k).Left = 0
      LabWH(k).Width = 83
      LabNumColors(k).Left = 1
      LabNumColors(k).Top = LabWH(k).Top - LabNumColors(k).Height - 1
      LabNumColors(k).BackColor = &HE0E0E0
      picSEL(k).Visible = False
   
      picSmall(k).MouseIcon = LoadResPicture("REDARROW", vbResCursor)
      picSmall(k).MousePointer = vbCustom
   Next k
   
   If picSmall(0).Point(1, 1) <> TColorBGR Then
      TColorBGR = picSmall(0).Point(1, 1) + 8
      For k = 0 To 2
         With picSmall(k)
            .Width = 48 + 2
            .Height = 48 + 2
            .BackColor = TColorBGR
         End With
         LabNumColors(k).BackColor = &HE0E0E0
      Next k
   End If
      
   If picSmall(0).Point(1, 1) <> TColorBGR Then
      TColorBGR = picSmall(0).Point(1, 1)
      For k = 0 To 2
         With picSmall(k)
            .Width = 48 + 2
            .Height = 48 + 2
            .BackColor = TColorBGR
         End With
         LabNumColors(k).BackColor = &HE0E0E0 'TColorBGR
      Next k
   End If
   
   picSmallFrame.BackColor = &H808080 'TColorBGR
   
   picPANEL.Width = 48 * GridMult
   picPANEL.Height = 48 * GridMult
   picPANEL.BackColor = TColorBGR
   
   For k = 0 To 1
      Line1(k).Visible = False
      Box1(k).Visible = False
      Ellipse1(k).Visible = False
   Next k
   shpCirc.Visible = False
   ' Start colors
   picPAL.Width = 386 '258
   picVisColor.Width = 15
   picGrid.Width = 15
   
   If Not FileExists(AppPathSpec$ & "GFX.ini") Then
      PALIndex = 4
   End If
   cmdPAL_Click PALIndex
   'QBColors picPAL
   
   QBColors picGrid, 2
   Picture1.Line (picGrid.Left - 1, picGrid.Top - 1)-(picGrid.Left + picGrid.Width, picGrid.Top + picGrid.Height), 0, B
   
   Select Case PALIndex
   Case 0: QBColors picVisColor, 1 '2
   Case 1: ShortBandedPAL picVisColor, 2
   Case 2: LongBandedPAL picVisColor, 2
   Case 3: GreyPAL picVisColor, 2
   Case 4: CenteredPAL picVisColor, 2
   End Select
   'QBColors picVisColor, 3
   Picture1.Line (picVisColor.Left - 1, picVisColor.Top - 1)-(picVisColor.Left + picVisColor.Width, picVisColor.Top + picVisColor.Height), 0, B
   
   
   LColor = 0
   RColor = vbWhite
   LabColor(0).BackColor = LColor
   LabColor(1).BackColor = RColor
   picSmallBU.BackColor = TColorBGR
   
   picC.Height = 391
   picC.Width = picC.Height
   picC.BackColor = RGB(180, 181, 182)
   
   Picture1.DrawWidth = 2
   Picture1.PSet (8 * 16 - 1, picVisColor.Top - 3), vbRed

   aScroll = False
   scrWidth.Max = MaxWidth
   scrHeight.Max = MaxHeight
   
   scrHeight.Max = 64 ' to allow Ht = 4
   
   aScroll = True
   
   ' Headers Tools, Effect, Select
   Label8(0).Left = picToolbar.Left + 1
   picElements.Left = picToolbar.Left + optTools(13).Left + 3
   picElements.CurrentY = 0
   picElements.CurrentX = 50
   picElements.DrawStyle = 2
   picElements.Print "Effects";
   picElements.Line (100, 2)-(114, 10), 0, B
   
   Label8(2).Left = picToolbar.Left + optTools(20).Left - 2
   
   ' Start images
   For ImageNum = 2 To 0 Step -1
      picORG(ImageNum).BackColor = TColorBGR
      picORG(ImageNum).Width = ImageWidth(ImageNum)
      picORG(ImageNum).Height = ImageHeight(ImageNum)
   Next ImageNum
   
   'aScroll = False
   For ImageNum = 2 To 0 Step -1
      scrWidth.Value = ImageWidth(ImageNum)
      scrHeight.Value = 65 - ImageHeight(ImageNum)
   Next ImageNum
   aScroll = True
   
   LineErase.y1 = picErase.Top + 1
   LineErase.y2 = picErase.Top + picErase.Height - 1
   LineErase.Visible = False
   
   LabTest.Top = LabSpec.Top + LabSpec.Height + 2
End Sub

'#### QueryUnload & INI code ####

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim k As Long
Dim Form As Form
   
   If UnloadMode = 0 Then
      k = MsgBox("If you need to save or re-save an image(s) press No." & vbCrLf & _
      "Press Yes to Exit.", vbQuestion + vbYesNo, "Exit TinyGFX ?")
      If k = vbNo Then
         Cancel = True
         Exit Sub
      End If
   End If
   
   If aClipBoard = True Then
      k = MsgBox("Clipboard used,  Clear?", vbQuestion + vbYesNo, "Exitting")
      If k = vbYes Then
         Clipboard.Clear
      End If
   End If
   
   For k = 0 To 2
      KILLSAVS k
   Next k
   
   TURNOFFCURSOR
   
   IniSpec$ = AppPathSpec$ & IniTitle$ & ".ini"
   If FileExists(IniSpec$) Then
      Kill IniSpec$
   End If
   
   'WriteINI(Title$, TheKey$, Info$, ISpec$)
   WriteINI "RecentFiles", "NumRecentFiles", Str$(NumRecentFiles), IniSpec$
   If NumRecentFiles > 0 Then
      For k = 1 To NumRecentFiles
         'WriteINI "RecentFiles", Str$(k) & ".", mnuRecentFiles(k).Caption, IniSpec$
         WriteINI "RecentFiles", Str$(k) & ".", FileArray$(k), IniSpec$
      Next k
   End If
   
   WriteINI "LastFolder", "LastFolder", CPath$, IniSpec$
   
   
   WriteINI "Preferences", "1.", Str$(OptimizeNumber), IniSpec$
   WriteINI "Preferences", "2.", Str$(ToneNumber), IniSpec$
   WriteINI "Preferences", "3.", Str$(AspectNumber), IniSpec$
   WriteINI "Preferences", "4.", Str$(MarkPixels), IniSpec$ ' 9,10,11
   WriteINI "Preferences", "5.", Str$(ExtractorBackups), IniSpec$
   ' aAlphaRestricted =  True, Checked,   Alpha OFF
   ' aAlphaRestricted = False, Unchecked, Alpha ON
   WriteINI "Preferences", "6.", Str$(aAlphaRestricted), IniSpec$
   WriteINI "Preferences", "7.", Str$(aQueryOC), IniSpec$
   
   WriteINI "GridCul", "GridCul", Trim$(Str$(GridLinesCul)), IniSpec$
   WriteINI "Palette", "PALIndex", Trim$(Str$(PALIndex)), IniSpec$
   WriteINI "frmRotator", "frmRotateLeft", Trim$(Str$(frmRotateLeft)), IniSpec$
   WriteINI "frmRotator", "frmRotateTop", Trim$(Str$(frmRotateTop)), IniSpec$
   
   If hwndHelp <> 0 Then
      HtmlHelp Me.hwnd, "", HH_CLOSE_ALL, 0
   End If

   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form

'   Unload Me
'   End
   Form_Terminate
End Sub

Private Sub Form_Terminate()
   End
End Sub


Private Sub GetINI_Info()
' Preferences
'Public OptimizeNumber As Integer   ' 0,1  Opt or Non-opt saving
'Public ToneNumber As Long          ' 3,4  Halftone or ColorOnColor mode
'Public AspectNumber As Long        ' 6,7  Keep or ignore aspect
'Public MarkPixels As Long          ' 9,10 Marked/Unmarked oixels
Dim k As Long, kn As Long
Dim TheNAM$, ThePATH$, a$
Dim Ret$
   
   ' Default preferences
   OptimizeNumber = 0      ' (Optimized saving)
   ToneNumber = 3          ' (HALFTONE)
   AspectNumber = 6        ' (Keep Aspect ratio)
   MarkPixels = 9          ' (Mark transparent pixels square)
   ExtractorBackups = 14   ' (Back up last Extractor image)

   GridLinesCul = RGB(128, 128, 128)
   IniSpec$ = AppPathSpec$ & IniTitle$ & ".ini"
   
   If GetINI("RecentFiles", "NumRecentFiles", Ret$, IniSpec$) Then
      NumRecentFiles = Val(Ret$)
      FileCount = NumRecentFiles
   End If
   
   If NumRecentFiles > 0 Then
      mnuRecentFiles(0).Visible = True  '-------
      k = 1
      kn = 1
      Do
         GetINI "RecentFiles", Str$(kn) & ".", Ret$, IniSpec$
         If Len(Ret$) <> 0 And FileExists(Ret$) Then
            FileArray$(k) = Ret$
            'mnuRecentFiles(k).Caption = Ret$
            ' Shorten
            TheNAM$ = GetFileName(FileArray$(k))
            ThePATH$ = GetPath(FileArray$(k))
            a$ = TheNAM$ & ": " & ThePATH$
            mnuRecentFiles(k).Caption = LTrim$(Str$(k)) & ". " & ShortenString$(a$, 42)

            
            'mnuRecentFiles(k).Caption = LTrim$(Str$(k)) & ". " & ShortenString$(Ret$, 40)
            
            mnuRecentFiles(k).Visible = True
            k = k + 1
            kn = kn + 1
            If kn > NumRecentFiles Then Exit Do
         Else
            kn = kn + 1
            If kn > NumRecentFiles Then Exit Do
         End If
      Loop
      NumRecentFiles = k - 1
      FileCount = NumRecentFiles
   End If
   
   If GetINI("LastFolder", "LastFolder", Ret$, IniSpec$) Then
      CPath$ = Ret$
   End If
  
   
   For k = 1 To 7
      If GetINI("Preferences", Str$(k) & ".", Ret$, IniSpec$) Then
         Select Case k
         Case 1
            OptimizeNumber = Val(Ret$) ' 0,1
         Case 2
            ToneNumber = Val(Ret$)     ' 3,4
         Case 3
            AspectNumber = Val(Ret$)   ' 6,7
         Case 4
            MarkPixels = Val(Ret$)     ' 9,10,11
         Case 5
            ExtractorBackups = Val(Ret$) ' 13,14
         Case 6
            aAlphaRestricted = CBool(Ret$)  ' 255,True or 0,False 16
         Case 7
            aQueryOC = CBool(Ret$)  ' 255,True or 0,False 18
         End Select
      Else
         Exit For
      End If
   Next k
   mnuPreferences_Click (CInt(OptimizeNumber))
   mnuPreferences_Click (CInt(ToneNumber))
   mnuPreferences_Click (CInt(AspectNumber))
   
   mnuPreferences(9).Checked = False
   mnuPreferences(10).Checked = False
   mnuPreferences(11).Checked = False
   mnuPreferences(CInt(MarkPixels)).Checked = True
   mnuPreferences(13).Checked = False
   mnuPreferences(14).Checked = False
   mnuPreferences(CInt(ExtractorBackups)).Checked = True

   If aAlphaRestricted Then
      mnuPreferences(16).Checked = True  ' Alpha restricted, checked
      mnuOpenIntoImage(4).Visible = True
   Else   ' aAlphaRestricted = False
      mnuPreferences(16).Checked = False ' Alpha used, unchecked
      mnuOpenIntoImage(4).Visible = False
   End If
   
   If aQueryOC Then
      mnuPreferences(18).Checked = True  ' Query Original colors, checked
   Else   ' aQueryOC = False
      mnuPreferences(18).Checked = False ' aQueryOC unchecked
   End If
   
   
   If GetINI("GridCul", "GridCul", Ret$, IniSpec$) Then
      GridLinesCul = Val(Ret$)
   End If

   If GetINI("Palette", "PALIndex", Ret$, IniSpec$) Then
      PALIndex = Val(Ret$)
      If PALIndex = 5 Then ' ie System
         PALIndex = 4
      End If
   End If
   If GetINI("frmRotator", "frmRotateLeft", Ret$, IniSpec$) Then
      frmRotateLeft = Val(Ret$)
      GetINI "frmRotator", "frmRotateTop", Ret$, IniSpec$
      frmRotateTop = Val(Ret$)
   End If
   
End Sub

Private Sub mnuExit_Click()
   Form_QueryUnload 1, 0
End Sub

Public Sub KILLSAVS(Im As Long)
' im = 0,1 or 2 ImageNum
Dim gen As Long
   
   For gen = 1 To MaxGen(Im)
      BUSpec$ = AppPathSpec$ & "SAV"
      BUSpec$ = BUSpec$ & Trim$(Str$(Im)) & Trim$(Str$(gen)) & ".dat"
      If FileExists(BUSpec$) Then Kill BUSpec$
      CurrentGen(Im) = 0
   Next gen
   MaxGen(Im) = 0
   cmdUNDO(Im).Enabled = False
   cmdRedo(Im).Enabled = False
   LabCurGen(Im) = 0 'CurrentGen(ImageNum)
   LabMaxGen(Im) = 0 'MaxGen(ImageNum)
End Sub

Private Sub mnuHelp_Click()
   If LabMAC <> "Images" Then
      RestoreImage
      Exit Sub
   End If

   HelpSpec$ = AppPathSpec$ & "tinygfx.chm"
   If FileExists(HelpSpec$) Then
       hwndHelp = HtmlHelp(hwnd, HelpSpec$, HH_DISPLAY_TOPIC, 0)
   Else
      HelpSpec$ = AppPathSpec$ & "GFXHelp.txt"
      MsgBox "tinygfx.chm missing", vbInformation + vbApplicationModal, "TinyGFX help"
      HelpSpec$ = AppPathSpec$ & "GFXHelp.txt"
      If Not FileExists(HelpSpec$) Then
         MsgBox "GFXHelp.txt missing", vbInformation + vbApplicationModal, "TinyGFX help"
         Exit Sub
      End If
      frmHelp.Show vbModal, Me
   End If
End Sub

