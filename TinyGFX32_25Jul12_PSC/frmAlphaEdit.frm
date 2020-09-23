VERSION 5.00
Begin VB.Form frmAlphaEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Alpha Editing"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   543
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   508
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   6870
      ScaleHeight     =   2325
      ScaleWidth      =   660
      TabIndex        =   41
      Top             =   1155
      Width           =   690
      Begin VB.OptionButton optTool 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   150
         Picture         =   "frmAlphaEdit.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   " Dot "
         Top             =   285
         Width           =   315
      End
      Begin VB.OptionButton optTool 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   150
         Picture         =   "frmAlphaEdit.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   " Line "
         Top             =   645
         Width           =   315
      End
      Begin VB.OptionButton optTool 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   150
         Picture         =   "frmAlphaEdit.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   " Box "
         Top             =   1020
         Width           =   315
      End
      Begin VB.CommandButton cmdUNDO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Undo Last "
         Height          =   450
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1800
         Width           =   600
      End
      Begin VB.OptionButton optTool 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   150
         Picture         =   "frmAlphaEdit.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   " Solid box "
         Top             =   1365
         Width           =   315
      End
      Begin VB.Label LabTool 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Line"
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
         Left            =   45
         TabIndex        =   47
         Top             =   0
         Width           =   525
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6285
      ScaleHeight     =   345
      ScaleWidth      =   1170
      TabIndex        =   37
      Top             =   5445
      Width           =   1200
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   40
         Top             =   60
         Width           =   360
      End
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G"
         ForeColor       =   &H0000C000&
         Height          =   240
         Index           =   1
         Left            =   405
         TabIndex        =   39
         Top             =   60
         Width           =   360
      End
      Begin VB.Label LabRGB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   780
         TabIndex        =   38
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   5040
      ScaleHeight     =   1050
      ScaleWidth      =   975
      TabIndex        =   34
      Top             =   120
      Width           =   1005
      Begin VB.CommandButton cmdAlphaGrad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Alpha gradient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   " Edge gradient from Set Alpha to small values in 6 steps "
         Top             =   45
         Width           =   855
      End
      Begin VB.ComboBox cboSteps 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   35
         ToolTipText     =   " Steps for Alpha gradient "
         Top             =   675
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   6285
      ScaleHeight     =   1035
      ScaleWidth      =   1170
      TabIndex        =   23
      Top             =   5835
      Width           =   1200
      Begin VB.Label LabA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alpha"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   105
         TabIndex        =   32
         ToolTipText     =   " Underlying Alpha value "
         Top             =   90
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set  Alpha"
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
         Left            =   75
         TabIndex        =   26
         Top             =   465
         Width           =   1020
      End
      Begin VB.Label LabAlpha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   315
         TabIndex        =   25
         Top             =   720
         Width           =   405
      End
      Begin VB.Label LabVal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   630
         TabIndex        =   24
         ToolTipText     =   " Underlying Alpha value "
         Top             =   60
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdFillTAreas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fill transparent areas with set color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4290
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7350
      Width           =   1845
   End
   Begin VB.PictureBox picSmallColors 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   3630
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   17
      ToolTipText     =   " Basic underlying colors. "
      Top             =   255
      Width           =   960
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6345
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7710
      Width           =   1020
   End
   Begin VB.PictureBox picSmallMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   420
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   11
      Top             =   255
      Width           =   960
   End
   Begin VB.PictureBox picSmallAlpha 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   2565
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   10
      Top             =   255
      Width           =   960
   End
   Begin VB.CommandButton cmdACCCAN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2790
      TabIndex        =   7
      Top             =   7710
      Width           =   1215
   End
   Begin VB.CommandButton cmdACCCAN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   225
      TabIndex        =   6
      Top             =   7710
      Width           =   1215
   End
   Begin VB.PictureBox picGrey 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   6375
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   5
      Top             =   540
      Width           =   405
   End
   Begin VB.PictureBox picGrey 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   6375
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   4
      Top             =   4725
      Width           =   405
   End
   Begin VB.PictureBox picGrey 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3870
      Index           =   1
      Left            =   6375
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   3
      Top             =   840
      Width           =   405
   End
   Begin VB.PictureBox picEditC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   6000
      Left            =   180
      ScaleHeight     =   396
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   396
      TabIndex        =   1
      Top             =   1260
      Width           =   6000
      Begin VB.PictureBox picEDIT 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   2145
         Left            =   0
         ScaleHeight     =   143
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   191
         TabIndex        =   2
         Top             =   0
         Width           =   2865
         Begin VB.Shape Box1 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   15  'Merge Pen Not
            Height          =   420
            Index           =   1
            Left            =   1170
            Top             =   840
            Width           =   450
         End
         Begin VB.Shape Box1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            Height          =   420
            Index           =   0
            Left            =   1005
            Top             =   315
            Width           =   450
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            DrawMode        =   15  'Merge Pen Not
            Index           =   1
            X1              =   177
            X2              =   134
            Y1              =   31
            Y2              =   66
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   6  'Inside Solid
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            Index           =   0
            X1              =   173
            X2              =   130
            Y1              =   20
            Y2              =   55
         End
      End
   End
   Begin VB.PictureBox picSmallEdit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   1500
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   255
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Drawing Tools "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6855
      TabIndex        =   48
      Top             =   765
      Width           =   690
   End
   Begin VB.Label LabLColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   7200
      TabIndex        =   33
      ToolTipText     =   " Right color : click to set "
      Top             =   4710
      Width           =   225
   End
   Begin VB.Label LabCul 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C0C000&
      Height          =   240
      Left            =   6945
      TabIndex        =   31
      Top             =   5100
      Width           =   450
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Set color"
      Height          =   420
      Index           =   4
      Left            =   6960
      TabIndex        =   30
      Top             =   3585
      Width           =   450
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " L   R"
      Height          =   195
      Index           =   3
      Left            =   6975
      TabIndex        =   29
      Top             =   4485
      Width           =   450
   End
   Begin VB.Label LabLColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   6945
      TabIndex        =   28
      ToolTipText     =   " Left color: click to set  "
      Top             =   4710
      Width           =   225
   End
   Begin VB.Label LabTest 
      BackColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   30
      TabIndex        =   27
      Top             =   8355
      Width           =   1005
   End
   Begin VB.Label LabClick 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Click to set Alpha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   21
      Top             =   7335
      Width           =   3990
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Transparent"
      Height          =   255
      Index           =   2
      Left            =   6510
      TabIndex        =   20
      Top             =   7335
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   135
      Index           =   1
      Left            =   6315
      Picture         =   "frmAlphaEdit.frx":1628
      Top             =   7395
      Width           =   135
   End
   Begin VB.Label LabSetCul 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6945
      TabIndex        =   19
      ToolTipText     =   " Set color "
      Top             =   4020
      Width           =   450
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Original colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   3615
      TabIndex        =   18
      Top             =   -15
      Width           =   1365
   End
   Begin VB.Label LabXY 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X = 64, Y = 64"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6300
      TabIndex        =   15
      Top             =   6975
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Mask"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   420
      TabIndex        =   14
      Top             =   0
      Width           =   585
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Alpha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2580
      TabIndex        =   13
      Top             =   0
      Width           =   540
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   1530
      TabIndex        =   12
      Top             =   0
      Width           =   570
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Transparent Alpha"
      Height          =   390
      Index           =   1
      Left            =   6150
      TabIndex        =   9
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opaque Alpha"
      Height          =   420
      Index           =   0
      Left            =   6270
      TabIndex        =   8
      Top             =   5025
      Width           =   645
   End
End
Attribute VB_Name = "frmAlphaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmAlphaEdit.frm

Option Explicit

Private aMDown As Boolean
Private GridCol As Long       ' Fixed so can be ignored
Private AlphaValue As Byte    ' 0 to 255
Private svDATACUL() As Byte   ' Save original data for Reset
Private tempDATACUL() As Byte ' Saved data for Undo Last
Private CulValue As Long      ' 255-AlphaValue

Private pMaskData() As Long

' Saved ix,iy from picEDIT_MouseDown & _MouseUp
Private six() As Long, siy() As Long
Private NPoints As Long

' Drawing tools
Private ATool As Integer   '0 Dot, 1 Line, 2 Box

' For Alpha Gradient
Private GradSteps As Long

'Public
'Private xStart As Single
'Private yStart  As Single
'Private xend As Single
'Private yend As Single

Private IACMNum As Long    ' 0,1,2,3  Image, Alpha, Original colors, Mask
Private OColor As Long     ' Set when editing Original colors




Private Sub cboSteps_Click()
Dim a$
   GradSteps = 2 + Val(cboSteps.ListIndex)
    a$ = " Edge gradient from Set Alpha to small values in" ' 6 steps "
    a$ = a$ & Str$(GradSteps) & " steps "
    cmdAlphaGrad.ToolTipText = a$
End Sub

' NB pixels outside the mask (ie mask color = white)
' are transparent and cannot be changed
' Inside the mask (ie mask color <> white),
' if alpha=0 then make mask transparent

' Orignal colors in
' DATACULSRC#(0 to 3,ImageWidth(ImageNum)-1,ImageHeight(ImageNum)-1)

Private Sub cmdACCCan_Click(Index As Integer)
   aAlphaEdit = False
   If Index = 0 Then 'ACCEPT
      If ImBPP(ImageNum) = 32 Then  ' ? not nec
         picSmallEdit.Picture = picSmallEdit.Image
         Reconcile picSmallEdit, ImageNum  ' Ensure Alpha and Image aligned
      End If
      TransferSmallToLarge picSmallEdit, Form1.picSmall(ImageNum)
      aAlphaEdit = True
   End If
   Unload Me
End Sub

Private Sub cmdAlphaGrad_Click()
' Sep 08
Dim iy As Long, ix As Long
Dim ABYTE As Long
Dim ab(0 To 3) As Byte
Dim aT As Boolean
Dim j As Long, k As Long
Dim NP As Long

Dim OrgR As Byte, OrgG As Byte, OrgB As Byte
Dim NewR As Byte, NewG As Byte, NewB As Byte
Dim SR As Byte, SG As Byte, SB As Byte
Dim BR As Byte, BG As Byte, BB As Byte
Dim Salpha As Single
Dim salpham As Single

Dim SetAlpha() As Byte
Dim AlphaIncr As Single
Dim DATACULSTO() As Byte
Dim DATACULBU() As Byte

   ReDim DATACULBU(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   Select Case ImageNum
   Case 0
      FILL3D DATACULBU(), DATACUL0()
   Case 1
      FILL3D DATACULBU(), DATACUL1()
   Case 2
      FILL3D DATACULBU(), DATACUL2()
   End Select
   
   ReDim DATACULSTO(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   Select Case ImageNum
   Case 0
      FILL3D DATACULSTO(), DATACUL0()
   Case 1
      FILL3D DATACULSTO(), DATACUL1()
   Case 2
      FILL3D DATACULSTO(), DATACUL2()
   End Select

   ReDim SetAlpha(0 To GradSteps - 1)
   
   SetAlpha(0) = AlphaValue
   If AlphaValue = 255 Then SetAlpha(0) = 254
   AlphaIncr = SetAlpha(0) / GradSteps
   For k = 1 To GradSteps - 1
      If AlphaIncr < SetAlpha(k - 1) Then
         SetAlpha(k) = SetAlpha(k - 1) - AlphaIncr
      Else
         SetAlpha(k) = 1
      End If
      If SetAlpha(k) <= 0 Then SetAlpha(k) = 1
   Next k
   k = k - 1
   
   LngToRGB LabSetCul.BackColor, BR, BG, BB
   
   LngToRGB TColorBGR, SR, SG, SB
   
   ReDim pMaskData(0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   Dim BMIH As BITMAPINFOHEADER
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = ImageWidth(ImageNum)
      .biHeight = ImageHeight(ImageNum)
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With
   
   
   SaveTemp  ' Save DATACUL#() in tempDATACUL()
   
   
   '  0 1 0
   '  1 0 1
   '  0 1 0

   For k = 0 To GradSteps - 1
   
      NPoints = 0
      ReDim six(0 To 100), siy(0 To 100)
   
      For iy = 0 To ImageHeight(ImageNum) - 1
      For ix = 0 To ImageWidth(ImageNum) - 1
         aT = False
         ABYTE = DATACULSTO(3, ix, iy)
         
         If ABYTE = 0 Then ' Cross seems better
            
'            ab(0) = 0
'            ab(1) = 0
'            ab(2) = 0
'            ab(3) = 0
            
            If iy > 0 Then ab(0) = DATACULSTO(3, ix, iy - 1)
            If iy < ImageHeight(ImageNum) - 1 Then ab(3) = DATACULSTO(3, ix, iy + 1)
            
            If ix > 0 Then ab(1) = DATACULSTO(3, ix - 1, iy)
            If ix < ImageWidth(ImageNum) - 1 Then ab(2) = DATACULSTO(3, ix + 1, iy)
            
            For j = 0 To 3
               If ab(j) <> 0 Then
                  aT = True
                  Exit For
               End If
            Next j
            
            If aT Then
               If NPoints > 0 Then
                  If ix <> six(NPoints - 1) Or iy <> siy(NPoints - 1) Then
                     six(NPoints) = ix
                     siy(NPoints) = ImageHeight(ImageNum) - 1 - iy
                     NPoints = NPoints + 1
                  End If
               Else
                  six(0) = ix
                  siy(0) = ImageHeight(ImageNum) - 1 - iy
                  NPoints = NPoints + 1
               End If
               If NPoints > UBound(six()) Then
                  ReDim Preserve six(NPoints + 100)
                  ReDim Preserve siy(NPoints + 100)
               End If
               
               AlphaValue = SetAlpha(k)
               CulValue = RGB(255 - SetAlpha(k), 255 - SetAlpha(k), 255 - SetAlpha(k))
               Select Case ImageNum
               Case 0
                  DATACUL0(3, ix, iy) = SetAlpha(k)
                  DATACUL0(0, ix, iy) = BB
                  DATACUL0(1, ix, iy) = BG
                  DATACUL0(2, ix, iy) = BR
               Case 1
                  DATACUL1(3, ix, iy) = SetAlpha(k)
                  DATACUL1(0, ix, iy) = BB
                  DATACUL1(1, ix, iy) = BG
                  DATACUL1(2, ix, iy) = BR
               Case 2
                  DATACUL2(3, ix, iy) = SetAlpha(k)
                  DATACUL2(0, ix, iy) = BB
                  DATACUL2(1, ix, iy) = BG
                  DATACUL2(2, ix, iy) = BR
               End Select
                     
               SetPixelV picSmallMask.hdc, ix, ImageHeight(ImageNum) - 1 - iy, vbBlack
               SetPixelV picSmallColors.hdc, ix, ImageHeight(ImageNum) - 1 - iy, RGB(BR, BG, BB)
               
            End If
         End If
      Next ix
      Next iy
      
      IACMNum = 0
      aMDown = True
      ATool = 0   ' Dot
      picEDIT_MouseUp 1, 0, 0, 0

      Select Case ImageNum
      Case 0
         FILL3D DATACULSTO(), DATACUL0()
      Case 1
         FILL3D DATACULSTO(), DATACUL1()
      Case 2
         FILL3D DATACULSTO(), DATACUL2()
      End Select
   
   Next k
   
   If GetDIBits(picSmallMask.hdc, picSmallMask.Image, 0, ImageHeight(ImageNum), _
      pMaskData(0, 0), BMIH, 0) = 0 Then
      MsgBox "DIB ERROR", vbCritical, "Getting Mask Data"
      Exit Sub
   End If
      
   SaveTemp  ' Save DATACUL#() in tempDATACUL()
   LoadLast  ' Restore DATACUL#() from tempDATACUL()
   ' Restore tempDATACUL()
   FILL3D tempDATACUL(), DATACULBU()   ' dest, src
   cmdUNDO.Enabled = True  ' Enable AlphaGrad to be undone
   
   IACMNum = -1
   ShowIACM
   GRIDDER

End Sub

Private Sub cmdFillTAreas_Click()
Dim ix As Long, iy As Long
Dim Trans As Long
Dim mix As Long, miy As Long

   SaveTemp
   
   For iy = 0 To ImageHeight(ImageNum) - 1
   For ix = 0 To ImageWidth(ImageNum) - 1
      Trans = 0
      Select Case ImageNum
      Case 0
         If DATACUL0(3, ix, ImageHeight(ImageNum) - 1 - iy) = 0 Then Trans = 1
      Case 1
         If DATACUL1(3, ix, ImageHeight(ImageNum) - 1 - iy) = 0 Then Trans = 1
      Case 2
         If DATACUL2(3, ix, ImageHeight(ImageNum) - 1 - iy) = 0 Then Trans = 1
      End Select
      If Trans = 1 Then
         CulTo32bppArrays OColor, ix, iy, 0 'AlphaValue=0
         SetPixelV picSmallColors.hdc, ix, iy, OColor
         miy = iy * GridMult + 1
         mix = ix * GridMult + 1
         picEDIT.Line (mix, miy)-(mix + GridMult - 2, miy + GridMult - 2), OColor, BF
      End If
   Next ix
   Next iy
   picSmallColors.Picture = picSmallColors.Image
   cmdUNDO.Enabled = True
End Sub

Private Sub cmdReset_Click()
Dim svIACMNum As Long
   svIACMNum = IACMNum
   Select Case ImageNum
   Case 0
         FILL3D DATACUL0(), svDATACUL()
   Case 1
         FILL3D DATACUL1(), svDATACUL()
   Case 2
         FILL3D DATACUL2(), svDATACUL()
   End Select
   AlphaValue = 255
   CulValue = 0
   LabAlpha = AlphaValue
   Initialize
   cmdUNDO.Enabled = False
   IACMNum = svIACMNum - 1
   ShowIACM
End Sub


Private Sub cmdUNDO_Click()
   LoadLast
End Sub

Private Sub LabClick_Click()
   Select Case IACMNum
   Case 0   ' Image
      LabClick.Caption = "Edit Alpha"
      cmdFillTAreas.Enabled = False
   Case 1   ' Alpha
      LabClick.Caption = "Edit Alpha"
      cmdFillTAreas.Enabled = False
   Case 2   ' Original colors
      LabClick.Caption = "Edit colors, Alpha = 0 outside Mask"
      cmdFillTAreas.Enabled = True
      LabAlpha = AlphaValue
   Case 3   ' Mask
      LabClick.Caption = "No editing"
      cmdFillTAreas.Enabled = False
   End Select
End Sub


'Private Sub ShowIACM()
Private Sub ShowIACM()
   IACMNum = IACMNum + 1
   If IACMNum > 3 Then IACMNum = 0
   Select Case IACMNum
   Case 0  ' Image
      TransferSmallToLarge picSmallEdit, picEDIT
      GRIDDER
      'aIACM = True
      Label3(0).ForeColor = vbRed
      Label3(1).ForeColor = 0
      Label3(2).ForeColor = 0
      Label3(3).ForeColor = 0
      'cmdToggleIACM.Caption = "&Show Alpha"
      GreyPic
      LabSetCul.Visible = False
      LabLColor(0).Visible = False
      LabLColor(1).Visible = False
      Label2(0).Visible = True
      Label2(1) = "Transparent Alpha"
      Label2(1).Visible = True
      Label2(3).Visible = False
      Label2(4).Visible = False
      
      Label2(2).Visible = True
      Image1(1).Visible = True
   Case 1   ' Alpha
      TransferSmallToLarge picSmallAlpha, picEDIT
      GRIDDER
      aMDown = False
      Label3(0).ForeColor = 0
      Label3(1).ForeColor = vbRed
      Label3(2).ForeColor = 0
      Label3(3).ForeColor = 0
      GreyPic
      LabSetCul.Visible = False
      LabLColor(0).Visible = False
      LabLColor(1).Visible = False
      Label2(0).Visible = True
      Label2(1) = "Transparent Alpha"
      Label2(1).Visible = True
      Label2(3).Visible = False
      Label2(4).Visible = False
   
      Label2(2).Visible = True
      Image1(1).Visible = True
   Case 2   ' Original colors
      TransferSmallToLarge picSmallColors, picEDIT
      GRIDDER
      Label3(0).ForeColor = 0
      Label3(1).ForeColor = 0
      Label3(2).ForeColor = vbRed
      Label3(3).ForeColor = 0
      
      ' Show palette as main window
      Select Case PALIndex
      Case 0: QBColors picGrey(1), 3
      Case 1: ShortBandedPAL picGrey(1), 1
      Case 2: LongBandedPAL picGrey(1), 1
      Case 3: GreyPAL picGrey(1), 1
      Case 4: CenteredPAL picGrey(1), 1
      End Select
      
      LabSetCul.Visible = True
      LabLColor(0).Visible = True
      LabLColor(1).Visible = True
      Label2(0).Visible = False
      Label2(1) = "   Set     Colors"
      Label2(1).Visible = True
      Label2(3).Visible = True
      Label2(4).Visible = True
      
      Label2(2).Visible = False
      Image1(1).Visible = False
   Case 3   ' Mask
      TransferSmallToLarge picSmallMask, picEDIT
      GRIDDER
      aMDown = False
      LabSetCul.Visible = False
      Label3(0).ForeColor = 0
      Label3(1).ForeColor = 0
      Label3(2).ForeColor = 0
      Label3(3).ForeColor = vbRed
      GreyPic
      LabLColor(0).Visible = False
      LabLColor(1).Visible = False
      Label2(0).Visible = False
      Label2(1).Visible = False
      Label2(3).Visible = False
      Label2(4).Visible = False
   
      Label2(2).Visible = True
      Image1(1).Visible = True
   End Select
   LabClick_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' IACM 0123
   Select Case KeyCode
   Case vbKeyI ' Image
      IACMNum = -1
      ShowIACM
   Case vbKeyA
      IACMNum = 0
      ShowIACM
   Case vbKeyO
      IACMNum = 1
      ShowIACM
   Case vbKeyM
      IACMNum = 2
      ShowIACM
   Case vbKeyS
      ShowIACM
   End Select
End Sub

Private Sub LabLColor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case Index
   Case 0
      OColor = LabLColor(0).BackColor
   Case 1
      OColor = LabLColor(1).BackColor
   End Select
   LabSetCul.BackColor = OColor
End Sub

Private Sub optTool_Click(Index As Integer)
   ATool = Index
   Select Case Index
   Case 0: LabTool = "Dot"
   Case 1: LabTool = "Line"
   Case 2: LabTool = "Box"
   Case 3: LabTool = "SBox"
   End Select
End Sub

Private Sub picSmallAlpha_Click()
IACMNum = 0
ShowIACM
End Sub

Private Sub picSmallColors_Click()
IACMNum = 1
ShowIACM
End Sub

Private Sub picSmallEdit_Click()
IACMNum = -1
ShowIACM
End Sub

Private Sub picSmallMask_Click()
IACMNum = 2
ShowIACM
End Sub

Private Sub Form_Load()
   
   If EditCreate = 0 Then
      Caption = "Alpha editing for Image" & Str$(ImageNum + 1)
      cmdAlphaGrad.Enabled = False
      cboSteps.Enabled = False
   Else
      Caption = "Creating Alpha for Image" & Str$(ImageNum + 1)
      cmdAlphaGrad.Enabled = True
      cboSteps.Enabled = True
   End If
   Left = 0
   Top = 90
   
   Line1(0).Visible = False
   Line1(1).Visible = False
   Box1(0).Visible = False
   Box1(1).Visible = False
   
   ReDim svDATACUL(0 To 3, ImageWidth(ImageNum) - 1, ImageHeight(ImageNum) - 1)
   ReDim tempDATACUL(0 To 3, ImageWidth(ImageNum) - 1, ImageHeight(ImageNum) - 1)
   
   Select Case ImageNum
   Case 0
      FILL3D svDATACUL(), DATACUL0()
   Case 1
      FILL3D svDATACUL(), DATACUL1()
   Case 2
      FILL3D svDATACUL(), DATACUL2()
   End Select
   
   Initialize
   cmdUNDO.Enabled = False

   frmAlphaEdit.KeyPreview = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not aAlphaEdit Then
      Select Case ImageNum
      Case 0
         FILL3D DATACUL0(), svDATACUL()
      Case 1
         FILL3D DATACUL1(), svDATACUL()
      Case 2
         FILL3D DATACUL2(), svDATACUL()
      End Select
      aAlphaEdit = False
      Erase six(), siy()
   End If
End Sub

Private Sub picEDIT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k As Long
Dim ix As Long, iy As Long
Dim xleft As Single, xright As Single
Dim ytop As Single, ybelow As Single
Dim Cul As Long
Dim pr As Byte, pg As Byte, pb As Byte
Dim CulM As Long
   
   NPoints = 0
   ReDim six(100), siy(100)
   ix = X \ GridMult
   If ix < 0 Then ix = 0
   If ix > ImageWidth(ImageNum) - 1 Then ix = ImageWidth(ImageNum) - 1
   iy = Y \ GridMult
   If iy < 0 Then iy = 0
   If iy > ImageHeight(ImageNum) - 1 Then iy = ImageHeight(ImageNum) - 1
   LabXY = "X =" & Str$(ix) & ", Y =" & Str$(iy)
   LabXY.Refresh
   
   If IACMNum = 3 Then Exit Sub   ' Mask
   
   If Button = vbRightButton Then
      If IACMNum = 2 Then  ' OC
         Cul = picEDIT.Point(X, Y)
         If Cul >= 0 And Cul <> GridCol Then
            LabLColor(0).BackColor = Cul
            Form1.LabColor(0).BackColor = Cul
            LColor = Cul
         End If
      ElseIf IACMNum = 0 Or IACMNum = 1 Then  ' Image or Alpha
         Cul = picSmallAlpha.Point(ix, iy)
         If Cul >= 0 And Cul <> GridCol Then
            LngToRGB Cul, pr, pg, pb
            Cul = 255 - pr
            LabCul.BackColor = RGB(pr, pr, pr)
            LabVal.Caption = Str$(Cul)
            LabRGB(0) = pr
            LabRGB(1) = pg
            LabRGB(2) = pb
            AlphaValue = 255 - pr
            LabAlpha = 255 - pr
         End If
      End If
      Exit Sub
   End If
   
   aMDown = True
   six(0) = ix
   siy(0) = iy
   NPoints = NPoints + 1
   
   xleft = (ix) * GridMult + 1
   xright = xleft + GridMult - 2
   ytop = (iy) * GridMult + 1
   ybelow = ytop + GridMult - 2
   
   ' AlphaValue=set Alpha value 0 to 255
   CulValue = RGB(255 - AlphaValue, 255 - AlphaValue, 255 - AlphaValue)
   CulM = picSmallMask.Point(ix, iy)
   If CulM <> vbWhite Then  ' Inside Mask
      SaveTemp  ' Save DATACUL#() in tempDATACUL()
      Select Case ATool
      Case 0
         picEDIT.Line (xleft, ytop)-(xright, ybelow), vbRed, B
      Case 1
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
      Case 2, 3  ' Box, Solid box
         xStart = X
         yStart = Y
         xend = X
         yend = Y
         StartOnLine
         For k = 0 To 1
            Box1(k).Visible = True
            Box1(k).Move xStart, yStart, 0, 0
         Next k
      
      End Select
   Else  ' Outside Mask  White
      If IACMNum = 2 Then ' OC Outside Mask
         SaveTemp  ' Save DATACUL#() in tempDATACUL()
         Select Case ATool
         Case 0
            picEDIT.Line (xleft, ytop)-(xright, ybelow), vbRed, B
         Case 1
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
         Case 2, 3  ' Box, Solid box
            xStart = X
            yStart = Y
            xend = X
            yend = Y
            StartOnLine
            For k = 0 To 1
               Box1(k).Visible = True
               Box1(k).Move xStart, yStart, 0, 0
            Next k
         End Select
      Else
         NPoints = 0
      End If
   End If

End Sub

Private Sub picEDIT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ix As Long, iy As Long
Dim xleft As Single, xright As Single
Dim ytop As Single, ybelow As Single
Dim Cul As Long
Dim pr As Byte, pg As Byte, pb As Byte
Dim CulM As Long
   ix = X \ GridMult
   If ix < 0 Then ix = 0
   If ix > ImageWidth(ImageNum) - 1 Then ix = ImageWidth(ImageNum) - 1
   iy = Y \ GridMult
   If iy < 0 Then iy = 0
   If iy > ImageHeight(ImageNum) - 1 Then iy = ImageHeight(ImageNum) - 1
   LabXY = "X =" & Str$(ix) & ", Y =" & Str$(iy)
   LabXY.Refresh
   Cul = picEDIT.Point(X, Y)
   If Cul <> GridCol Then
      If Cul < 0 Then Cul = 0
      LngToRGB Cul, pr, pg, pb
      LabRGB(0) = pr
      LabRGB(1) = pg
      LabRGB(2) = pb
      LabCul.BackColor = Cul
   End If
   
   If aMDown Then
      
      If NPoints > 0 Then
         If ix <> six(NPoints - 1) Or iy <> siy(NPoints - 1) Then
            six(NPoints) = ix
            siy(NPoints) = iy
            NPoints = NPoints + 1
         End If
      End If
      
      If NPoints > UBound(six()) Then
         ReDim Preserve six(NPoints + 100)
         ReDim Preserve siy(NPoints + 100)
      End If
      
      xleft = (ix) * GridMult + 1
      xright = xleft + GridMult - 2
      ytop = (iy) * GridMult + 1
      ybelow = ytop + GridMult - 2
   
      CulM = picSmallMask.Point(ix, iy)
      If CulM <> vbWhite Then   'Black inside mask
         Select Case ATool
         Case 0   ' Dot
            picEDIT.Line (xleft, ytop)-(xright, ybelow), vbRed, B
         Case 1  ' Line
            xend = X
            yend = Y
            ' Draw shape line
            Line1(0).x2 = xend
            Line1(0).y2 = yend
            Line1(1).x2 = xend
            Line1(1).y2 = yend
         Case 2, 3  ' Box, Solid box
            xend = X
            yend = Y
            DrawBox1 frmAlphaEdit, picEDIT, X, Y, 1
         End Select
      Else  ' Outside Mask  White
         If IACMNum = 2 Then  ' OC Outside Mask
            Select Case ATool
            Case 0
               picEDIT.Line (xleft, ytop)-(xright, ybelow), vbRed, B
            Case 1
               xend = X
               yend = Y
               ' Draw shape line
               Line1(0).x2 = xend
               Line1(0).y2 = yend
               Line1(1).x2 = xend
               Line1(1).y2 = yend
            Case 2, 3  ' Box, Solid box
               xend = X
               yend = Y
               DrawBox1 frmAlphaEdit, picEDIT, X, Y, 1
            Case 4   ' Fill Alpha
            End Select
         End If
      End If
   
   Else   ' Mouse Up just Mouse Move
      
      Cul = picSmallAlpha.Point(ix, iy)
         If Cul = GridCol Then Exit Sub
         
         If Cul >= 0 Then
            LngToRGB Cul, pr, pg, pb
            Cul = 255 - pr
            If IACMNum <> 2 Then ' ie Mask
               LabCul.BackColor = RGB(pr, pr, pr)
            End If
            LabVal.Caption = Str$(Cul)
         End If
   
   End If   ' If aMDown Then
End Sub

Private Sub picEDIT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ix As Long, iy As Long
Dim iyy As Long

'' Line
'Dim xxs As Single, yys As Single
'Dim xxe As Single, yye As Single
'' Box
'Dim bxxs As Single
'Dim bxxe As Single
'Dim byys As Single
'Dim byye As Single


Dim OrgR As Byte, OrgG As Byte, OrgB As Byte
Dim NewR As Byte, NewG As Byte, NewB As Byte
Dim SR As Byte, SG As Byte, SB As Byte
Dim ABYTE As Byte
Dim Salpha As Single
Dim salpham As Single

Dim BR As Byte, BG As Byte, BB As Byte
Dim pr As Byte, pg As Byte, pb As Byte
Dim xleft As Single, xright As Single
Dim ytop As Single, ybelow As Single
Dim NP As Long
Dim CulM As Long

   If IACMNum = 3 Then Exit Sub   ' Mask
   
   If aMDown Then
      
      If ATool <> 0 Then  ' Line, Box or Solid Box
         GRID_PIC_Coords X, Y
      End If
      
      LngToRGB TColorBGR, BR, BG, BB
      LngToRGB TColorRGB, SR, SG, SB
      Salpha = AlphaValue / 255
      salpham = 1 - Salpha
      
      
      '======================================================================
      If IACMNum <> 2 Then  ' Image or Alpha
         
         For NP = 0 To NPoints - 1
            ix = six(NP)
            iy = siy(NP)
            
            iyy = ImageWidth(ImageNum) - iy - 1
            CulM = picSmallMask.Point(ix, iy)
            If CulM = vbBlack Then  ' Inside Mask
               '===========================================
               ' Show changed alphas
               SetPixelV picSmallAlpha.hdc, ix, iy, CulValue
              '===========================================
               ' Get original colors
               Select Case ImageNum
               Case 0
                  ABYTE = DATACUL0(3, ix, ImageHeight(ImageNum) - 1 - iy)
                  OrgB = DATACUL0(0, ix, ImageHeight(ImageNum) - 1 - iy)
                  OrgG = DATACUL0(1, ix, ImageHeight(ImageNum) - 1 - iy)
                  OrgR = DATACUL0(2, ix, ImageHeight(ImageNum) - 1 - iy)
               Case 1
                  ABYTE = DATACUL1(3, ix, ImageHeight(ImageNum) - 1 - iy)
                  OrgB = DATACUL1(0, ix, ImageHeight(ImageNum) - 1 - iy)
                  OrgG = DATACUL1(1, ix, ImageHeight(ImageNum) - 1 - iy)
                  OrgR = DATACUL1(2, ix, ImageHeight(ImageNum) - 1 - iy)
               Case 2
                  ABYTE = DATACUL2(3, ix, ImageHeight(ImageNum) - 1 - iy)
                  OrgB = DATACUL2(0, ix, ImageHeight(ImageNum) - 1 - iy)
                  OrgG = DATACUL2(1, ix, ImageHeight(ImageNum) - 1 - iy)
                  OrgR = DATACUL2(2, ix, ImageHeight(ImageNum) - 1 - iy)
               End Select
   '            '===========================================
   '            ' Re-convolve with new alpha
               NewR = SR * salpham + OrgR * Salpha
               NewG = SG * salpham + OrgG * Salpha
               NewB = SB * salpham + OrgB * Salpha
            
               If AlphaValue = 0 Then    'salpha=0, salpham=1, make transparent BGR = 197,195,194
                  SetPixelV picSmallEdit.hdc, ix, iy, RGB(NewB, NewG, NewR)
                  ' Also make mask transparent
                  SetPixelV picSmallMask.hdc, ix, iy, vbWhite
                  pMaskData(ix, ImageHeight(ImageNum) - 1 - iy) = vbWhite
               Else  ' AlphaValue>0
                  If NewR = SR And NewG = SG And NewB = SB Then
                     SetPixelV picSmallEdit.hdc, ix, iy, RGB(SR + 1, SG - 1, SB)
                  Else
                     SetPixelV picSmallEdit.hdc, ix, iy, RGB(NewR, NewG, NewB)
                  End If
                  ' Make mask opaque
                  SetPixelV picSmallMask.hdc, ix, iy, 0
                  pMaskData(ix, ImageHeight(ImageNum) - 1 - iy) = 0
               End If
               ' Overwrite red sq with Culvalue
               xleft = (ix) * GridMult + 1
               xright = xleft + GridMult - 2
               ytop = (iy) * GridMult + 1
               ybelow = ytop + GridMult - 2
               picEDIT.Line (xleft, ytop)-(xright, ybelow), CulValue, BF
               '===========================================
               ' Set new Alphas
               Select Case ImageNum
               Case 0: DATACUL0(3, ix, ImageHeight(ImageNum) - iy - 1) = AlphaValue
               Case 1: DATACUL1(3, ix, ImageHeight(ImageNum) - iy - 1) = AlphaValue
               Case 2: DATACUL2(3, ix, ImageHeight(ImageNum) - iy - 1) = AlphaValue
               End Select
               '===========================================
            Else  ' OUtside Mask, Alpha =0
               ' When Mask = transparent, cannot change anything
               ' except in main image by drawing with a color
               ' or transparent color
               pMaskData(ix, ImageHeight(ImageNum) - 1 - iy) = vbWhite
            End If
         
         Next NP
      '======================================================================
         
         picSmallAlpha.Picture = picSmallAlpha.Image
         picSmallEdit.Picture = picSmallEdit.Image
         picSmallMask.Picture = picSmallMask.Image
         picSmallColors.Picture = picSmallColors.Image
         If IACMNum = 0 Then
            TransferSmallToLarge picSmallEdit, picEDIT
         End If
         
         GRIDDER
      
      Else   ' Original colors
            For NP = 0 To NPoints - 1
               ix = six(NP)
               iy = siy(NP)
               CulValue = picSmallAlpha.Point(ix, iy)
               LngToRGB CulValue, pr, pg, pb
               AlphaValue = 255 - pr
               Salpha = AlphaValue / 255
               salpham = 1 - Salpha
               If picSmallMask.Point(ix, iy) <> vbWhite Then   ' OC Inside Mask
               
                  CulTo32bppArrays OColor, ix, iy, AlphaValue
                  LngToRGB OColor, pr, pg, pb
                  NewR = SR * salpham + pr * Salpha
                  NewG = SG * salpham + pg * Salpha
                  NewB = SB * salpham + pb * Salpha
                  SetPixelV picSmallEdit.hdc, ix, iy, RGB(NewR, NewG, NewB) 'OColor
                  
                  SetPixelV picSmallAlpha.hdc, ix, iy, CulValue
                  SetPixelV picSmallColors.hdc, ix, iy, OColor
               
               Else ' OC outside Mask
                  CulTo32bppArrays OColor, ix, iy, 0 'AlphaValue=0
                  SetPixelV picSmallColors.hdc, ix, iy, OColor
               End If

               xleft = (ix) * GridMult + 1
               xright = xleft + GridMult - 2
               ytop = (iy) * GridMult + 1
               ybelow = ytop + GridMult - 2
               picEDIT.Line (xleft, ytop)-(xright, ybelow), OColor, BF
            
            Next NP
            
            picSmallColors.Picture = picSmallColors.Image
            picSmallEdit.Picture = picSmallEdit.Image
      End If
   
   End If
   
   aMDown = False

End Sub

' Grey Palette  0-255 or Color palette
Private Sub picGrey_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
Dim pr As Byte, pg As Byte, pb As Byte

   Cul = picGrey(Index).Point(X, Y)
   If Cul >= 0 Then
      LngToRGB Cul, pr, pg, pb
      If IACMNum = 2 Then  ' Original colors
         LabCul.BackColor = Cul
      Else
         Cul = 255 - pr
         LabCul.BackColor = RGB(pr, pr, pr)
         LabVal.Caption = Str$(Cul)
      End If
      LabRGB(0) = pr
      LabRGB(1) = pg
      LabRGB(2) = pb
   End If
End Sub

Private Sub picGrey_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
Dim pr As Byte, pg As Byte, pb As Byte

   Cul = picGrey(Index).Point(X, Y)
   If Cul >= 0 Then
      If IACMNum = 2 Then   ' Original colors
         OColor = Cul   'A new originalcolor
         LabCul.BackColor = Cul
         LabSetCul.BackColor = Cul
      Else
         LngToRGB Cul, pr, pg, pb
         AlphaValue = 255 - pr
         LabAlpha = 255 - pr
      End If
   End If
End Sub


Private Sub Initialize()
Dim k As Long
   ' Orignal colors in DATACULSRC#(0 to 3,W-1,H-1)
   
   AlphaValue = 255
   aAlphaEdit = False

   GridCol = RGB(1, 128, 3)
   
   Border_Locate
   
   cboSteps.Clear
   For k = 2 To 6
      cboSteps.AddItem LTrim$(Str$(k))
   Next k
   cboSteps.ListIndex = 0
   GradSteps = 2
   
   ' Display Image
   picSmallEdit.BackColor = TColorRGB
   ' Dest  <- Src
   ' Display picSmallEdit from Form1.picSmall(ImageNum)
   BitBlt picSmallEdit.hdc, 0, 0, Form1.picSmall(ImageNum).Width, Form1.picSmall(ImageNum).Height, _
          Form1.picSmall(ImageNum).hdc, 0, 0, vbSrcCopy
   picSmallEdit.Picture = picSmallEdit.Image
   
   ' Display Alphas
   ShowAlpha ImageNum, picSmallAlpha, picSmallAlpha
   picSmallAlpha.Picture = picSmallAlpha.Image
   
   ' Display mask  needs picSmallEdit
   ShowMask picSmallEdit, picSmallMask 'Src, Dest
   picSmallMask.Picture = picSmallMask.Image
   
   VisColor = 0 'vbRed 'TColorBGR
   'vbWhite
   ' Show Original colors with no modification
   If ImBPP(ImageNum) <> 32 Then
      ShowWithTColor ImageNum, Form1.picSmall(ImageNum), picSmallColors, 1
   Else
      ShowWithTColor ImageNum, picSmallColors, picSmallColors, 1
   End If
   picSmallColors.Picture = picSmallColors.Image
   
   
   OColor = 0  ' Default Original color
   LabSetCul.BackColor = 0
   LabLColor(0).BackColor = LColor
   LabLColor(1).BackColor = RColor
   

   ' Get pMaskData() from picSmallMask picture box
   ReDim pMaskData(0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   Dim BMIH As BITMAPINFOHEADER
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = ImageWidth(ImageNum)
      .biHeight = ImageHeight(ImageNum)
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With
   If GetDIBits(picSmallMask.hdc, picSmallMask.Image, 0, ImageHeight(ImageNum), _
      pMaskData(0, 0), BMIH, 0) = 0 Then
      MsgBox "DIB ERROR", vbCritical, "Getting Mask Data"
      Exit Sub
   End If
   
   optTool(0).Value = True
   optTool_Click 0
   
   
   Line1(0).Visible = False
   Line1(1).Visible = False
   Box1(0).Visible = False
   Box1(1).Visible = False
   
   IACMNum = -1
   ShowIACM
   GRIDDER
End Sub

Private Sub Border_Locate()
   picSmallEdit.Height = ImageHeight(ImageNum)
   picSmallEdit.Width = ImageWidth(ImageNum)
   Line (picSmallEdit.Left - 1, picSmallEdit.Top - 1)- _
   (picSmallEdit.Left + picSmallEdit.Width, picSmallEdit.Top + picSmallEdit.Height), GridCol, B
   
   picSmallAlpha.Height = ImageHeight(ImageNum)
   picSmallAlpha.Width = ImageWidth(ImageNum)
   Line (picSmallAlpha.Left - 1, picSmallAlpha.Top - 1)- _
   (picSmallAlpha.Left + picSmallAlpha.Width, picSmallAlpha.Top + picSmallAlpha.Height), GridCol, B
   
   picSmallMask.Height = ImageHeight(ImageNum)
   picSmallMask.Width = ImageWidth(ImageNum)
   Line (picSmallMask.Left - 1, picSmallMask.Top - 1)- _
   (picSmallMask.Left + picSmallMask.Width, picSmallMask.Top + picSmallMask.Height), GridCol, B
   
   picSmallColors.Height = ImageHeight(ImageNum)
   picSmallColors.Width = ImageWidth(ImageNum)
   Line (picSmallColors.Left - 1, picSmallColors.Top - 1)- _
   (picSmallColors.Left + picSmallColors.Width, picSmallColors.Top + picSmallColors.Height), GridCol, B
   
   
   picEDIT.Height = ImageHeight(ImageNum) * GridMult
   picEDIT.Width = ImageWidth(ImageNum) * GridMult
   picEDIT.Left = (picEditC.Width - picEDIT.Width) \ 2 - 2
   picEDIT.Top = (picEditC.Height - picEDIT.Height) \ 2 - 2
   picEDIT.Picture = picEDIT.Image

End Sub

Private Sub GRIDDER()
Dim k As Long
Dim cx1 As Single, cx2 As Single
Dim cy1 As Single, cy2 As Single
   ' GRID
      ' Horz lines
   For k = 0 To GridMult * ImageHeight(ImageNum) Step GridMult
      picEDIT.Line (0, k)-(picEDIT.Width, k), GridCol
   Next k
   ' Vert lines
   For k = 0 To GridMult * ImageWidth(ImageNum) Step GridMult
      picEDIT.Line (k, 0)-(k, picEDIT.Height), GridCol
   Next k
   ' Bottom & Right lines to complete grid appearance
   cx1 = picEDIT.Left
   cx2 = cx1 + picEDIT.Width
   cy1 = picEDIT.Top
   cy2 = cy1 + picEDIT.Height
   picEditC.Line (cx2, cy1)-(cx2, cy2), GridCol  ' Vert
   picEditC.Line (cx1, cy2)-(cx2 + 1, cy2), GridCol  ' Horz
   
   JustGrid

End Sub

Private Sub JustGrid()
Dim ix As Long, iy As Long
Dim mix As Long, miy As Long
Dim Cul As Long
      For iy = 0 To ImageHeight(ImageNum) - 1
      miy = iy * GridMult + 1
      For ix = 0 To ImageWidth(ImageNum) - 1
         If IACMNum < 2 Then  ' Image or Alpha
            Cul = pMaskData(ix, ImageHeight(ImageNum) - iy - 1) And &HFFFFFF
            If Cul >= 0 Then
               If Cul = vbWhite Then   ' outside Mask
                  mix = ix * GridMult + 1
                  'sq for transparency
                  picEDIT.Line (mix + 1, miy + 1)-(mix + GridMult - 3, miy + GridMult - 3), GridCol, B
               End If
            End If
         Else  ' IACMNum=2 Original colors 0r = 3 Mask
            Cul = pMaskData(ix, ImageHeight(ImageNum) - iy - 1) And &HFFFFFF
            If Cul >= 0 Then
               If IACMNum = 3 Then ' Mask
                  If Cul = vbWhite Then   ' Outside mask
                     mix = ix * GridMult + 1
                     picEDIT.Line (mix, miy)-(mix + GridMult - 2, miy + GridMult - 2), TColorBGR, BF
                     'sq for transparency
                     picEDIT.Line (mix + 1, miy + 1)-(mix + GridMult - 3, miy + GridMult - 3), GridCol, B
                  End If
               Else  ' OC
               End If
            End If
         End If
      Next ix
      Next iy
End Sub

Private Sub GreyPic()
Dim k As Long
   ' Grey Palette
   picGrey(1).Picture = LoadPicture
   For k = 0 To 255
      picGrey(1).Line (0, 255 - k)-(picGrey(1).Width, 255 - k), RGB(k, k, k)
   Next k
   k = k - 1
   picGrey(1).Picture = picGrey(1).Image
End Sub

Private Sub GRID_PIC_Coords(X As Single, Y As Single)
' Private ATool
' Line
Dim xxs As Single, yys As Single
Dim xxe As Single, yye As Single
' Box
Dim bxxs As Single
Dim bxxe As Single
Dim byys As Single
Dim byye As Single
   
Dim k As Long
   
   Select Case ATool
   Case 1   'Line
      Line1(0).Visible = False
      Line1(1).Visible = False
      xend = X
      yend = Y
      ' Convert GRID coords to picSmall coords
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
      
      ReDim six(100), siy(100)
      NPoints = 0
      BresLine xxs, yys, xxe, yye  ' Set six(), siy() & NPoints for Line tool
   Case 2 ' Box
      Box1(0).Visible = False
      Box1(1).Visible = False
      'xend = X
      'yend = Y
      ReDim six(100), siy(100)
      NPoints = 0
      bxxs = Box1(0).Left \ GridMult
      bxxe = bxxs + Box1(0).Width \ GridMult - 1
      byys = Box1(0).Top \ GridMult
      byye = byys + Box1(0).Height \ GridMult - 1
      ' 4 lines
      BresLine bxxs, byys, bxxe, byys
      BresLine bxxe, byys, bxxe, byye
      BresLine bxxe, byye, bxxs, byye
      BresLine bxxs, byye, bxxs, byys
   Case 3   ' Solid box
      ' byye-byys+1 lines
      Box1(0).Visible = False
      Box1(1).Visible = False
      ReDim six(100), siy(100)
      NPoints = 0
      bxxs = Box1(0).Left \ GridMult
      bxxe = bxxs + Box1(0).Width \ GridMult - 1
      byys = Box1(0).Top \ GridMult
      byye = byys + Box1(0).Height \ GridMult - 1
      For k = 0 To byye - byys
         BresLine bxxs, byys + k, bxxe, byys + k
      Next k
   End Select
End Sub

Private Sub BresLine(ByVal ix1 As Long, ByVal iy1 As Long, _
                     ByVal ix2 As Long, ByVal iy2 As Long)

' BresLine ix1,iy1,ix2,iy2  ' start and end points
' Used for collecting ix,iy values along a lines
' into six(), siy()
Dim iix As Long, iiy As Long
Dim idx As Long, idy As Long
Dim jkstep As Long
Dim incx As Long
Dim id As Long
Dim ainc As Long, binc As Long
Dim jj As Long, kk As Long

   ' Reject lines outside image W & H
   If ix1 >= 0 Or ix2 >= 0 Then
   If ix1 < ImageWidth(ImageNum) Or ix2 < ImageWidth(ImageNum) Then
   If iy1 >= 0 Or iy2 >= 0 Then
   If iy1 < ImageHeight(ImageNum) Or iy2 < ImageHeight(ImageNum) Then
      
      idx = Abs(ix2 - ix1)
      idy = Abs(iy2 - iy1)
      jkstep = 1
      incx = 1
      If idx < idy Then   '-- Steep slope
         
         If iy1 > iy2 Then jkstep = -1
         If ix2 < ix1 Then incx = -1
         id = 2 * idx - idy
         ainc = 2 * (idx - idy)   '-ve
         binc = 2 * idx
         jj = iy1: kk = iy2: iix = ix1
         For iiy = jj To kk Step jkstep
            ' Reject any point outside rect
            If iix >= 0 Then
            If iix < ImageWidth(ImageNum) Then
            If iiy >= 0 Then
            If iiy < ImageHeight(ImageNum) Then
               six(NPoints) = iix
               siy(NPoints) = iiy
               NPoints = NPoints + 1
               If NPoints > UBound(six()) Then
                  ReDim Preserve six(NPoints + 100)
                  ReDim Preserve siy(NPoints + 100)
               End If
            End If
            End If
            End If
            End If
            If id > 0 Then
               id = id + ainc
               iix = iix + incx
            Else
               id = id + binc
            End If
         Next iiy
      
      Else                '-- Shallow slope
         
         If ix1 > ix2 Then jkstep = -1
         If iy2 < iy1 Then incx = -1
         id = 2 * idy - idx
         ainc = 2 * (idy - idx)   '-ve
         binc = 2 * idy
         jj = ix1: kk = ix2: iix = iy1
      
         For iiy = jj To kk Step jkstep
            ' Reject any point outside rect
            ' NB switch in x,y
            If iiy >= 0 Then
            If iiy < ImageWidth(ImageNum) Then
            If iix >= 0 Then
            If iix < ImageHeight(ImageNum) Then
               six(NPoints) = iiy
               siy(NPoints) = iix
               NPoints = NPoints + 1
               If NPoints > UBound(six()) Then
                  ReDim Preserve six(NPoints + 100)
                  ReDim Preserve siy(NPoints + 100)
               End If
            End If
            End If
            End If
            End If
            If id > 0 Then
               id = id + ainc
               iix = iix + incx
            Else
               id = id + binc
            End If
         Next iiy
      
      End If
   
   End If
   End If
   End If
   End If
End Sub

Private Sub SaveTemp()
   Select Case ImageNum
   Case 0
      FILL3D tempDATACUL(), DATACUL0()
   Case 1
      FILL3D tempDATACUL(), DATACUL1()
   Case 2
      FILL3D tempDATACUL(), DATACUL2()
   End Select
   cmdUNDO.Enabled = True
End Sub

Private Sub LoadLast()
Dim ix As Long, iy As Long
Dim OrgR As Byte, OrgG As Byte, OrgB As Byte
Dim NewR As Byte, NewG As Byte, NewB As Byte
Dim SR As Byte, SG As Byte, SB As Byte
Dim ABYTE As Byte
Dim Salpha As Single
Dim salpham As Single

   ' Restore DATACUL#()
   Select Case ImageNum
   Case 0
      FILL3D DATACUL0(), tempDATACUL()
   Case 1
      FILL3D DATACUL1(), tempDATACUL()
   Case 2
      FILL3D DATACUL2(), tempDATACUL()
   End Select

   ' Restore Alphas from DATACUL#()
   ShowAlpha ImageNum, picSmallAlpha, picSmallAlpha
   picSmallAlpha.Picture = picSmallAlpha.Image
   
   
   If ImBPP(ImageNum) = 32 Then
      ' Restore Mask Image & Original colors
      LngToRGB TColorRGB, SR, SG, SB
      
      For iy = 0 To ImageHeight(ImageNum) - 1
      For ix = 0 To ImageWidth(ImageNum) - 1
         
         Select Case ImageNum
         Case 0:
            OrgB = DATACUL0(0, ix, iy)
            OrgG = DATACUL0(1, ix, iy)
            OrgR = DATACUL0(2, ix, iy)
            ABYTE = DATACUL0(3, ix, iy)
         Case 1
            OrgB = DATACUL1(0, ix, iy)
            OrgG = DATACUL1(1, ix, iy)
            OrgR = DATACUL1(2, ix, iy)
            ABYTE = DATACUL1(3, ix, iy)
         Case 2
            OrgB = DATACUL2(0, ix, iy)
            OrgG = DATACUL2(1, ix, iy)
            OrgR = DATACUL2(2, ix, iy)
            ABYTE = DATACUL2(3, ix, iy)
         End Select
         
         ' Restore Mask
         If ABYTE <> 0 Then
               SetPixelV picSmallMask.hdc, ix, ImageHeight(ImageNum) - 1 - iy, vbBlack
               pMaskData(ix, iy) = vbBlack
         Else
               SetPixelV picSmallMask.hdc, ix, ImageHeight(ImageNum) - 1 - iy, vbWhite
               pMaskData(ix, iy) = vbWhite
         End If
         
         If picSmallMask.Point(ix, ImageHeight(ImageNum) - 1 - iy) = vbBlack Then ' Inside Mask
            ' Restore Image
            Salpha = ABYTE / 255
            salpham = 1 - Salpha
            ' Re-convolve with alpha
            NewR = SR * salpham + OrgR * Salpha
            NewG = SG * salpham + OrgG * Salpha
            NewB = SB * salpham + OrgB * Salpha
            
            SetPixelV picSmallEdit.hdc, ix, ImageHeight(ImageNum) - 1 - iy, RGB(NewR, NewG, NewB)
      
         Else  ' Outside mask make transparent
            SetPixelV picSmallEdit.hdc, ix, ImageHeight(ImageNum) - 1 - iy, RGB(SB, SG, SR)
         End If
         ' Restore Original Colors
         SetPixelV picSmallColors.hdc, ix, ImageHeight(ImageNum) - 1 - iy, RGB(OrgR, OrgG, OrgB)
         
      Next ix
      Next iy
   
   Else  ' Incoming < 32bpp ie Create
   End If
   
   'picSmallMask.Picture = picSmallMask.Image
   'picSmallEdit.Picture = picSmallEdit.Image
   'picSmallColors.Picture = picSmallColors.Image
   
   NPoints = 0
   ATool = 0
   optTool_Click 0
   optTool(0).Value = True
   
   cmdUNDO.Enabled = False
   ix = IACMNum
   For iy = 0 To 3
      IACMNum = iy
      ShowIACM
   Next iy
   IACMNum = ix - 1
   ShowIACM
End Sub

