VERSION 5.00
Begin VB.Form frmSelect 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Icon selector"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3840
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdACCCan 
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
      Height          =   300
      Index           =   1
      Left            =   1980
      TabIndex        =   10
      Top             =   4575
      Width           =   1305
   End
   Begin VB.CommandButton cmdACCCan 
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
      Height          =   300
      Index           =   0
      Left            =   435
      TabIndex        =   9
      Top             =   4575
      Width           =   1305
   End
   Begin VB.CommandButton cmdSEL 
      Caption         =   "Select All pictures"
      Height          =   345
      Index           =   3
      Left            =   945
      TabIndex        =   0
      Top             =   3855
      Width           =   1710
   End
   Begin VB.CommandButton cmdSEL 
      Caption         =   "Select picture 3"
      Height          =   330
      Index           =   2
      Left            =   1830
      TabIndex        =   8
      Top             =   2700
      Width           =   1515
   End
   Begin VB.CommandButton cmdSEL 
      Caption         =   "Select picture 2"
      Height          =   330
      Index           =   1
      Left            =   1830
      TabIndex        =   7
      Top             =   1470
      Width           =   1515
   End
   Begin VB.CommandButton cmdSEL 
      Caption         =   "Select picture 1"
      Height          =   330
      Index           =   0
      Left            =   1830
      TabIndex        =   6
      Top             =   300
      Width           =   1515
   End
   Begin VB.PictureBox picSEL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   2
      Left            =   405
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   2640
      Width           =   960
   End
   Begin VB.PictureBox picSEL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   1
      Left            =   405
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   1440
      Width           =   960
   End
   Begin VB.PictureBox picSEL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   0
      Left            =   405
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   11
      Top             =   255
      Width           =   960
   End
   Begin VB.Label LabToIm 
      BackColor       =   &H00404040&
      Caption         =   "starting at Image 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   1185
      TabIndex        =   15
      Top             =   4245
      Width           =   1335
   End
   Begin VB.Label LabToIm 
      BackColor       =   &H00404040&
      Caption         =   "to Image #"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   2085
      TabIndex        =   14
      Top             =   3090
      Width           =   1095
   End
   Begin VB.Label LabToIm 
      BackColor       =   &H00404040&
      Caption         =   "to Image #"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   2085
      TabIndex        =   13
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Label LabToIm 
      BackColor       =   &H00404040&
      Caption         =   "to Image #"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   2085
      TabIndex        =   12
      Top             =   675
      Width           =   930
   End
   Begin VB.Image imSel 
      Height          =   315
      Index           =   3
      Left            =   2760
      Picture         =   "frmSelect.frx":0000
      Top             =   3870
      Width           =   240
   End
   Begin VB.Image imSel 
      Height          =   315
      Index           =   2
      Left            =   3420
      Picture         =   "frmSelect.frx":0186
      Top             =   2700
      Width           =   240
   End
   Begin VB.Image imSel 
      Height          =   315
      Index           =   1
      Left            =   3420
      Picture         =   "frmSelect.frx":030C
      Top             =   1500
      Width           =   240
   End
   Begin VB.Image imSel 
      Height          =   315
      Index           =   0
      Left            =   3420
      Picture         =   "frmSelect.frx":0492
      Top             =   330
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
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
      Height          =   270
      Index           =   2
      Left            =   90
      TabIndex        =   5
      Top             =   2625
      Width           =   225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
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
      Height          =   270
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   1425
      Width           =   225
   End
   Begin VB.Label Label1 
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
      Height          =   270
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   240
      Width           =   225
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdACCCan_Click(Index As Integer)
Dim r As RECT
   If Index = 0 Then  ' Accept
   
   Else  ' Cancel
      Response$ = ""
   End If
  ' Avoid click through
   GetWindowRect Form1.picSmallFrame.hwnd, r
   SetCursorPos r.Left + 130, r.Top + 90
   Unload Me
End Sub

Private Sub cmdSEL_Click(Index As Integer)
Dim k As Long
   For k = 0 To 3
      imSel(k).Visible = False
   Next k
   Select Case Index
   Case 0:  Response$ = "1"
   Case 1:  Response$ = "2"
   Case 2:  Response$ = "3"
   Case 3:  Response$ = "A"
   End Select
   imSel(Index).Visible = True
End Sub

Private Sub Form_Load()
Dim fnum As Integer
Dim k As Long
Dim BStream() As Byte
Dim IcoStream() As Byte

      fnum = FreeFile
      Open FileSpec$ For Binary As #fnum
      ReDim BStream(0 To LOF(fnum) - 1)
      Get #fnum, , BStream
      Close #fnum
      
      If NumIcons = 2 Then
         picSEL(2).Visible = False
         cmdSEL(2).Visible = False
      Else
         picSEL(2).Visible = True
         cmdSEL(2).Visible = True
      End If
      
      For k = 0 To NumIcons - 1
         If ETools2ctICON(k, BStream(), IcoStream()) Then
            Form1.picSmallBU.Picture = LoadPicture
            Set Form1.picSmallBU.Picture = PictureFromByteStream(IcoStream)
            Form1.picSmallBU.Picture = Form1.picSmallBU.Image
            ImageNum = k   'TempImageNum
            TransferImageTofrmSEL  ' One Transfer picSmallBU to picSEL(ImageNum)
         Else
            GoTo SELERROR
         End If
         If k = 2 Then Exit For
      Next k
      Response$ = "A"
      For k = 0 To 2
         imSel(k).Visible = False
         LabToIm(k) = "to Image" & Str$(TempImageNum + 1)
      Next k
      Exit Sub
'==========
SELERROR:
Response$ = ""
Unload Me
End Sub

Private Sub TransferImageTofrmSEL()
   
   With picSEL(ImageNum)
      .Width = Form1.picSmallBU.Width
      .Height = Form1.picSmallBU.Height
   End With
   BitBlt picSEL(ImageNum).hdc, 0, 0, Form1.picSmallBU.Width, Form1.picSmallBU.Height, _
          Form1.picSmallBU.hdc, 0, 0, vbSrcCopy
   picSEL(ImageNum).Picture = picSEL(ImageNum).Image

End Sub



