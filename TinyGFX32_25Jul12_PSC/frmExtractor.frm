VERSION 5.00
Begin VB.Form frmExtractor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Extractor"
   ClientHeight    =   7320
   ClientLeft      =   150
   ClientTop       =   0
   ClientWidth     =   5640
   ForeColor       =   &H00000000&
   LinkTopic       =   "frmExtractor"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkStandard 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Standard icon sizes only."
      Height          =   225
      Left            =   3345
      TabIndex        =   52
      Top             =   6450
      Width           =   2205
   End
   Begin VB.CommandButton cmdSave32 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save As 32bpp image"
      Height          =   435
      Left            =   4380
      TabIndex        =   50
      Top             =   525
      Width           =   1185
   End
   Begin VB.CheckBox chkAlpha 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Check1"
      Height          =   225
      Left            =   4710
      TabIndex        =   49
      Top             =   4590
      Width           =   225
   End
   Begin VB.CommandButton cmdGotoEnd 
      Caption         =   "GOTO END"
      Height          =   255
      Left            =   1860
      TabIndex        =   44
      Top             =   6075
      Width           =   1575
   End
   Begin VB.CommandButton cmdGotoStart 
      Caption         =   "GOTO START"
      Height          =   255
      Left            =   1860
      TabIndex        =   43
      Top             =   5745
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   330
      Left            =   4830
      TabIndex        =   41
      Top             =   60
      Width           =   765
   End
   Begin VB.CommandButton cmdTransTo 
      Caption         =   "Image 3"
      Height          =   285
      Index           =   2
      Left            =   4530
      TabIndex        =   35
      Top             =   3465
      Width           =   765
   End
   Begin VB.CommandButton cmdTransTo 
      Caption         =   "Image 2"
      Height          =   285
      Index           =   1
      Left            =   4530
      TabIndex        =   34
      Top             =   3045
      Width           =   765
   End
   Begin VB.CommandButton cmdTransTo 
      Caption         =   "Image 1"
      Height          =   285
      Index           =   0
      Left            =   4530
      TabIndex        =   32
      Top             =   2625
      Width           =   765
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "&Open file"
      Height          =   330
      Left            =   75
      TabIndex        =   22
      Top             =   75
      Width           =   915
   End
   Begin VB.CommandButton cmdPrevNine 
      Caption         =   "SHOW PREVIOUS"
      Height          =   255
      Left            =   165
      TabIndex        =   21
      Top             =   6075
      Width           =   1590
   End
   Begin VB.PictureBox PICTBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   4500
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   11
      Top             =   1305
      Width           =   960
   End
   Begin VB.CommandButton cmdNextNine 
      Caption         =   "SHOW NEXT"
      Height          =   255
      Left            =   165
      TabIndex        =   9
      Top             =   5745
      Width           =   1590
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   8
      Left            =   2940
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   8
      Top             =   4320
      Width           =   960
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   7
      Left            =   1605
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   4320
      Width           =   960
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   6
      Left            =   195
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   6
      Top             =   4320
      Width           =   960
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   5
      Left            =   2940
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   2670
      Width           =   960
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   4
      Left            =   1590
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      Top             =   2700
      Width           =   960
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   3
      Left            =   195
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   3
      Top             =   2700
      Width           =   960
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   2
      Left            =   2940
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   1065
      Width           =   960
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   1
      Left            =   1545
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   1080
      Width           =   960
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   0
      Left            =   195
      ScaleHeight     =   85.333
      ScaleMode       =   0  'User
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   1065
      Width           =   960
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "File ="
      Height          =   195
      Left            =   165
      TabIndex        =   51
      Top             =   6720
      Width           =   435
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Preference for 32bpp images:-"
      Height          =   600
      Left            =   4470
      TabIndex        =   48
      Top             =   3960
      Width           =   780
   End
   Begin VB.Label LabAlpha 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "LabAlpha"
      Height          =   375
      Left            =   4470
      TabIndex        =   47
      Top             =   4875
      Width           =   825
   End
   Begin VB.Label LabBU 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "LabBU"
      Height          =   195
      Left            =   1620
      TabIndex        =   46
      Top             =   135
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Info = W x H, bpp: 0 (ico or cur) / 1 (bmp)"
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
      Left            =   180
      TabIndex        =   45
      Top             =   5460
      Width           =   3660
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   4305
      TabIndex        =   42
      Top             =   420
      Width           =   15
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   15
      TabIndex        =   40
      Top             =   5415
      Width           =   5610
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      TabIndex        =   39
      Top             =   15
      Width           =   5655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Click on image to move to transfer box"
      Height          =   420
      Left            =   3765
      TabIndex        =   38
      Top             =   5835
      Width           =   1500
   End
   Begin VB.Label LabFName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   150
      TabIndex        =   37
      Top             =   6945
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Transfer box"
      Height          =   255
      Left            =   4515
      TabIndex        =   36
      Top             =   1050
      Width           =   945
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Transfer to :-"
      Height          =   225
      Left            =   4515
      TabIndex        =   33
      Top             =   2370
      Width           =   930
   End
   Begin VB.Label LabN 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   8
      Left            =   2940
      TabIndex        =   31
      Top             =   3750
      Width           =   465
   End
   Begin VB.Label LabN 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   7
      Left            =   1590
      TabIndex        =   30
      Top             =   3765
      Width           =   465
   End
   Begin VB.Label LabN 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   6
      Left            =   195
      TabIndex        =   29
      Top             =   3765
      Width           =   465
   End
   Begin VB.Label LabN 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   5
      Left            =   2955
      TabIndex        =   28
      Top             =   2145
      Width           =   465
   End
   Begin VB.Label LabN 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   4
      Left            =   1590
      TabIndex        =   27
      Top             =   2145
      Width           =   465
   End
   Begin VB.Label LabN 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   3
      Left            =   195
      TabIndex        =   26
      Top             =   2145
      Width           =   465
   End
   Begin VB.Label LabN 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   2
      Left            =   2940
      TabIndex        =   25
      Top             =   465
      Width           =   465
   End
   Begin VB.Label LabN 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   1
      Left            =   1530
      TabIndex        =   24
      Top             =   465
      Width           =   465
   End
   Begin VB.Label LabN 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   0
      Left            =   195
      TabIndex        =   23
      Top             =   465
      Width           =   465
   End
   Begin VB.Label LabS 
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   8
      Left            =   2940
      TabIndex        =   20
      Top             =   4035
      Width           =   960
   End
   Begin VB.Label LabS 
      BackColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   7
      Left            =   1605
      TabIndex        =   19
      Top             =   4035
      Width           =   960
   End
   Begin VB.Label LabS 
      BackColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   6
      Left            =   195
      TabIndex        =   18
      Top             =   4035
      Width           =   960
   End
   Begin VB.Label LabS 
      BackColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   5
      Left            =   2970
      TabIndex        =   17
      Top             =   2430
      Width           =   960
   End
   Begin VB.Label LabS 
      BackColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   4
      Left            =   1590
      TabIndex        =   16
      Top             =   2430
      Width           =   960
   End
   Begin VB.Label LabS 
      BackColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   3
      Left            =   195
      TabIndex        =   15
      Top             =   2430
      Width           =   960
   End
   Begin VB.Label LabS 
      BackColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   2
      Left            =   2940
      TabIndex        =   14
      Top             =   750
      Width           =   960
   End
   Begin VB.Label LabS 
      BackColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   1
      Left            =   1545
      TabIndex        =   13
      Top             =   750
      Width           =   960
   End
   Begin VB.Label LabS 
      BackColor       =   &H00E0E0E0&
      Height          =   210
      Index           =   0
      Left            =   195
      TabIndex        =   12
      Top             =   750
      Width           =   1200
   End
   Begin VB.Label LabTot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   150
      TabIndex        =   10
      Top             =   6405
      Width           =   2850
   End
End
Attribute VB_Name = "frmExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmExtractor.frm

Option Explicit

Private Declare Function BitBlt Lib "gdi32.dll" _
   (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
   ByVal nWidth As Long, ByVal nHeight As Long, _
   ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
   ByVal dwRop As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
   (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
   
Private Type IcoCurHeader
  iStart As Integer     ' 0
  iType As Integer      ' 1  icon, 2 cursor
  iCount As Integer     ' number of images (1 for single)
  bWidth As Byte        ' Width
  bHeight As Byte       ' Height
  bColorCount As Byte   ' Number of colors (0)
  bRes As Byte          ' Reserved (0)
  iPlanes As Integer    ' Color Planes (1),   or HotX
  ibpp As Integer       ' Bits per pixel (1,4,8,24,32),  or HotY
  LImageSize As Long    ' 40 +[PAL] + XORSize + ANDSize
  LImageOffset As Long  ' Offset to BMIH (22 for single)
End Type
Private IcoHdr As IcoCurHeader

Private FileSize As Long
Private FileString$
Private FileStream() As Byte
Private BStream() As Byte
Private ICOBytePointers() As Long
Private NumICOBytePointers As Long
Private ICOPointers() As Long
Private NumICOPointers As Long
Private IcoWidth As Long
Private IcoHeight As Long
Private ICONPlanes As Integer
Private ICOBPP As Integer
Private ICOWidthN() As Long
Private ICOHeightN() As Long
Private ICOBPPN() As Integer
Private ICOCURN() As Integer
Private ICOorCUR As Integer
Private CURHOTX() As Integer
Private CURHOTY() As Integer
Private PalSize As Long
Private XORSize As Long
Private XORSize32 As Long
Private ANDSize As Long
Private PICNum As Long

Private ICOBMPN() As Long ' 0 ICO(CUR), 1 BMP
Private BMPSize As Long
Private BMPSize32 As Long
Private BMPSizeN() As Long

Private NumIcons As Integer
Private NumBMPs As Integer
Private StartImageNum As Long
Private ImageIndex As Long


Private aChk As Boolean

Private aTBox As Boolean

Private aStanChk As Boolean

'Private svTColor As Long

'Public AppPathSpec$, CurrPath$, FileSpec$, SavePath$, SaveSpec$
Private CommonDialog1 As cOSDialog

Private Sub chkAlpha_Click()
   If Not aChk Then Exit Sub
   If chkAlpha.Value = Checked Then
      aAlphaRestricted = True   ' Alpha restricted
      LabAlpha = "Alpha restricted"
      cmdSave32.Enabled = False
      'PICTBox.Picture = LoadPicture
   Else
      aAlphaRestricted = False  ' Alpha used
      LabAlpha = "Alpha used"
   End If
   If Len(FileSpec$) <> 0 Then
      ShowNine
   End If
End Sub

Private Sub chkStandard_Click()
   aStanChk = chkStandard.Value
   cmdExtract_Click
End Sub


Private Sub cmdSave32_Click()
Dim NImages As Long
Dim BMPorICO As Long
Dim ptr4066 As Long
Dim ImWidth As Long
Dim ImHeight As Long
Dim ImBPP As Long
Dim k As Long
Dim siz As Long
Dim sizX As Long
Dim sizA As Long
Dim FSpec$
Dim Ext$
Dim fnum As Integer


' Refs:
' FileStream() available
' ICOWidthN(NumICOPointers) = IcoWidth
' ICOHeightN(NumICOPointers) = IcoHeight
' ICOBPPN(NumICOPointers) = ICOBPP
' ICOBMPN(k)
' ICOPointers(k)
'
   k = StartImageNum + ImageIndex
   ' ImageIndex from PIC_Click(Index as Integer) transferring to PICTBox
   
   NImages = NumIcons + NumBMPs
   If NImages = 0 Then Exit Sub
   
   BMPorICO = ICOBMPN(k)  ' Denotes 0 ICO(CUR), 1 BMP
   ptr4066 = ICOPointers(k) ' points to 40 for ico or 66(B) for bmp
   'FileStream(ptr4066) ' = 40 for ico or = 66(B) for bmp
   ' ICO ptr4066+40 -> BGRA
   ' BMP ptr4066+54 -> BGRA
   ImWidth = ICOWidthN(k)
   ImHeight = ICOHeightN(k)
   ImBPP = ICOBPPN(k)
   
   If BMPorICO = 0 And FileStream(ptr4066) = 40 Then
   
      ICOFileSpec FSpec$  ' Returns FSpec$ *.ico or *.cur
      
      If Len(FSpec$) = 0 Then Exit Sub
            
      'Private IcoHdr As IcoCurHeader
      sizX = 40 + (ImWidth * ImHeight / 2) * 4
      sizA = (((ImWidth + 7) \ 8 + 3) And &HFFFFFFFC) * ImHeight \ 2
      siz = sizX + sizA
      With IcoHdr
         .iStart = 0
         .iType = 1
         .iCount = 1
         .bWidth = CByte(ImWidth)
         .bHeight = CByte(ImHeight / 2)
         .bColorCount = 0
         .bRes = 0
         .iPlanes = 1   ' or HotX
         .ibpp = 32     ' or HotY
         .LImageSize = siz
         .LImageOffset = 22
      End With
      
      Ext$ = UCase$(Right$(FSpec$, 3))
      If Ext$ = "CUR" Then  'And ICOorCUR = 2 Then ' If image is an icon will changed
         IcoHdr.iType = 2                          ' to a *.cur with HotSpot =[0,0]
         IcoHdr.iPlanes = CURHOTX(k) ' HotX
         IcoHdr.ibpp = CURHOTX(k)    ' HotY
      End If
      
      fnum = FreeFile
      Open FSpec$ For Binary As fnum
      Put fnum, , IcoHdr
      For k = ptr4066 To ptr4066 + siz - 1
         Put fnum, , FileStream(k)
      Next k
      Close
   
   ElseIf BMPorICO = 1 And FileStream(ptr4066) = 66 Then ' BMP
      
      BMPFileSpec FSpec$
      
      If Len(FSpec$) = 0 Then Exit Sub
   
      siz = 54 + (ImWidth * ImHeight) * 4
      fnum = FreeFile
      Open FSpec$ For Binary As fnum
      For k = ptr4066 To ptr4066 + siz - 1
         Put fnum, , FileStream(k)
      Next k
      Close
   End If
End Sub

Private Sub ICOFileSpec(FSpec$)
Dim a$
Dim resp As Long
Dim Title$, Filt$, InDir$
Dim FIndex As Long
   FSpec$ = "" 'AppPathSpec$ & "Test32bpp.ico"
   a$ = "Save As ICO or CUR" & vbCrLf
   a$ = a$ & "Yes for ICO,    No for CUR,     Cancel for neither  "
   resp = MsgBox(a$, vbQuestion + vbYesNoCancel + vbDefaultButton4, "Saving 32 bpp image")
   ' 6 Yes
   ' 7 No
   ' 2 Cancel
   Select Case resp
   Case vbYes
      Title$ = "Save As 32bpp icon"
      Filt$ = "Icon (*.ico)|*.ico"
      InDir$ = ""
      Set CommonDialog1 = New cOSDialog
      CommonDialog1.ShowSave FSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
      Set CommonDialog1 = Nothing
      If Len(FSpec$) <> 0 Then
         FixExtension FSpec$, ".ico"
      End If
   Case vbNo
      Title$ = "Save As 32bpp cursor"
      Filt$ = "Cursor (*.cur)|*.cur"
      InDir$ = ""
      Set CommonDialog1 = New cOSDialog
      CommonDialog1.ShowSave FSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
      Set CommonDialog1 = Nothing
      If Len(FSpec$) <> 0 Then
         FixExtension FSpec$, ".cur"
      End If
   Case vbCancel
      FSpec$ = ""
   End Select
End Sub

Private Sub BMPFileSpec(FSpec$)
Dim Title$, Filt$, InDir$
Dim FIndex As Long
   FSpec$ = "" 'AppPathSpec$ & "Test32bpp.bmp"
   Title$ = "Save As 32bpp bmp file"
   Filt$ = "Bitmap (*.bmp)|*.bmp"
   InDir$ = ""
   Set CommonDialog1 = New cOSDialog
   CommonDialog1.ShowSave FSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex
   Set CommonDialog1 = Nothing
   If Len(FSpec$) <> 0 Then
      FixExtension FSpec$, ".bmp"
   End If
End Sub

Private Sub Form_Activate()
' Fill IcoHdr
   With IcoHdr
      .iStart = 0
      .iType = 1
      .iCount = 1
      .bWidth = 16
      .bHeight = 16
      .bColorCount = 0
      .bRes = 0
      .iPlanes = 1
      .ibpp = 8
      .LImageSize = 0
      .LImageOffset = 0
  End With

End Sub

Private Sub Form_Load()

Dim k As Long
'   AppPathSpec$ = App.Path
'   If Right$(AppPathSpec$, 1) <> "\" Then AppPathSpec$ = AppPathSpec$ & "\"
'   CurrPath$ = AppPathSpec$
   
   PICTBox.BackColor = TColorBGR
   For k = 0 To 8
      LabS(k).Width = 80
      PIC(k).BackColor = TColorBGR
   Next k
   aTBox = False
   
   ' Extractorbackups  13 All, 14 Just last
   If ExtractorBackups = 13 Then
      LabBU = "Preference:  Back up All"
   Else
      LabBU = "Preference:  Back up Final"
   End If
   aChk = False
   If aAlphaRestricted Then
      LabAlpha = "Alpha restricted"
      chkAlpha.Value = Checked
   Else
      LabAlpha = "Alpha used"
      chkAlpha.Value = Unchecked
   End If
   aChk = True
   cmdSave32.Enabled = False
End Sub

Private Sub cmdGotoStart_Click()
   PICTBox.SetFocus
   If NumICOPointers = 0 Then Exit Sub
   StartImageNum = 0
   ShowNine
End Sub

Private Sub cmdGotoEnd_Click()
Dim S As Long

   PICTBox.SetFocus
   If NumICOPointers = 0 Then Exit Sub
   S = StartImageNum
   StartImageNum = NumICOPointers - (NumICOPointers Mod 9)
   If (NumICOPointers Mod 9) = 0 Then
      StartImageNum = StartImageNum - 9
   End If
   If StartImageNum < 0 Then
      StartImageNum = 0
   End If
   ShowNine
End Sub

Private Sub cmdNextNine_Click()
   PICTBox.SetFocus
   If NumICOPointers = 0 Then Exit Sub
   
   If StartImageNum + 8 < NumICOPointers - 1 Then
      StartImageNum = StartImageNum + 9
   End If
   
   ShowNine ' <<<<<<<
End Sub

Private Sub cmdPrevNine_Click()
   If NumICOPointers = 0 Then Exit Sub
      StartImageNum = StartImageNum - 9
      If StartImageNum < 0 Then
         StartImageNum = 0
      End If
      ShowNine
End Sub

Private Sub cmdExtract_Click()
'Open File button
Dim p As Long, p2 As Long
Dim k As Long, k2 As Long
Dim H$
Dim Title$, Filt$, InDir$
Dim FIndex As Long
Dim Ext$
Dim TempFileSpec$
Dim fnum As Integer
   
   
'   For k = 0 To 8
'      PIC(k).Visible = True
'      PIC(k).Picture = LoadPicture
'      LabS(k) = ""
'      LabN(k) = ""
'   Next k
'   LabTot = ""
'   LabFName = ""
   
   Title$ = "Load any file"
   H$ = "bmp,exe,dll,frx,ctx,ico,cur,icl,res,ani,ica,ocx"
   H$ = H$ & "|*.bmp;*.exe;*.dll;*.frx;*.ctx;*.ico;*.cur;*.icl;*.res;*.ani;*.ica;*.ocx"
   
   H$ = H$ & "|All files (*.*)|*.*"
   
   Filt$ = H$
   TempFileSpec$ = ""
   InDir$ = CurrPath$
   
   Set CommonDialog1 = New cOSDialog

   CommonDialog1.ShowOpen TempFileSpec$, Title$, Filt$, InDir$, "", Me.hwnd, FIndex

   Set CommonDialog1 = Nothing
   
   ' FIndex = 1 "bmp,exe,dll,frx,ctx,ico,cur,icl,res,ani,ica,ocx"
   '        = 2 "|All files (*.*)|*.*"
   '        = 3 "|All in Folder (*.*)|*.*"

   'If Len(FileSpec$) <> 0 Then
      If Len(TempFileSpec$) = 0 Then
'            StartImageNum = 0
'            NumICOPointers = 0
'            LabTot = ""
'            NumIcons = 0
'            NumBMPs = 0
'            For k2 = 0 To 8
'               PIC(k2).Visible = False
'            Next k2
            Exit Sub
      End If
   
      
      
      FileSpec$ = TempFileSpec$
   
      CurrPath$ = FileSpec$
      LabFName = " " & GetFileName(FileSpec$)

      fnum = FreeFile
      Open FileSpec$ For Binary Access Read As fnum
      FileSize = LOF(fnum)
'      ReDim FileStream(0 To FileSize - 1)
'      Get #1, , FileStream
      FileString$ = Space$(FileSize)
      Get fnum, , FileString$
      Close
      
      ReDim ICOBytePointers(0 To 10)
      NumICOBytePointers = 0
      
      H$ = Chr$(40) & Chr$(0) & Chr$(0) & Chr$(0)    ' BMIH 40,0,0,0
      p2 = 1
      
      NumIcons = 0
      NumBMPs = 0
      
      '''''''''''''''''''
      PICTBox.Picture = LoadPicture
      aTBox = False
      cmdSave32.Enabled = False
      
      
      Do
         NumICOBytePointers = NumICOBytePointers + 1
         If NumICOBytePointers > 10 Then
            ReDim Preserve ICOBytePointers(NumICOBytePointers + 10)
         End If
         p = InStr(p2, FileString$, H$) ' FileString$ starts at 1
'         p = InStrB(p2, FileStream, H$) ' FileString$ starts at 1
         If p = 0 Then
            NumICOBytePointers = NumICOBytePointers - 1
            Exit Do
         End If
         ICOBytePointers(NumICOBytePointers - 1) = p
         
         
         p2 = p + 1
         If p2 >= FileSize - 22 Then
            NumICOBytePointers = NumICOBytePointers - 1
            Exit Do
         End If
      Loop
      
      ' Transfer characters to byte array
      ReDim FileStream(0 To FileSize - 1)
      CopyMemory FileStream(0), ByVal FileString$, FileSize
      
' Instead of:-
'      Open FileSpec$ For Binary As #1
'      ReDim FileStream(0 To FileSize - 1)
'      Get #1, , FileStream
'      Close
      
      
      FileString$ = " "
      
      
      If NumICOBytePointers = 0 Then
         MsgBox "No ICONS            ", vbOKOnly, "ETools2ctor"
         
         '''''''''''''''
         'Clear PICs
         For k = 0 To 8
            PIC(k).Visible = True
            PIC(k).Width = 64
            PIC(k).Height = 64
            PIC(k).Picture = LoadPicture
            LabS(k) = ""
            LabN(k) = ""
         Next k
         LabTot = ""
         LabFName = ""
         StartImageNum = 0
         NumICOPointers = 0
         Exit Sub
         '''''''''''''''
      End If
      
'k = FileStream(0)
      
      ' Check for valid BitMapInfoHeaders BMIH
      ReDim ICOPointers(0 To NumICOBytePointers - 1)
      ReDim ICOWidthN(0 To NumICOBytePointers - 1)
      ReDim ICOHeightN(0 To NumICOBytePointers - 1)
      ReDim ICOBPPN(0 To NumICOBytePointers - 1)
      ReDim ICOBMPN(0 To NumICOBytePointers - 1)
      ReDim ICOCURN(0 To NumICOBytePointers - 1)
      
      ReDim CURHOTX(0 To NumICOBytePointers - 1)
      ReDim CURHOTY(0 To NumICOBytePointers - 1)
      
      ReDim BMPSizeN(0 To NumICOBytePointers - 1)
      
      NumIcons = 0
      NumBMPs = 0
      
      NumICOBytePointers = NumICOBytePointers - 1
      'NumICOBytePointers = number of 40,0,0,0 possible BMHIs
      'ICOBytePointers(k) = points to 40,0,0,0
      NumICOPointers = 0
      For k = 0 To NumICOBytePointers '- 1
         If ICOBytePointers(k) + 40 < FileSize Then
            ' Get possible BMHI info
            CopyMemory IcoWidth, FileStream(ICOBytePointers(k) + 4 - 1), 4
            CopyMemory IcoHeight, FileStream(ICOBytePointers(k) + 8 - 1), 4
            CopyMemory ICONPlanes, FileStream(ICOBytePointers(k) + 12 - 1), 2
            CopyMemory ICOBPP, FileStream(ICOBytePointers(k) + 14 - 1), 2
            
            ' Check for likely valid BMHI
            If IcoWidth >= 1 Then ''''''
            If IcoWidth <= MaxWidth Then
            If IcoHeight >= 1 Then ''''
            If IcoHeight <= 2 * MaxHeight Then
            If ICONPlanes = 1 Then
            
If Not aStanChk Then GoTo AllSizes ' Simplest

' For standard icons only, no bmps:-
If IcoWidth = IcoHeight \ 2 Then
If (IcoWidth = 16 Or IcoWidth = 24 Or IcoWidth = 32 Or IcoWidth = 48 Or IcoWidth = 64) Then

AllSizes:
               Select Case ICOBPP
               Case 1, 4, 8, 24, 32
                  ICOBMPN(NumICOPointers) = 0 ' denotes ICO or CUR
                  p = ICOBytePointers(k) - 15   ' 15 back from 40,0,0,0
'k2 = FileStream(p) ' = 40
                  If p >= 0 Then
                     ' Test if bmp or not
                     If FileStream(p) = 66 And FileStream(p + 1) = 77 Then  ' = BM ie a BMP
                        If IcoHeight > MaxHeight Then GoTo Nextk
                        ICOBMPN(NumICOPointers) = 1   ' Denotes a BMP
                        ICOPointers(NumICOPointers) = p  ' = 22 for first BMP
                        ICOWidthN(NumICOPointers) = IcoWidth
                        ICOHeightN(NumICOPointers) = IcoHeight
                        ICOBPPN(NumICOPointers) = ICOBPP
                        ICOCURN(NumICOPointers) = 0
                        CURHOTX(NumICOPointers) = 0
                        CURHOTY(NumICOPointers) = 0
                        NumICOPointers = NumICOPointers + 1
                        NumBMPs = NumBMPs + 1
                     Else
                        ' FileStream(p) <> 66 And FileStream(p + 1) <> 77  ' <> BM
                        ICOPointers(NumICOPointers) = ICOBytePointers(k) - 1
                        
                        p = ICOBytePointers(k) - 21  ' Points to Ico(1) or Cur(2)
                        ICOCURN(NumICOPointers) = FileStream(p)
                        ICOorCUR = ICOCURN(NumICOPointers)
                        CURHOTX(NumICOPointers) = 0
                        CURHOTY(NumICOPointers) = 0
                        If ICOorCUR = 2 Then
                           p = ICOBytePointers(k) - 13
                           CURHOTX(NumICOPointers) = FileStream(p)      ' HotX
                           CURHOTY(NumICOPointers) = FileStream(p + 2)  ' HotY
                        End If
                        ICOWidthN(NumICOPointers) = IcoWidth
                        ICOHeightN(NumICOPointers) = IcoHeight
                        ICOBPPN(NumICOPointers) = ICOBPP
                        NumICOPointers = NumICOPointers + 1
                        NumIcons = NumIcons + 1
                     End If
                  Else
                        ICOPointers(NumICOPointers) = ICOBytePointers(k) - 1
                        ICOWidthN(NumICOPointers) = IcoWidth
                        ICOHeightN(NumICOPointers) = IcoHeight
                        ICOBPPN(NumICOPointers) = ICOBPP
                        CURHOTX(NumICOPointers) = 0
                        CURHOTY(NumICOPointers) = 0
                        NumICOPointers = NumICOPointers + 1
                        NumIcons = NumIcons + 1
                  End If
               End Select
            
            End If
            End If
            End If
            End If
            End If
End If
End If
         End If
Nextk:
      Next k
'   End If


   'NumICOPointers = pointers to likely valid BMHIs
   'ICOPointers(NumICOPointers) to BMHI
   
   StartImageNum = 0
   LabTot = Str$(NumICOPointers) & " images" & Str$(NumIcons) & " icons" & Str$(NumBMPs) & " bmps"
   
   Ext$ = UCase$(Right$(FileSpec$, 3))

   If NumICOPointers = 0 Then
      NumICOPointers = 0
      MsgBox "No images or too big ", vbInformation, "Extractor"
      
      '''''''''''''''
      'Clear PICs
      For k = 0 To 8
         PIC(k).Visible = True
         PIC(k).Width = 64
         PIC(k).Height = 64
         PIC(k).Picture = LoadPicture
         LabS(k) = ""
         LabN(k) = ""
      Next k
      LabFName = ""
      StartImageNum = 0
      NumICOPointers = 0
      LabTot = ""
      Exit Sub
      '''''''''''''''
   End If
   
   If NumICOPointers > 0 Then
      If StartImageNum + 8 > NumICOPointers - 1 Then
         StartImageNum = 0
         For k2 = NumICOPointers - 1 To 8
            PIC(k2).Visible = False
         Next k2
      End If
      ShowNine
   Else
      For k2 = 0 To 8
         PIC(k2).Visible = False
      Next k2
   End If
End Sub


Private Sub ShowNine()
Dim k As Long, k2 As Long
Dim kk As Long
Dim n As Long
Dim NB As Byte, NG As Byte, NR As Byte, NA As Byte
Dim TB As Byte, TG As Byte, TR As Byte
Dim Salpha As Single
Dim kb As Byte

   PICTBox.SetFocus
   PICNum = 0
   If NumICOPointers = 0 Then Exit Sub
   For k = 0 To 8
      PIC(k).Visible = True
      PIC(k).Picture = LoadPicture
      If k > NumICOPointers - 1 Then
         PIC(k).Visible = False
      End If
      LabS(k) = ""
      LabN(k) = ""
   Next k
   
   'NumICOPointers = number of pointers to likely valid BMHIs
   'ICOPointers(NumICOPointers) = pointers to BMHIs
   
   For k = StartImageNum To StartImageNum + 8 'NumICOPointers - 1
      If k > NumICOPointers - 1 Then
         For k2 = PICNum To 8
            PIC(k2).Visible = False
         Next k2
         Exit For
      End If
      
      If ICOBMPN(k) = 0 Then ' ICO(CUR)(ANI)
            ' TEST
            '   With IcoHdr
            '      .iStart = 0
            '      .iType = 1
            '      .iCount = 1
            '      .bWidth = 16
            '      .bHeight = 16
            '      .bColorCount = 0
            '      .bRes = 0
            '      .iPlanes = 1  ' or HotX
            '      .iBpp = 8     ' or HotY
            '      .LImageSize = 0
            '      .LImageOffset = 22
            '  End With
            
      
            With IcoHdr
               .bWidth = ICOWidthN(k)
               .bHeight = ICOHeightN(k) \ 2
               .ibpp = ICOBPPN(k)
            End With
            ' Calc Image size. For 32x32 ANDSize = 128
            ANDSize = (((ICOWidthN(k) + 7) \ 8 + 3) And &HFFFFFFFC) * ICOHeightN(k) \ 2
            
            Select Case ICOBPPN(k)
            Case 1:  PalSize = 8
               XORSize = (((ICOWidthN(k) + 7) \ 8 + 3) And &HFFFFFFFC) * ICOHeightN(k) \ 2
            Case 4:  PalSize = 64
               XORSize = (((ICOWidthN(k) + 1) \ 2 + 3) And &HFFFFFFFC) * ICOHeightN(k) \ 2
            Case 8:  PalSize = 1024
               XORSize = ((ICOWidthN(k) + 3) And &HFFFFFFFC) * ICOHeightN(k) \ 2
            Case 24: PalSize = 0
               kk = (3 * ICOWidthN(k) + 3) And &HFFFFFFFC
               kk = kk - 3 * ICOWidthN(k) ' Pad size
               XORSize = (3 * ICOWidthN(k) + kk) * ICOHeightN(k) \ 2
            Case 32: PalSize = 0
               ' For converting to 24bpp
               kk = (3 * ICOWidthN(k) + 3) And &HFFFFFFFC
               kk = kk - 3 * ICOWidthN(k) ' Pad size
               XORSize = (3 * ICOWidthN(k) + kk) * ICOHeightN(k) \ 2
               
               XORSize32 = (4 * ICOWidthN(k)) * ICOHeightN(k) \ 2
            End Select
            
            IcoHdr.LImageSize = 40 + PalSize + XORSize + ANDSize
            
            If ICOBPPN(k) = 32 Then
               ' 1st. Build BMP header
               ' 2nd. Get BGRA image bytes from ICO, CUR or ANI Filestream
               ' 3rd. Modify BGR bytes using (alpha byte)/255
               ' 4th. Set alpha byte to 255
               ' This only gives correct image for a background
               ' color = TColorBGR
               BMPSize32 = 14 + 40 + (ICOWidthN(k) * ICOHeightN(k)) * 4
               BMPOffset = 54
               
               ReDim BStream(0 To BMPSize32 - 1)
               
               BStream(0) = 66
               BStream(1) = 77
               LngToRGB BMPSize32, BStream(2), BStream(3), BStream(4)
               BStream(6) = 0
               BStream(7) = 0
               BStream(8) = 0
               BStream(9) = 0
               LngToRGB BMPOffset, BStream(10), BStream(11), BStream(12)
               ' Start at BMHI ie bmp offset @ 14,
               For kk = 0 To 40 + XORSize32 - 1    ' 40 + XORSize  , 40 for BMHI
                  kb = FileStream(ICOPointers(k) + kk) 'ICOPointers(0)= 22 offset 22 points to BMHI 40
                  BStream(kk + 14) = kb   ' +14 to allow BM header to be put in
               Next kk
               BStream(22) = CLng(ICOHeightN(k) \ 2)
         '      ' Modify BStream(54 To EOF) BGRs
         '      ' For alpha
               LngToRGB TColorBGR, TR, TG, TB
               For kk = 54 To UBound(BStream) Step 4
                  NB = CByte(BStream(kk))
                  NG = CByte(BStream(kk + 1))
                  NR = CByte(BStream(kk + 2))
                  NA = CByte(BStream(kk + 3))
                  If aAlphaRestricted Then
                     ' Only lets colors through where NA=255
                     If NA <> 255 Then NA = 0
                  End If
                  Salpha = (NA / 255)
                  NB = TB * (1 - Salpha) + NB * Salpha
                  NG = TG * (1 - Salpha) + NG * Salpha
                  NR = TR * (1 - Salpha) + NR * Salpha
                  BStream(kk) = CByte(NB) ' BGR
                  BStream(kk + 1) = CByte(NG)
                  BStream(kk + 2) = CByte(NR)
                  BStream(kk + 3) = 255
               Next kk
               
               IcoHeight = ICOHeightN(k) \ 2
               PIC(PICNum).Width = ICOWidthN(k)
               PIC(PICNum).Height = IcoHeight
      
            Else  ' < 32 bpp
               ReDim BStream(0 To (22 + IcoHdr.LImageSize) - 1)
               CopyMemory BStream(0), IcoHdr.iStart, 13
               '14,15,16,17  IcoHdr.LImageSize
               '18,19,20,21  22
               CopyMemory BStream(14), IcoHdr.LImageSize, 4
               BStream(18) = 22
               CopyMemory BStream(22), FileStream(ICOPointers(k)), IcoHdr.LImageSize
               PIC(PICNum).Width = ICOWidthN(k)
               PIC(PICNum).Height = ICOHeightN(k) \ 2
               IcoHeight = ICOHeightN(k) \ 2
            End If
      ' TEST
      'If ICOBPPN(k) = 1 Then
      '   Open "BStream.txt" For Binary As #2
      '   Put #2, , BStream()
      '   Close
      'End If
      Else   ' BMP
         
         BMPSize = 14 + 40
         BMPSize32 = BMPSize
         
         Select Case ICOBPPN(k)
         Case 1
               BMPSize = BMPSize + 8 + (((ICOWidthN(k) + 7) \ 8 + 3) And &HFFFFFFFC) * ICOHeightN(k)
         Case 4
               BMPSize = BMPSize + 64 + (((ICOWidthN(k) + 1) \ 2 + 3) And &HFFFFFFFC) * ICOHeightN(k)
         Case 8
               BMPSize = BMPSize + 1024 + ((ICOWidthN(k) + 3) And &HFFFFFFFC) * ICOHeightN(k)
         Case 24
               kk = (3 * ICOWidthN(k) + 3) And &HFFFFFFFC
               kk = kk - 3 * ICOWidthN(k)
               BMPSize = BMPSize + (3 * ICOWidthN(k) + kk) * ICOHeightN(k)
         Case 32
               ' For converting to 24bpp
               kk = (3 * ICOWidthN(k) + 3) And &HFFFFFFFC
               kk = kk - 3 * ICOWidthN(k)
               
               BMPSize = BMPSize + (3 * ICOWidthN(k) + kk) * ICOHeightN(k)
               
               BMPSize32 = BMPSize32 + (4 * ICOWidthN(k)) * ICOHeightN(k)
         End Select
         
'         ReDim BStream(0 To BMPSize - 1)
'         CopyMemory BStream(0), FileStream(ICOPointers(k)), BMPSize
         
'         IcoHeight = ICOHeightN(k)
'         PIC(PICNum).Width = ICOWidthN(k)
'         PIC(PICNum).Height = ICOHeightN(k)
         
         If ICOBPPN(k) = 32 Then ' Want to convert 32bpp BMP to 24bpp
            ' 1st. Build BMP header
            ' 2nd. Get BGRA image bytes from BMP Filestream
            ' 3rd. Modify BGR bytes using (alpha byte)/255
            ' 4th. Set alpha byte to 255
            ' This only gives correct image for a background
            ' color = TColorBGR
            ' As in Sub ProcessInFile
            ReDim BStream(0 To BMPSize32 - 1)
            BMPOffset = 54
            BStream(0) = 66
            BStream(1) = 77
            LngToRGB BMPSize, BStream(2), BStream(3), BStream(4)
            BStream(6) = 0
            BStream(7) = 0
            BStream(8) = 0
            BStream(9) = 0
            LngToRGB BMPOffset, BStream(10), BStream(11), BStream(12)
            ' Start at BMHI ie bmp offset @ 14. No multiple bmp files exist.
            For kk = 0 To 40 + (4 * ICOWidthN(k)) * ICOHeightN(k) - 1   ' 40 for BMHI + BGRAs
               kb = FileStream(14 + kk) ' Offset 14 points to BMHI 40
               BStream(kk + 14) = kb
            Next kk
            LngToRGB TColorBGR, TR, TG, TB
            For kk = 54 To UBound(BStream) Step 4
               NB = CByte(BStream(kk))
               NG = CByte(BStream(kk + 1))
               NR = CByte(BStream(kk + 2))
               NA = CByte(BStream(kk + 3))
               If aAlphaRestricted Then
                  ' Only lets colors through where NA=255
                  If NA <> 255 Then NA = 0
               End If
               
               Salpha = (NA / 255)
               NB = TB * (1 - Salpha) + NB * Salpha
               NG = TG * (1 - Salpha) + NG * Salpha
               NR = TR * (1 - Salpha) + NR * Salpha
               BStream(kk) = CByte(NB)
               BStream(kk + 1) = CByte(NG)
               BStream(kk + 2) = CByte(NR)
               BStream(kk + 3) = 255
            Next kk

         Else   ' BMP BPP <32
            ReDim BStream(0 To BMPSize - 1)
            CopyMemory BStream(0), FileStream(ICOPointers(k)), BMPSize
         End If
      
         IcoHeight = ICOHeightN(k)
         PIC(PICNum).Width = ICOWidthN(k)
         PIC(PICNum).Height = ICOHeightN(k)
      
      End If
      
      '-------------------------------------------------------------------------
      Set PIC(PICNum).Picture = PictureFromByteStream(BStream())
      PIC(PICNum).Picture = PIC(PICNum).Image
      
      LabN(PICNum) = k + 1
      PICNum = PICNum + 1
      IcoWidth = ICOWidthN(k)
      LabS(n) = Str$(IcoWidth) & " x" & Str$(IcoHeight) & " ," & Str$(ICOBPPN(k)) & ":" & Str$(ICOBMPN(k))
      n = n + 1
      If FileStream(ICOPointers(k)) + IcoHdr.LImageSize > FileSize - 1 Then
         Exit For
      End If
   
   Next k

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Erase FileStream, BStream
   NumICOPointers = 0
   Unload Me
End Sub

Private Sub cmdClose_Click()
Dim k As Long
   Erase FileStream, BStream
   NumICOPointers = 0
' TEST
   k = ImageNum
   Unload Me
End Sub

Private Sub LabN_Click(Index As Integer)
   PIC_Click Index
End Sub

Private Sub LabS_Click(Index As Integer)
   PIC_Click Index
End Sub

Private Sub PIC_Click(Index As Integer)
Dim W As Long
Dim H As Long
Dim k As Long
Dim NImages As Long
   
   
   NImages = NumIcons + NumBMPs
   If NImages = 0 Then Exit Sub
   

   ImageIndex = Index  ' Needed for Save As 32bpp
   k = StartImageNum + ImageIndex

   If k <= UBound(ICOBPPN()) Then
      If ICOBPPN(k) = 32 And aAlphaRestricted = False Then
         cmdSave32.Enabled = True
      Else
         cmdSave32.Enabled = False
      End If
      
      If ICOBMPN(k) = 0 Then
         cmdSave32.Caption = "Save As 32bpp ico/cur"
      Else
         cmdSave32.Caption = "Save As 32bpp bmp"
      End If
   Else
      Exit Sub
   End If
   W = PIC(Index).Width
   H = PIC(Index).Height
   ' Transfer
   PICTBox.Picture = LoadPicture
   PICTBox.Width = W
   PICTBox.Height = H
   BitBlt PICTBox.hdc, 0, 0, W, H, PIC(Index).hdc, 0, 0, vbSrcCopy
   PICTBox.Picture = PICTBox.Image
   aTBox = True
End Sub

Private Sub cmdTransTo_Click(Index As Integer)
Dim cx1 As Single, cx2 As Single
Dim cy1 As Single, cy2 As Single
Dim k As Long
Dim NImages

Dim ix As Long, iy As Long
Dim kk As Long

Dim ptr4066 As Long
Dim ptrBGRA As Long
Dim ImWidth As Long
Dim ImHeight As Long


   If Not aTBox Then Exit Sub
   
   NImages = NumIcons + NumBMPs
   If NImages = 0 Then Exit Sub
   
   k = StartImageNum + ImageIndex
   ' ImageIndex from PIC_Click(Index as Integer) transferring to PICTBox

   'Public BPP(0 To 2) As Long
   If k <= UBound(ICOBPPN()) Then
      If ICOBPPN(k) = 32 And aAlphaRestricted Then
         ICOBPPN(k) = 24
      End If
      bpp(Index) = ICOBPPN(k)
      ImBPP(Index) = bpp(Index)
      
      If ICOorCUR = 2 Then
         'p = ICOBytePointers(k) - 13
         HotX(Index) = CURHOTX(k) '= FileStream(p)    ' HotX
         HotY(Index) = CURHOTY(k) '= FileStream(p + 2) ' HotY
      End If
   End If
   
   With Form1.picSmall(Index)
      .Width = PICTBox.Width
      .Height = PICTBox.Height
      .BackColor = TColorBGR
      .Picture = LoadPicture
      .Picture = .Image
   End With
   With Form1
      .picPANEL.BackColor = TColorBGR
      .picPANEL.Picture = LoadPicture
      .picPANEL.Picture = .picPANEL.Image
   End With
   
   BitBlt Form1.picSmall(Index).hdc, 0, 0, PICTBox.Width, PICTBox.Height, _
          PICTBox.hdc, 0, 0, vbSrcCopy
   ' Can change TColorBGR !!??
   
   Form1.picSmall(Index).Picture = Form1.picSmall(Index).Image
   
   ImageHeight(Index) = PICTBox.Height
   ImageWidth(Index) = PICTBox.Width
   
   Form1.LabWH(Index) = "W x H =" & Str$(ImageWidth(Index)) & " x" & Str$(ImageHeight(Index))
   
   ' Draw around current picSmall
   ImageNum = CLng(Index)
   With Form1
      cx1 = .picSmall(ImageNum).Left - 1
      cx2 = cx1 + .picSmall(ImageNum).Width + 1
      cy1 = .picSmall(ImageNum).Top - 1
      cy2 = cy1 + .picSmall(ImageNum).Height + 1
      .picSmallFrame.Cls
      .picSmallFrame.DrawStyle = 2
      .picSmallFrame.Line (cx1 - 1, cy1 - 1)-(cx2 + 1, cy2 + 1), 0, B
   End With
   
'---------------------------------------------------------------------
   If ICOBPPN(k) = 32 Then ' Extract Alpha bytes from FileStream
      ptr4066 = ICOPointers(k) ' points to 40 for ico or 66(B) for bmp
      'FileStream(ptr4066) ' = 40 for ico or = 66(B) for bmp
      ' ICO ptr4066+40 -> BGRA
      ' BMP ptr4066+54 -> BGRA
      ImWidth = ICOWidthN(k)
      ImHeight = ICOHeightN(k)
      
' Dim TextB as byte
'      TestB = FileStream(ptr4066)
   
      If ICOBMPN(k) = 0 Then  ' ICO,CUR
         ptrBGRA = ptr4066 + 40
         ImHeight = ImHeight \ 2
      Else ' BMP
         ptrBGRA = ptr4066 + 54
      End If
      ReDim DATACULSRC(0 To 3, ImWidth - 1, ImHeight - 1)
      For kk = ptrBGRA To ptrBGRA + 4 * (ImWidth * ImHeight) - 1 Step 4
         DATACULSRC(0, ix, iy) = FileStream(kk + 0)
         DATACULSRC(1, ix, iy) = FileStream(kk + 1)
         DATACULSRC(2, ix, iy) = FileStream(kk + 2)
         DATACULSRC(3, ix, iy) = FileStream(kk + 3)
         ix = ix + 1
         If ix > ImWidth - 1 Then
            ix = 0
            iy = iy + 1
         End If
      Next kk
         
      Select Case ImageNum
      Case 0
         ReDim DATACUL0(0 To 3, ImWidth - 1, ImHeight - 1)
         FILL3D DATACUL0(), DATACULSRC()
      Case 1
         ReDim DATACUL1(0 To 3, ImWidth - 1, ImHeight - 1)
         FILL3D DATACUL1(), DATACULSRC()
      Case 2
         ReDim DATACUL2(0 To 3, ImWidth - 1, ImHeight - 1)
         FILL3D DATACUL2(), DATACULSRC()
      End Select

   End If   ' If ICOBPPN(k) = 32 Then

   ' Extractorbackups  13 All, 14 Just last
   If ExtractorBackups = 14 Then ' Just back up final image, so clear any others
      Form1.KILLSAVS CLng(Index)
   End If
   Form1.BackUp (ImageNum)
   Form1.LabSpec = "EX"
   NameSpec$(ImageNum) = "Extracted (bpp=" & Str$(bpp(Index)) & ")"
   ImFileSpec$(ImageNum) = FileSpec$

End Sub

