VERSION 5.00
Begin VB.Form frmRotator 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Rotator"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3315
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   221
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdACCCAN 
      Caption         =   "Cancel/Close"
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
      Left            =   1815
      TabIndex        =   16
      Top             =   4080
      Width           =   1305
   End
   Begin VB.CommandButton cmdACCCAN 
      Caption         =   "Accept/Close"
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
      Left            =   285
      TabIndex        =   15
      Top             =   4080
      Width           =   1305
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00E0E0E0&
      Height          =   1170
      Left            =   195
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   10
      Top             =   1365
      Width           =   2925
      Begin VB.OptionButton optSpace 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Replace by Right color"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   14
         Top             =   750
         Width           =   2490
      End
      Begin VB.OptionButton optSpace 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Replace by transparent color"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   13
         Top             =   405
         Width           =   2490
      End
      Begin VB.OptionButton optSpace 
         BackColor       =   &H00E0E0E0&
         Caption         =   "No replacement"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   12
         Top             =   90
         Width           =   2490
      End
   End
   Begin VB.CheckBox chkAA 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anti-alias"
      Height          =   240
      Left            =   2160
      TabIndex        =   9
      Top             =   1095
      Width           =   930
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
      Height          =   330
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   1305
      Picture         =   "frmRotator.frx":0000
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   3
      Top             =   2775
      Width           =   720
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   5
      Left            =   780
      Max             =   180
      Min             =   -180
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Value           =   1
      Width           =   1680
   End
   Begin VB.Label LabRot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+90"
      Height          =   255
      Index           =   3
      Left            =   2835
      TabIndex        =   20
      Top             =   705
      Width           =   315
   End
   Begin VB.Label LabRot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+45"
      Height          =   255
      Index           =   2
      Left            =   2475
      TabIndex        =   19
      Top             =   705
      Width           =   315
   End
   Begin VB.Label LabRot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-45"
      Height          =   255
      Index           =   1
      Left            =   435
      TabIndex        =   18
      Top             =   705
      Width           =   315
   End
   Begin VB.Label LabRot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-90"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   17
      Top             =   705
      Width           =   315
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1050
      Left            =   1140
      Top             =   2730
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vacated space :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   11
      Top             =   1110
      Width           =   1710
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "deg"
      Height          =   195
      Left            =   2790
      TabIndex        =   8
      Top             =   240
      Width           =   330
   End
   Begin VB.Label LabSelRect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select rectangle OFF"
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   2280
      TabIndex        =   7
      Top             =   2895
      Width           =   840
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "by :"
      Height          =   240
      Index           =   2
      Left            =   1755
      TabIndex        =   6
      Top             =   255
      Width           =   270
   End
   Begin VB.Label LabIMNum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   285
      Left            =   1380
      TabIndex        =   5
      Top             =   225
      Width           =   240
   End
   Begin VB.Label LabDeg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "deg"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2055
      TabIndex        =   2
      Top             =   225
      Width           =   705
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Rotate image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   255
      Width           =   1230
   End
End
Attribute VB_Name = "frmRotator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

' Getting image to data array
Private Declare Function GetDIBits Lib "gdi32.dll" _
   (ByVal aHDC As Long, ByVal hBitmap As Long, _
   ByVal nStartScan As Long, ByVal nNumScans As Long, _
   ByRef lpBits As Any, _
   ByRef BInfo As BITMAPINFOHEADER, _
   ByVal wUsage As Long) As Long

Private Declare Function SetDIBits Lib "gdi32.dll" _
   (ByVal hdc As Long, _
    ByVal hBitmap As Long, _
    ByVal nStartScan As Long, _
    ByVal nNumScans As Long, _
    ByRef lpBits As Any, _
    ByRef lpBI As BITMAPINFOHEADER, _
    ByVal wUsage As Long) As Long

Private Declare Function BitBlt Lib "gdi32.dll" _
   (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
   ByVal nWidth As Long, ByVal nHeight As Long, _
   ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
   ByVal dwRop As Long) As Long


Private Ndeg As Long

'Private picDATAREG() As Byte
Private DATACULORG() As Byte

Private PW As Long, PH As Long
Private aAlias As Long
Private aSpace As Long
'Private aErase As Boolean

Private BCulR As Byte, BCulG As Byte, BCulB As Byte, BCulA As Byte
Private Const pi# = 3.14159265359
Private Const d2r# = pi# / 180

Private aShow As Boolean


Private Sub Form_Load()

' Enter with ImageNum, LOX etc
Dim ix As Long, iy As Long
Dim ABYTE As Byte
Dim BMIH As BITMAPINFOHEADER
   aShow = False

   frmRotator.Left = frmRotateLeft
   frmRotator.Top = frmRotateTop

   If aSelect Then
      LabSelRect = "Select Rectangle ON"
   Else
      LabSelRect = "Select Rectangle OFF"
   End If
   
   ' TColorBGR
   LngToRGB TColorBGR, BCulR, BCulG, BCulB
'   BCulR = 194
'   BCulG = 195
'   BCulB = 197
   
'   ' Set at mnuRotator
'   LOX = 0: HIX = PW - 1
'   LOY = 0: HIY = PH - 1
   
   LabIMNum = ImageNum + 1
   
   PW = ImageWidth(ImageNum)
   PH = ImageHeight(ImageNum)
   
   With PIC
      .Width = PW
      .Height = PH
      .BackColor = TColorBGR
      .Picture = LoadPicture
      .Picture = .Image
   End With
   
   PIC.Left = Image1.Left + (Image1.Width - PW) \ 2
   
   ' Copy Form1.picSmall() to PIC
   BitBlt PIC.hdc, 0, 0, PW, PH, _
          Form1.picSmall(ImageNum).hdc, 0, 0, vbSrcCopy
   PIC.Picture = PIC.Image
   
   GetTheBitsBGR ImageNum, Form1.picSmall(ImageNum)   ' picSmall(ImageNum) -->> PICDATAREG()

   ReDim picDATAORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   picDATAORG() = picDATAREG()   ' Needed in case SELREC
   
   chkAA.Enabled = True
   
   '---------------------------------------------------------
   
   If ImBPP(ImageNum) = 32 Then
      ' ReDim DATACULORG()  & fill with DATACUL#()
      ReDim DATACULORG(0 To 3, ImageWidth(ImageNum) - 1, ImageHeight(ImageNum) - 1)
      Select Case ImageNum
      Case 0: FILL3D DATACULORG(), DATACUL0()   ' Dest,Src
      Case 1: FILL3D DATACULORG(), DATACUL1()
      Case 2: FILL3D DATACULORG(), DATACUL2()
      End Select
   End If
   '---------------------------------------------------------
   aTransfer = False
   'aErase = False
   aAlias = 0
   aSpace = 1   ' Vacated space replacement (1 transparent color)
   Ndeg = 0
   HScroll1.Value = 0
   ' Replace with transparent color
   optSpace(1).Value = True
   optSpace_Click 1
   aShow = True
End Sub

Private Sub cmdReset_Click()
Dim BMIH As BITMAPINFOHEADER

   If ImBPP(ImageNum) = 32 Then  ' Reset DATACUL#()
      Select Case ImageNum
         Case 0: FILL3D DATACUL0(), DATACULORG()
         Case 1: FILL3D DATACUL1(), DATACULORG()
         Case 2: FILL3D DATACUL2(), DATACULORG()
      End Select
   End If
   picDATAREG() = picDATAORG()   ' Reset picDATA
'------------------------------------------
   ' Rest angle to zero
   aShow = False
      HScroll1.Value = 0
      Ndeg = 0
      LabDeg = Ndeg
   aShow = True
'------------------------------------------
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = PIC.Width
      .biHeight = PIC.Height
      .biBitCount = 32
   End With
      
   ' Reset PIC from picDATAREG() or picDATAORG()
   If SetDIBits(PIC.hdc, PIC.Image, _
      0, PIC.Height, picDATAREG(0, 0, 0), BMIH, 0) = 0 Then
      MsgBox "SetDIBits Error", vbCritical, "Rotate"
   End If
   PIC.Picture = PIC.Image
   
   ' Show Reset on Form1
   With Form1
      BitBlt .picSmall(ImageNum).hdc, 0, 0, PIC.Width, PIC.Height, _
             PIC.hdc, 0, 0, vbSrcCopy
      .picSmall(ImageNum).Picture = .picSmall(ImageNum).Image
      .DrawGrid
   End With
   
   aTransfer = False
End Sub

Private Sub Transfer()
   With Form1
      BitBlt .picSmall(ImageNum).hdc, 0, 0, PIC.Width, PIC.Height, _
             PIC.hdc, 0, 0, vbSrcCopy
      .picSmall(ImageNum).Picture = .picSmall(ImageNum).Image
      
      .picSmallBU.Width = PIC.Width
      .picSmallBU.Height = PIC.Height
      BitBlt .picSmallBU.hdc, 0, 0, PIC.Width, PIC.Height, _
             PIC.hdc, 0, 0, vbSrcCopy
      .picSmallBU.Picture = .picSmallBU.Image
      
      BitBlt .picORG(ImageNum).hdc, 0, 0, PIC.Width, PIC.Height, _
             PIC.hdc, 0, 0, vbSrcCopy
      .picSmallBU.Picture = .picSmallBU.Image
      
      .DrawGrid
   End With
End Sub

Private Sub HScroll1_Scroll()
   Call HScroll1_Change
End Sub

Private Sub HScroll1_Change()
   Ndeg = HScroll1.Value
   LabDeg = Ndeg
 
   If Not aShow Then Exit Sub
   PIC.SetFocus
   If Ndeg = 0 Then
      cmdReset_Click
   Else
      Rotate CSng(Ndeg)
   End If
End Sub


Private Sub chkAA_Click()
   aAlias = chkAA.Value
   HScroll1_Scroll
   'cmdReset_Click
End Sub

Private Sub LabRot_Click(Index As Integer)
Dim v As Long
   Select Case Index
   Case 0: v = -90
   Case 1: v = -45
   Case 2: v = 45
   Case 3: v = 90
   End Select
   HScroll1.Value = v
End Sub

Private Sub optSpace_Click(Index As Integer)
   aSpace = Index
   cmdReset_Click
End Sub

Private Sub Rotate(zPROTATE As Single)
' zPROTATE  degrees  -180 - + 180
Dim BMIH As BITMAPINFOHEADER

Dim ixc As Single, iyc As Single
Dim ixd As Long, iyd As Long
Dim xs As Single, ys As Single
Dim ixs As Long, iys As Long
Dim zang As Double
Dim zcos As Double, zsin As Double
'Dim CH As Long, CW As Long
Dim CH As Single, CW As Single

Dim idx As Single

' For AA
Dim xsf As Single, ysf As Single
Dim ixsp1 As Single, iysp1 As Single  ' ixs+1 & iys+1

Dim CulB As Long, CulG As Long, CulR As Long, CulA As Long
Dim CulB0 As Long, CulG0 As Long, CulR0 As Long, CulA0 As Long
Dim CulB1 As Long, CulG1 As Long, CulR1 As Long, CulA1 As Long

'  Put newly rotated image into picDATAREG()
'  also, if 32bpp, DATACUL#() has rotated bytes including Alpha
   
   CH = HIY - LOY
   CW = HIX - LOX
   
   'ABYTE = picDATAORG(3, ix, iy)  ' will = 255 everywhere
   'ABYTE = picDATAREG(3, ix, iy)  ' will = 255 everywhere
   
   Select Case aSpace
   Case 0   ' Unchanged, ensure changed back to original
      picDATAREG() = picDATAORG()   ' Reset picDATA
      If ImBPP(ImageNum) = 32 Then  ' Reset DATACUL#()
         Select Case ImageNum
            Case 0: FILL3D DATACUL0(), DATACULORG()    ' Dest, Src
            Case 1: FILL3D DATACUL1(), DATACULORG()
            Case 2: FILL3D DATACUL2(), DATACULORG()
         End Select
      End If
   Case 1   ' Fill picDATAREG() rectangle (selection) with Transparent color
      LngToRGB TColorBGR, BCulR, BCulG, BCulB
      BCulA = 0
   Case 2   ' Fill picDATAREG() & DATACUL#() rectangle (selection) with RColor color
      LngToRGB RColor, BCulR, BCulG, BCulB
      BCulA = 255
   End Select
   
   If aSpace = 1 Or aSpace = 2 Then
      For iys = LOY To HIY
      For ixs = LOX To HIX
         picDATAREG(0, ixs, iys) = BCulB
         picDATAREG(1, ixs, iys) = BCulG
         picDATAREG(2, ixs, iys) = BCulR
         picDATAREG(3, ixs, iys) = BCulA 'picDATAORG(3, ixs, iys), 0 or 255
      Next ixs
      Next iys
   End If
      
   If aSpace = 2 Then    ' Fill with Right Color
      If ImBPP(ImageNum) = 32 Then
         For iys = LOY To HIY
         For ixs = LOX To HIX
             Select Case ImageNum
             Case 0
                DATACUL0(0, ixs, iys) = BCulB
                DATACUL0(1, ixs, iys) = BCulG
                DATACUL0(2, ixs, iys) = BCulR
                DATACUL0(3, ixs, iys) = BCulA 'DATACULORG(3, ixs, iys)
             Case 1
                DATACUL1(0, ixs, iys) = BCulB
                DATACUL1(1, ixs, iys) = BCulG
                DATACUL1(2, ixs, iys) = BCulR
                DATACUL1(3, ixs, iys) = BCulA 'DATACULORG(3, ixs, iys)
             Case 2
                DATACUL2(0, ixs, iys) = BCulB
                DATACUL2(1, ixs, iys) = BCulG
                DATACUL2(2, ixs, iys) = BCulR
                DATACUL2(3, ixs, iys) = BCulA 'DATACULORG(3, ixs, iys)
            End Select
         Next ixs
         Next iys
      End If
   End If
   
   ' Some anomaly with |45| degrees ??
   If zPROTATE = 45 Then zPROTATE = 45.1
   If zPROTATE = -45 Then zPROTATE = -45.1
   If zPROTATE = 135 Then zPROTATE = 135.1
   If zPROTATE = -135 Then zPROTATE = -135.1
   
   zang = zPROTATE * d2r#
   
   zcos = Cos(zang)
   If Abs(zcos) < 10 ^ -3 Then zcos = 0
   zsin = Sin(-zang)
   If Abs(zsin) < 10 ^ -3 Then zsin = 0
   ixc = CSng(LOX + CW / 2)
   iyc = CSng(LOY + CH / 2)
   
   ' Adjust center maintains shape for 90 rot but
   ' offsets for even dimensions ??
   'If ixc - Int(ixc) <> 0 Then ixc = Int(ixc) + 1
   'If iyc - Int(iyc) <> 0 Then iyc = Int(iyc) + 1
   
   If aAlias = 0 Then
      For iyd = 0 To PH - 1
         idx = CW / 2
         For ixd = 0 To PW - 1
            xs = ixc + (ixd - ixc) * zcos + (iyd - iyc) * zsin
            ys = iyc + (iyd - iyc) * zcos - (ixd - ixc) * zsin
            If xs = Int(xs) Then xs = xs + 0.001
            If ys = Int(ys) Then ys = ys + 0.001
            ixs = CLng(xs)
            iys = CLng(ys)
            ' Check that rotated source pixel lies within rectangle
            If ixs >= LOX Then
            If ixs <= HIX Then
            If iys >= LOY Then
            If iys <= HIY Then
               picDATAREG(0, ixd, iyd) = picDATAORG(0, ixs, iys)
               picDATAREG(1, ixd, iyd) = picDATAORG(1, ixs, iys)
               picDATAREG(2, ixd, iyd) = picDATAORG(2, ixs, iys)
               picDATAREG(3, ixd, iyd) = picDATAORG(3, ixs, iys)  ' alpha byte
               If ImBPP(ImageNum) = 32 Then
                  Select Case ImageNum
                  Case 0
                     DATACUL0(0, ixd, iyd) = DATACULORG(0, ixs, iys)
                     DATACUL0(1, ixd, iyd) = DATACULORG(1, ixs, iys)
                     DATACUL0(2, ixd, iyd) = DATACULORG(2, ixs, iys)
                     DATACUL0(3, ixd, iyd) = DATACULORG(3, ixs, iys)
                     'CulB = DATACUL0(0, ixd, iyd)
                     'CulG = DATACUL0(1, ixd, iyd)
                     'CulR = DATACUL0(2, ixd, iyd)
                     'CulA = DATACUL0(3, ixd, iyd)
                     'picDATAREG(3, ixd, iyd) = CulA 'picDATAORG(3, ixs, iys)  ' alpha byte
                  Case 1
                     DATACUL1(0, ixd, iyd) = DATACULORG(0, ixs, iys)
                     DATACUL1(1, ixd, iyd) = DATACULORG(1, ixs, iys)
                     DATACUL1(2, ixd, iyd) = DATACULORG(2, ixs, iys)
                     DATACUL1(3, ixd, iyd) = DATACULORG(3, ixs, iys)
                  Case 2
                     DATACUL2(0, ixd, iyd) = DATACULORG(0, ixs, iys)
                     DATACUL2(1, ixd, iyd) = DATACULORG(1, ixs, iys)
                     DATACUL2(2, ixd, iyd) = DATACULORG(2, ixs, iys)
                     DATACUL2(3, ixd, iyd) = DATACULORG(3, ixs, iys)
                  End Select
              End If
            End If
            End If
            End If
            End If
            If ImBPP(ImageNum) = 32 Then   ' Alphas blurred!
               'Replace_Alpha_DATACUL ixd, iyd, CByte(CulA), CByte(CulB), CByte(CulG), CByte(CulR), 1
               'ReconcileForAlphaXY ImageNum, ixd, iyd
            End If
         Next ixd
      Next iyd
         
   Else  ' aAlias  ' Problem with Alpha since altering TColotBGR values
      For iyd = 0 To PH - 1
         For ixd = 0 To PW - 1
            xs = ixc + (ixd - ixc) * zcos + (iyd - iyc) * zsin
            ys = iyc + (iyd - iyc) * zcos - (ixd - ixc) * zsin
            If xs = Int(xs) Then xs = xs + 0.001
            If ys = Int(ys) Then ys = ys + 0.001
            ixs = Int(xs)
            iys = Int(ys)
            ' Check that rotated source pixel lies within rectangle
            If ixs >= LOX Then
            If ixs <= HIX Then
            If iys >= LOY Then
            If iys <= HIY Then
               ixsp1 = ixs + 1
               If ixsp1 > HIX Then ixsp1 = HIX 'ixs - 1
               iysp1 = iys + 1
               If iysp1 > HIY Then iysp1 = HIY 'iys - 1
               
               xsf = xs - ixs
               ysf = ys - iys
               ' picDATAORG() = Colors from PIC
               ' ie TColorBGR is transparent area
               ' picDATAORG(3,ix,iy) = 255
               
               CulB = (1 - xsf) * picDATAORG(0, ixs, iys)
               CulG = (1 - xsf) * picDATAORG(1, ixs, iys)
               CulR = (1 - xsf) * picDATAORG(2, ixs, iys)
               'CulA = (1 - xsf) * picDATAORG(3, ixs, iys)
         
               CulB0 = CulB + xsf * picDATAORG(0, ixsp1, iys)
               CulG0 = CulG + xsf * picDATAORG(1, ixsp1, iys)
               CulR0 = CulR + xsf * picDATAORG(2, ixsp1, iys)
               'CulA0 = CulA + xsf * picDATAORG(3, ixsp1, iys)
               
               CulB = (1 - xsf) * picDATAORG(0, ixs, iysp1)
               CulG = (1 - xsf) * picDATAORG(1, ixs, iysp1)
               CulR = (1 - xsf) * picDATAORG(2, ixs, iysp1)
               'CulA = (1 - xsf) * picDATAORG(3, ixs, iysp1)
                  
               
               CulB1 = CulB + xsf * picDATAORG(0, ixsp1, iysp1)
               CulG1 = CulG + xsf * picDATAORG(1, ixsp1, iysp1)
               CulR1 = CulR + xsf * picDATAORG(2, ixsp1, iysp1)
               'CulA1 = CulA + xsf * picDATAORG(3, ixsp1, iysp1)
               
               CulB = (1 - ysf) * CulB0 + ysf * CulB1
               CulG = (1 - ysf) * CulG0 + ysf * CulG1
               CulR = (1 - ysf) * CulR0 + ysf * CulR1
               
               'CulA = (1 - ysf) * CulA0 + ysf * CulA1
               
                  If CulB > 255 Then
                     CulB = 255
                  ElseIf CulB < 0 Then
                     CulB = 0
                  End If
                  If CulG > 255 Then
                     CulG = 255
                  ElseIf CulG < 0 Then
                     CulG = 0
                  End If
                  If CulR > 255 Then
                     CulR = 255
                  ElseIf CulR < 0 Then
                     CulR = 0
                  End If
'                  If CulA > 255 Then
'                     CulA = 255
'                  ElseIf CulA < 0 Then
'                     CulA = 0
'                  End If
               
               picDATAREG(0, ixd, iyd) = CulB
               picDATAREG(1, ixd, iyd) = CulG
               picDATAREG(2, ixd, iyd) = CulR
               picDATAREG(3, ixd, iyd) = 255 'CulA   ' alpha byte
               
               If ImBPP(ImageNum) = 32 Then
               
                  ' DATACULORG#() = Original colors
                  ' with Alphas in DATACULORG(3,ix.iy)
                  
                  ' Picks up colors and Alpha from 4 pixels
                  ' so get an area average of Colors OK but
                  ' Alphas ??
                  
                  CulB = (1 - xsf) * DATACULORG(0, ixs, iys)
                  CulG = (1 - xsf) * DATACULORG(1, ixs, iys)
                  CulR = (1 - xsf) * DATACULORG(2, ixs, iys)
                  CulA = (1 - xsf) * DATACULORG(3, ixs, iys)
                  
                  CulB0 = CulB + xsf * DATACULORG(0, ixsp1, iys)
                  CulG0 = CulG + xsf * DATACULORG(1, ixsp1, iys)
                  CulR0 = CulR + xsf * DATACULORG(2, ixsp1, iys)
                  CulA0 = CulA + xsf * DATACULORG(3, ixsp1, iys)
                  
                  CulB = (1 - xsf) * DATACULORG(0, ixs, iysp1)
                  CulG = (1 - xsf) * DATACULORG(1, ixs, iysp1)
                  CulR = (1 - xsf) * DATACULORG(2, ixs, iysp1)
                  CulA = (1 - xsf) * DATACULORG(3, ixs, iysp1)
               
                  CulB1 = CulB + xsf * DATACULORG(0, ixsp1, iysp1)
                  CulG1 = CulG + xsf * DATACULORG(1, ixsp1, iysp1)
                  CulR1 = CulR + xsf * DATACULORG(2, ixsp1, iysp1)
                  CulA1 = CulA + xsf * DATACULORG(3, ixsp1, iysp1)
                  
                  CulB = (1 - ysf) * CulB0 + ysf * CulB1
                  CulG = (1 - ysf) * CulG0 + ysf * CulG1
                  CulR = (1 - ysf) * CulR0 + ysf * CulR1
                  CulA = (1 - ysf) * CulA0 + ysf * CulA1

                  'BCulR , BCulG, BCulB

                  If CulB > 255 Then
                     CulB = 255
                  ElseIf CulB < 0 Then
                     CulB = 0
                  End If
                  If CulG > 255 Then
                     CulG = 255
                  ElseIf CulG < 0 Then
                     CulG = 0
                  End If
                  If CulR > 255 Then
                     CulR = 255
                  ElseIf CulR < 0 Then
                     CulR = 0
                  End If
                  If CulA > 255 Then
                     CulA = 255
                  ElseIf CulA < 0 Then
                     CulA = 0
                  End If
                  Select Case ImageNum
                  Case 0
                     DATACUL0(0, ixd, iyd) = CulB
                     DATACUL0(1, ixd, iyd) = CulG
                     DATACUL0(2, ixd, iyd) = CulR
                     DATACUL0(3, ixd, iyd) = CulA
                  Case 1
                     DATACUL1(0, ixd, iyd) = CulB
                     DATACUL1(1, ixd, iyd) = CulG
                     DATACUL1(2, ixd, iyd) = CulR
                     DATACUL1(3, ixd, iyd) = CulA
                  Case 2
                     DATACUL2(0, ixd, iyd) = CulB
                     DATACUL2(1, ixd, iyd) = CulG
                     DATACUL2(2, ixd, iyd) = CulR
                     DATACUL2(3, ixd, iyd) = CulA
                  End Select
               
                  picDATAREG(0, ixd, iyd) = CulB
                  picDATAREG(1, ixd, iyd) = CulG
                  picDATAREG(2, ixd, iyd) = CulR
                  picDATAREG(3, ixd, iyd) = CulA   ' alpha byte
               
                  DealWithSmallNA ixd, iyd   ' Adjust picDATAREG
               
               End If
            End If
            End If
            End If
            End If
         
         Next ixd
      Next iyd
      
   End If

   ' Show rotated image in PIC from picDATAREG()
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = PIC.Width
      .biHeight = PIC.Height
      .biBitCount = 32
   End With
      
   If SetDIBits(PIC.hdc, PIC.Image, _
      0, PIC.Height, picDATAREG(0, 0, 0), BMIH, 0) = 0 Then
      MsgBox "SetDIBits Error", vbCritical, "Rotate"
   End If
   PIC.Picture = PIC.Image
      
   ' and show result on Form1
   With Form1
      BitBlt .picSmall(ImageNum).hdc, 0, 0, _
      .picSmall(ImageNum).Width, .picSmall(ImageNum).Height, PIC.hdc, 0, 0, vbSrcCopy
      .picSmall(ImageNum).Picture = .picSmall(ImageNum).Image
      .DrawGrid
   End With
   
   'Reconcile ImageNum
End Sub

Private Sub cmdACCCan_Click(Index As Integer)
   If Index = 1 Then    ' Cancel/Close
      aTransfer = False
   Else  ' Accept/Close
      aTransfer = True
   End If
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not aTransfer Then  ' Cancel/Close
      cmdReset_Click
   Else  ' aTransfer= True Accept/Close
      PIC.SetFocus
      If Ndeg = 0 Then
         cmdReset_Click
      End If
      Transfer ' BitBlt PIC to Form1.picSMall & picSmallBU & DrawGrid
   End If
   
   Erase DATACULORG(), picDATAORG()
   
   frmRotateLeft = frmRotator.Left
   frmRotateTop = frmRotator.Top
   Unload Me
End Sub

