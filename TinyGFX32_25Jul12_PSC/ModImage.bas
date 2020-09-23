Attribute VB_Name = "ModImage"
'ModImage.bas

Option Explicit

Public Tools As Long
Public Enum ETools
   [Dots] = 0
   [Lines]
   [Boxes]
   [BoxesSolid]
   [BoxesFilled]
   [Ellipse]
   [EllipseSolid]
   [EllipseFilled]
   [Fill]
   [Dropper]
   [Rot90CW]
   [Rot90ACW]
   [Text]
   
   [BlurIm]
   [LeftColorReliefIm]
   [InvertIm]
   [BrighterIm]
   [DarkerIm]
   [BorderIm]
   [ReplaceLbyR]
   
   [SelectON]
   [MoveSEL]
   [MoveCOPY]
   
   [BoxCenShade]
   [BoxVertShade]
   [BoxHorzShade]
   [EllipseCenShade]
   [EllipseVertShade]
   [EllipseHorzShade]
   [BoxDiagTLBRShade]
   [BoxDiagBLTRShade]
   [BoxVertCenShade]
   [BoxHorzCenShade]
   [BoxHVCenShade]
   
   [Tools2Blur]
   [Tools2Grey]
   [Tools2Invert]
   [Tools2Bright]
   [Tools2Dark]
   [Tools2Checker]
   [Tools2HLines]
   [Tools2VLines]
   [Tools2Random]
   [Tools2MoreBlue]
   [Tools2LessBlue]
   [Tools2MoreGreen]
   [Tools2LessGreen]
   [Tools2MoreRed]
   [Tools2LessRed]
   [MirrorL]
   [MirrorR]
   [MirrorT]
   [MirrorB]
   [ReflectLeft]
   [ReflectRight]
   [ReflectTop]
   [ReflectBottom]
   
End Enum

'Public Const TopToolNum As Long = 55

Public BMPFsize As Long
Public BMPOffset As Long
Public FSize As Long

' Ico info
Public NumIcons As Integer
Public IcoLOC(0 To 2) As Long   ' BMIH Location given at locs
                                ' 19, 35 & 51 from BOF Longs
' 40  Long
Public IcoWidth(0 To 2) As Integer
Public IcoHeight(0 To 2) As Long ' 2 x height
Public ICOBPP(0 To 2) As Integer    ' 1,4,8,24,32

''''''''''''''''''''''''''''''''''''''
Public ImBPP(0 To 2) As Long
Public bpp(0 To 2) As Long

' Cursor & multi-ico

Public HotX(0 To 2) As Integer, HotY(0 To 2) As Integer

Public Type IcoCurHeader
   ires As Integer      ' 0
   ityp As Integer      ' 1 ico, 2 cur
   inum As Integer      ' Num of images (1 for single)
End Type                ' Length = 6
Public IcoCurHdr As IcoCurHeader

Public Type IcoCurInformation
   bWidth As Byte
   bHeight As Byte
   Bres1 As Byte        ' NColors (NR)
   Bres2 As Byte        ' 0
   iHotX As Integer     ' 1 planes(NR) ico,   HotX cur
   iHotY As Integer     ' bpp(NR)      ico,   HotY cur
   LSize As Long        ' total size of (each) image
   Loffset As Long      ' Offset to start of (each) BMIH  (22 for single)
End Type                ' Length 16 B
Public IcoCurInfo(2) As IcoCurInformation

Public Type IcoCurBIH
   LBIH As Long         ' size of BMIH (40)
   LWidth As Long       ' width
   LHeight As Long      ' 2 * height
   inump As Integer     ' num planes = 1
   ibpp As Integer      ' bpp (1,4,8,24)
   L1 As Long           ' 0
   L2 As Long           ' 0
   L3 As Long           ' 0
   L4 As Long           ' 0
   L5 As Long           ' 0
   L6 As Long           ' 0
End Type                ' Length 40 B
Public IcoCurBMIH(2) As IcoCurBIH

Public XORARR() As Byte
Public ANDARR() As Byte

Public XORW(2)    As Long, ANDW(2)    As Long
Public XORSize(2) As Long, ANDSize(2) As Long

Public picSmallDATA() As Long ' (0 To W-1, 0 to H-1)
Public picDummy() As Long

''TEST
'Private picDATA() As Byte


Public picDATAREG() As Byte ' (0 to 3, 0 To W-1, 0 to H-1)
Public picDATAORG() As Byte

Public DATACULSRC() As Byte  ' (0 to 3, 0 To W-1, 0 to H-1)
Public DATACUL0() As Byte
Public DATACUL1() As Byte
Public DATACUL2() As Byte

Public DATACULBU() As Byte
Public DATACULORG() As Byte


Public PaletteSize(2) As Long

Public LColor As Long, RColor As Long
' Transparent colors
Public TColorBGR As Long   ' BGR
Public TColorRGB As Long  ' RGB
Public picMaskDATA() As Long

' For showing with other transparent colors
Public VisColor As Long   ' Test different TColorBGR

' Alpha
Public aAlphaEdit As Boolean
Public EditCreate As Long   ' 0,1

Public BMPbpp As Integer

Public ICOCURBPP() As Integer
Public Ico36bpp As Byte

Public svOptimizeNumber As Long

Public Response$   ' For multi-select icons
Public TempImageNum As Long  ' General back up Image number

' BackUp
Public SAVDAT(0 To 2) As Long  ' picSmall(ImageNum) Width & Height, ImBPP(ImageNum)
Public CurrentGen(0 To 3) As Long

Public SaveIndex As Integer    ' General Index back up

' Effects
Public aEffects As Boolean

' Select Rectangle params
Public aSelect As Boolean
Public aSelectDrawn As Boolean

Public aMoveSEL As Boolean
Public LOX As Long, HIX As Long
Public LOY As Long, HIY As Long


'#### File Checker ####
Public Sub CheckFile(FSpec$, Message$)
Dim FFileNum As Integer
Dim IntegerMarker(1 To 2) As Integer
Dim LongMarker As Long
Dim IcoHdrWidth As Byte
Dim IcoHdrHeight As Byte
Dim k As Long
Dim Exten$
   Message$ = ""
   FFileNum = FreeFile
   Open FSpec$ For Binary As #FFileNum
   FSize = LOF(FFileNum)
   Exten$ = UCase$(FindExtension$(FSpec$))
   Select Case Exten$
   Case "BMP"
      Get #FFileNum, , IntegerMarker(1) ' 19778 = 42 4Dh BM
      If IntegerMarker(1) <> 19778 Then Message$ = "Invalid bmp file "
      ' BMPbpp
      Get #FFileNum, 29, IntegerMarker(2)
      BMPbpp = IntegerMarker(2)
   Case "JPG", "JPEG"
      Get #FFileNum, , IntegerMarker(1)            ' -9985 = FF D8 BEndian
      Get #FFileNum, FSize - 1, IntegerMarker(2)   ' FF D9 BEndian
      If IntegerMarker(1) <> -9985 Then Message$ = "Invalid jpeg file, FF D8 marker 1 wrong ": GoTo CloseIt
' NOT CRITICAL ??
'     If IntegerMarker(2) <> -9729 Then Message$ = "Invalid jpeg file, FF D9 wrong End Marker ": GoTo CloseIt
   Case "GIF"
      Message$ = Space$(6)
      Get #FFileNum, , Message$ ' GIF87a or GIF89a
      If Message$ <> "GIF87a" And Message$ <> "GIF89a" Then
         Message$ = "Invalid gif file ": GoTo CloseIt
      Else
         Message$ = ""
      End If
' NOT CRITICAL ??
'      Message$ = Space$(1)
'      Get #FFileNum, FSize, Message$    ' ; 3B
'      If Message$ <> ";" Then
'         Message$ = "Invalid gif file, wrong End Marker (ie not 3B)": GoTo CloseIt
'      Else
'         Message$ = ""
'      End If
   Case "ICO", "CUR"
      Get #FFileNum, , IntegerMarker(1)    ' 0
      Get #FFileNum, , IntegerMarker(2)    ' 1 or 2
      Get #FFileNum, , IntegerMarker(2)    ' Num icons
      If IntegerMarker(1) <> 0 Then
         Message$ = "Invalid ico, cur file (First byte <> 0) "
      Else
         NumIcons = IntegerMarker(2)
         
         Get #FFileNum, , IcoHdrWidth     ' Need to cater for NumIcons>1
         Get #FFileNum, , IcoHdrHeight
         
         ' ETools2ct to max of 3 icons   IcoLOC() start of individual ico BHI headers
         ' 19, 19+16 = 35, 35+16 = 51
         Get #FFileNum, 19, IcoLOC(0)
         If NumIcons >= 2 Then Get #FFileNum, 35, IcoLOC(1)
         If NumIcons >= 3 Then Get #FFileNum, 51, IcoLOC(2)
         
         For k = 0 To NumIcons - 1
            ' These rely on correct icon header
            Select Case k
            Case 0: Get #FFileNum, 19 - 12, IcoHdrWidth
                    Get #FFileNum, 19 - 11, IcoHdrHeight
            Case 1: Get #FFileNum, 35 - 12, IcoHdrWidth
                    Get #FFileNum, 35 - 11, IcoHdrHeight
            Case 2: Get #FFileNum, 51 - 12, IcoHdrWidth
                    Get #FFileNum, 51 - 11, IcoHdrHeight
            End Select
            
            IcoLOC(k) = IcoLOC(k) + 1
            
            Get #FFileNum, IcoLOC(k), LongMarker
            If LongMarker <> 40 Then Message$ = "Error in ICO file (BMIH <> 40) or Vista icon": GoTo CloseIt
            
            Get #FFileNum, IcoLOC(k) + 4, IcoWidth(k)
            
            Get #FFileNum, IcoLOC(k) + 8, IcoHeight(k)
            
            Get #FFileNum, IcoLOC(k) + 14, ICOBPP(k)
            
            If IcoWidth(k) <> IcoHdrWidth Or _
               IcoHeight(k) <> 2 * IcoHdrHeight Then
                  Message$ = "Mismatch on widths & heights or" & vbCrLf & _
                  "wrong icon header or" & vbCrLf & _
                  "256x256 icon." & vbCrLf & "Could try Extractor menu."
                  GoTo CloseIt
            End If
            
            Select Case ICOBPP(k)
            Case 1, 4, 8, 24, 32
            Case Else
               Message$ = "Error in ICO file (BPP)": GoTo CloseIt
            End Select
            
            If k = 2 Then Exit For
            
         Next k
         k = k - 1
      End If
   
   Case Else
      Message$ = "Wrong file type"
   End Select
CloseIt:
   Close #FFileNum
End Sub

Public Function ETools2ctICON(ICONum As Long, FileStream() As Byte, IconStream() As Byte) As Boolean
' Only single or multiple icons or cursor come here
' ETools2ct icon ICONum from FileStream() into a single icon image in IconStream()
' Input:  ICONum (0,1 or 2), FileStream() = BStream() from caller)
' if 32bpp will contain BGRA XOR bytes
' Output: IconStream()
Dim a$
Dim k As Long
Dim ix As Long, iy As Long
Dim imSize(0 To 3) As Long
Dim imOffset(0 To 3) As Long
Dim BStream() As Byte
Dim NB0 As Byte, NG0 As Byte, NR0 As Byte
Dim NB As Byte, NG As Byte, NR As Byte, NA As Byte
Dim TB As Byte, TG As Byte, TR As Byte
Dim Ealpha As Single
Dim kk As Long

On Error GoTo ETools2ctICONERROR
   ' IcoWidth(ICONum),IcoHeight(ICONum) from CheckFile
   If IcoWidth(ICONum) > MaxWidth Or IcoHeight(ICONum) > 2 * MaxHeight Then
     a$ = "Icon width/height > Max for icon number " & Str$(ICONum + 1) & vbCr
     a$ = a$ & "(>" & Str$(MaxWidth) & " x" & Str$(MaxHeight) & " )"
     MsgBox a$, vbInformation, "Loading Icon file"
     Exit Function
   End If
   
   Select Case ICONum
   Case 0
         imSize(0) = FileStream(14) + 256& * FileStream(15) + 65536 * FileStream(16)
         imOffset(0) = FileStream(18) + 256& * FileStream(19) + 65536 * FileStream(20)
   Case 1
         imSize(1) = FileStream(30) + 256& * FileStream(31) + 65536 * FileStream(32)
         imOffset(1) = FileStream(34) + 256& * FileStream(35) + 65536 * FileStream(36)
   Case 2
         imSize(2) = FileStream(46) + 256& * FileStream(47) + 65536 * FileStream(48)
         imOffset(2) = FileStream(50) + 256& * FileStream(51) + 65536 * FileStream(52)
   End Select

   ReDim IconStream(0 To 6 + 16 + imSize(ICONum) - 1)
   
   For k = 0 To 5
     IconStream(k) = FileStream(k)
   Next k
   IconStream(4) = 1 ' 1 icon
   For k = 6 To 21
     IconStream(k) = FileStream(k + 16 * ICONum)
   Next k
   IconStream(18) = 22   ' Offset for single icon
   For k = 19 To 21      ' zero rest of offset bytes
     IconStream(k) = 0
   Next k
   For k = 22 To 22 + imSize(ICONum) - 1
     IconStream(k) = FileStream(imOffset(ICONum) + (k - 22))
   Next k
   
   k = IconStream(26) ' W
   kk = IconStream(30) 'H
   
   Ico36bpp = IconStream(36)
   
   'Ignore checksize
   'If Not Checksize(imSize(ICONum), k, kk \ 2, Ico36bpp) Then
   '  ETools2ctICON = False
   '  Exit Function
   'End If
   
   
   ' If not 32bpp then DONE here!
   
   If IconStream(36) = 32 Then ' Convert 32bpp XOR to 32bpp BMP
                'so that PictureFromByteStream(IcoStream) works
'Public BMPFsize As Long
'Public BMPOffset As Long
      'BM    14
      'BMHI  +40
      '+ 32x32x4  if 32x32 BGRA
      BMPFsize = 14 + 40 + (IcoWidth(ICONum) * IcoHeight(ICONum) \ 2) * 4
      BMPOffset = 54  ' Start of XOR
      
      ReDim BStream(0 To BMPFsize - 1)
      BStream(0) = 66
      BStream(1) = 77
      LngToRGB BMPFsize, BStream(2), BStream(3), BStream(4)
      BStream(6) = 0
      BStream(7) = 0
      BStream(8) = 0
      BStream(9) = 0
      LngToRGB BMPOffset, BStream(10), BStream(11), BStream(12)
      ' Halve ico height for BMP
      IconStream(30) = CByte(Val(IconStream(30)) \ 2)
      For k = 0 To (IcoWidth(ICONum) * IcoHeight(ICONum) / 2) * 4 + 39 ' 40-1
         BStream(k + 14) = IconStream(k + 22)
      Next k
      k = k - 1
      ReDim DATACULSRC(0 To 3, 0 To (IcoWidth(ICONum) - 1), 0 To IcoHeight(ICONum) / 2 - 1)
      
      iy = 0
      ix = 0
      ' Modify BStream(54 To EOF)
      ' For alpha
      LngToRGB TColorBGR, TR, TG, TB
      
      For k = 54 To UBound(BStream) Step 4
         NB0 = CByte(BStream(k))
         NG0 = CByte(BStream(k + 1))
         NR0 = CByte(BStream(k + 2))
         NA = CByte(BStream(k + 3))
         
         ' Original unchanged colors
         DATACULSRC(0, ix, iy) = NB0
         DATACULSRC(1, ix, iy) = NG0
         DATACULSRC(2, ix, iy) = NR0
         DATACULSRC(3, ix, iy) = NA
         
         If aAlphaRestricted Then
            ' Only lets colors through where NA=255
            If NA <> 255 Then NA = 0
            DATACULSRC(3, ix, iy) = NA
         End If
         
         ix = ix + 1
         If ix > IcoWidth(ICONum) - 1 Then
            ix = 0
            iy = iy + 1
         End If
               
         Ealpha = (NA / 255)
         NB = TB * (1 - Ealpha) + NB0 * Ealpha
         NG = TG * (1 - Ealpha) + NG0 * Ealpha
         NR = TR * (1 - Ealpha) + NR0 * Ealpha
      
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
      
      ReDim IconStream(0 To UBound(BStream))
      
      For k = 0 To UBound(BStream)
         IconStream(k) = BStream(k)   ' IconStream() for PictureFromByteStream(IcoStream)
      Next k
      'IconStream(36) = Ico36bpp
   End If
   ' Could use CopyMemory but fast enough and more descriptive as is
   ETools2ctICON = True
   Exit Function
'=========
ETools2ctICONERROR:
   MsgBox "ETools2ct Icon Error", vbCritical, "EXTRACT ICON"
   
End Function


Public Sub ShowMask(picS As PictureBox, picM As PictureBox)
' picS = picSmall(ImageNum) or picSmallEdit non-TColorBGR show as black
'        on picM (also picSmall(ImageNum)) or picSmallMask.
Dim sh As Long, sw As Long
Dim ix As Long, iy As Long
Dim Cul As Long

   sw = picS.Width
   sh = picS.Height
   For iy = 0 To sh - 1
   For ix = 0 To sw - 1
      Cul = picS.Point(ix, iy)
      If Cul >= 0 Then
         If Cul <> TColorBGR Then
            SetPixelV picM.hdc, ix, iy, 0
         Else
            SetPixelV picM.hdc, ix, iy, vbWhite
         End If
      End If
   Next ix
   Next iy
End Sub

Public Sub ShowAlpha(ImN As Long, picS As PictureBox, picM As PictureBox)
' IMN=ImageNum
' picS = picSmall(ImageNum) non-TColorBGR show as black
'        on picM (also picSmall(ImageNum)). Could add picMask(ImageNum).
' From Alpha Edit pics=picM=picSmallAlpha

Dim sh As Long, sw As Long
Dim ix As Long, iy As Long
Dim SR As Byte

   sw = picS.Width
   sh = picS.Height
   
   picM.Picture = LoadPicture
   picM.Width = sw
   picM.Height = sh
   picM.BackColor = vbWhite
   For iy = 0 To sh - 1
   For ix = 0 To sw - 1
      Select Case ImN
      Case 0
         SR = 255 - DATACUL0(3, ix, iy)
      Case 1
         SR = 255 - DATACUL1(3, ix, iy)
      Case 2
         SR = 255 - DATACUL2(3, ix, iy)
      End Select
      SetPixelV picM.hdc, ix, sh - iy - 1, RGB(SR, SR, SR)
   Next ix
   Next iy
End Sub


Public Sub ShowWithTColor(ImN As Long, picS As PictureBox, picM As PictureBox, Optional OV As Long = 0)
' picS = picSmall(ImageNum) non-TColorBGR show as original and TColorBGR as Test TColorBGR
'        from MouseDown on picVisColor
'        on picM (also picSmall(ImageNum)). Could add picM(ImageNum).
' From Alpha pics & picm = picSmallColors
Dim sh As Long, sw As Long
Dim BMIH As BITMAPINFOHEADER
Dim ix As Long, iy As Long
Dim AR As Byte, AG As Byte, ab As Byte
Dim VisR As Byte, VisG As Byte, VisB As Byte
Dim OrgR As Byte, OrgG As Byte, OrgB As Byte
Dim OrgRo As Byte, OrgGo As Byte, OrgBo As Byte
'Dim TColorRGB As Long
Dim Cul As Long
Dim Salpha As Single

   sw = picS.Width
   sh = picS.Height
   
   aDIBError = False
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = sw 'ImageWidth(ImageNum)
      .biHeight = sh  'ImageHeight(ImageNum)
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With

   ReDim picSmallDATA(0 To sw - 1, 0 To sh - 1)
   If GetDIBits(Form1.hdc, picS.Image, 0, sh, picSmallDATA(0, 0), BMIH, 0) = 0 Then
      MsgBox "DIB ERROR", vbCritical, "ShowWithTColor"
   End If
   
   picM.Picture = LoadPicture
   picM.Width = sw
   picM.Height = sh
   'picM.BackColor = vbWhite
   
   'LngToRGB TColorBGR, TR, TG, TB   ' Transparent colors  NU
   'TColorRGB = RGB(TB, TG, TR)
   If ImBPP(ImN) <> 32 Then
      LngToRGB VisColor, VisR, VisG, VisB ' Visbility colors
      For iy = 0 To sh - 1
      For ix = 0 To sw - 1
         Cul = picSmallDATA(ix, iy) And &HFFFFFF  ' NB 4byte nums can be -ve in VB
         If Cul = TColorRGB Then
            SetPixelV picM.hdc, ix, sh - iy - 1, VisColor
         Else
            LngToRGB Cul, AR, AG, ab   ' Actual displayed colors
            SetPixelV picM.hdc, ix, sh - iy - 1, RGB(ab, AG, AR)
         End If
      Next ix
      Next iy
   
   Else  ' ImBPP(ImN) = 32
   
      LngToRGB VisColor, VisR, VisG, VisB ' Visbility colors
      'If OV = 1 Then VisColor = TColorRGB
      For iy = 0 To sh - 1
      For ix = 0 To sw - 1
         Select Case ImN
         Case 0:
            Salpha = DATACUL0(3, ix, iy) / 255
            OrgBo = DATACUL0(0, ix, iy)
            OrgGo = DATACUL0(1, ix, iy)
            OrgRo = DATACUL0(2, ix, iy)
         Case 1: Salpha = DATACUL1(3, ix, iy) / 255
            OrgBo = DATACUL1(0, ix, iy)
            OrgGo = DATACUL1(1, ix, iy)
            OrgRo = DATACUL1(2, ix, iy)
         Case 2: Salpha = DATACUL2(3, ix, iy) / 255
            OrgBo = DATACUL2(0, ix, iy)
            OrgGo = DATACUL2(1, ix, iy)
            OrgRo = DATACUL2(2, ix, iy)
         End Select
         If OV = 0 Then  ' ie when Visibilty is pressed
               OrgR = VisR * (1 - Salpha) + OrgRo * Salpha
               OrgG = VisG * (1 - Salpha) + OrgGo * Salpha
               OrgB = VisB * (1 - Salpha) + OrgBo * Salpha
               SetPixelV picM.hdc, ix, sh - iy - 1, RGB(OrgR, OrgG, OrgB)
            ' So if salpha=0 VisRGB shows else a mixture
            ' between VisRGB and OrgRGB
         Else
               SetPixelV picM.hdc, ix, sh - iy - 1, RGB(OrgRo, OrgGo, OrgBo)
         End If
      Next ix
      Next iy
   End If
   If OV = 0 Then
      Form1.picSmallFrame.BackColor = VisColor
   End If
End Sub

Public Sub TransferSmallToLarge(picS As PictureBox, picL As PictureBox, Optional xs As Single = 0, Optional ys As Single = 0)
 
 ' Form1.DrawGrid:  TransferSmallToLarge picSmall(ImageNum), picPANEL
 ' frmRotator:   eg TransferSmallToLarge picSmallEdit, picEdit

Dim sh As Long, sw As Long
Dim BMIH As BITMAPINFOHEADER
   
   sw = picS.Width
   sh = picS.Height
   picL.Picture = LoadPicture
   
   aDIBError = False
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = sw 'ImageWidth(ImageNum)
      .biHeight = sh  'ImageHeight(ImageNum)
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With
   SetStretchBltMode picL.hdc, COLORONCOLOR

   ReDim picSmallDATA(0 To sw - 1, 0 To sh - 1)
   If GetDIBits(Form1.hdc, picS.Image, 0, sh, picSmallDATA(0, 0), BMIH, 0) = 0 Then
      MsgBox "DIB ERROR", vbCritical, "TransferSmallToLarge"
   End If
   
   StretchDIBits picL.hdc, xs, ys, picL.Width, picL.Height, 0, 0, _
      sw, sh, picSmallDATA(0, 0), BMIH, 0, vbSrcCopy

'Dim Cul As Long
'Cul = picSmallDATA(0, 0) And &HFFFFFF
' Equiv to
'Public GridMult As Long
'Dim kx As Long, ky As Long
'Dim xleft As Long, xright As Long
'Dim ytop As Long, ybelow As Long
'Dim pDrawCul As Long

'   For ky = 0 To ImageHeight(ImageNum) - 1
'   For kx = 0 To ImageWidth(ImageNum) - 1
'      pDrawCul = picS.Point(kx, ky)
'      If pDrawCul >= 0 Then
'            xleft = kx * GridMult + 1
'            xright = xleft + GridMult - 2
'            ytop = ky * GridMult + 1
'            ybelow = ytop + GridMult - 2
'            picL.Line (xleft, ytop)-(xright, ybelow), pDrawCul, BF
'      End If
'   Next kx
'   Next ky
End Sub

Public Sub TransferpicSEL(picS As PictureBox, picL As PictureBox, Optional xs As Single = 0, Optional ys As Single = 0)
 ' picS = picSmall(ImageNum) To picL = picPANEL [Used in Sub DrawGrid on Form1]
 ' or picSEL(ImageNum) to picPANEL [From picPANEL_MouseMove Case [MoveSEL][MoveCopy]
Dim sh As Long, sw As Long
Dim BMIH As BITMAPINFOHEADER
   
   sw = picS.Width
   sh = picS.Height
   'picL.Picture = LoadPicture
   
   aDIBError = False
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = sw 'ImageWidth(ImageNum)
      .biHeight = sh  'ImageHeight(ImageNum)
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With
   SetStretchBltMode picL.hdc, COLORONCOLOR

   ReDim picSmallDATA(0 To sw - 1, 0 To sh - 1)
   If GetDIBits(Form1.hdc, picS.Image, 0, sh, picSmallDATA(0, 0), BMIH, 0) = 0 Then
      MsgBox "DIB ERROR", vbCritical, "TransferpicSEL"
   End If
   
   StretchDIBits picL.hdc, xs, ys, sw * GridMult, sh * GridMult, 0, 0, _
      sw, sh, picSmallDATA(0, 0), BMIH, 0, vbSrcCopy

End Sub

Public Sub GetOriginalColors(ImN As Long)
' From mnuSaveBMP & ProcessOutFile
' and ProcessOutFile is called from
'  mnuSaveCURImage_Click
'  mnuSaveICOImage_Click
'  SaveSingleICOorCUR
'  SaveMultiICOs

' ImN = ImageNum AND ImBPP() = 32 only
   ReDim picDATAREG(0 To 3, 0 To ImageWidth(ImN) - 1, 0 To ImageHeight(ImN) - 1)
   Select Case ImN
   Case 0
      FILL3D picDATAREG(), DATACUL0()
   Case 1
      FILL3D picDATAREG(), DATACUL1()
   Case 2
      FILL3D picDATAREG(), DATACUL2()
   End Select
End Sub

'Public Sub GetPointOrgColor(ix As Long, iy As Long, OrgR As Byte, OrgG As Byte, OrgB As Byte)
'Dim AByte As Byte
'Dim SR As Byte, SG As Byte, SB As Byte
'Dim bAlpha As Single
'Dim sCorg As Single
'
'      LngToRGB TColorBGR, SR, SG, SB
'
'      bAlpha = AByte / 255
'      If bAlpha = 0 Or bAlpha = 1 Then
'         OrgB = picDATAREG(0, ix, iy)
'         OrgG = picDATAREG(1, ix, iy)
'         OrgR = picDATAREG(2, ix, iy)
'      Else
'         sCorg = picDATAREG(0, ix, iy) - SB * (1 - bAlpha)
'         sCorg = sCorg / bAlpha
'         If sCorg > 255 Then sCorg = 255
'         If sCorg < 0 Then sCorg = 0
'         OrgB = sCorg
'         sCorg = picDATAREG(1, ix, iy) - SG * (1 - bAlpha)
'         sCorg = sCorg / bAlpha
'         If sCorg > 255 Then sCorg = 255
'         If sCorg < 0 Then sCorg = 0
'         OrgG = sCorg
'         sCorg = picDATAREG(2, ix, iy) - SR * (1 - bAlpha)
'         sCorg = sCorg / bAlpha
'         If sCorg > 255 Then sCorg = 255
'         If sCorg < 0 Then sCorg = 0
'         OrgR = sCorg
'      End If
'' Keep original
''      picDATAREG(0, ix, iy) = OrgB
''      picDATAREG(1, ix, iy) = OrgG
''      picDATAREG(2, ix, iy) = OrgR
''      picDATAREG(3, ix, iy) = ABYTE
'End Sub

Public Sub GetTheBitsBGR(ImNum As Long, PIC As PictureBox)
' For BMP(RGB), GIF(RGB), ICO & CUR [Sub mnuSaveBMPImage & Sub PreProcess on Form1]
' Input: ImNum = ImageNum, picSmall(ImageNum) from Form1
' Output: to Public picDATAREG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
Dim BMIH As BITMAPINFOHEADER
   aDIBError = False
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = ImageWidth(ImNum)
      .biHeight = ImageHeight(ImNum)
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With
   ReDim picDATAREG(0 To 3, 0 To ImageWidth(ImNum) - 1, 0 To ImageHeight(ImNum) - 1)
   If GetDIBits(PIC.hdc, PIC.Image, 0, ImageHeight(ImNum), picDATAREG(0, 0, 0), BMIH, 0) = 0 Then
      MsgBox "DIB ERROR", vbCritical, "Sub GetTheBitsBGR"
      aDIBError = True
   End If
'TEST
'Dim ABYTE As Byte
'Dim ix As Long, iy As Long
'For iy = 0 To ImageHeight(ImageNum) - 1
'For ix = 0 To ImageWidth(ImageNum) - 1
'   ABYTE = picDATAREG(3, ix, iy)
'   If ABYTE <> 255 Then Stop
'Next ix
'Next iy
End Sub


Public Sub GetTheBitsLong(PIC As PictureBox)
' Input:  frm as Form1 or frmText, PIC = picSmall(ImageNum) from Form1 or picSmallText from frmText
' Output: to Public picSmallDATA(0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
'         for picSmall(ImageNum) or picSmallText

Dim BMIH As BITMAPINFOHEADER
   aDIBError = False
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = ImageWidth(ImageNum)
      .biHeight = ImageHeight(ImageNum)
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With
   ReDim picSmallDATA(0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   If GetDIBits(PIC.hdc, PIC.Image, 0, ImageHeight(ImageNum), picSmallDATA(0, 0), BMIH, 0) = 0 Then
      MsgBox "DIB ERROR", vbCritical, "Sub GetTheBitsLong"
      aDIBError = True
   End If
End Sub


Public Sub Rotate90(Index As Integer)
' Input : picSmallDATA(0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
' Output: picSmallDATA() Rotated 90 deg
Dim ix As Long, iy As Long
   
      ' Selection not allowed - gets cancelled
      ' Swap width & height
      ReDim picDummy(0 To ImageHeight(ImageNum) - 1, 0 To ImageWidth(ImageNum) - 1)
      ReDim DATACULSRC(0 To 3, 0 To ImageHeight(ImageNum) - 1, 0 To ImageWidth(ImageNum) - 1)
      
      For iy = 0 To ImageHeight(ImageNum) - 1
      For ix = 0 To ImageWidth(ImageNum) - 1
         If Index = [Rot90CW] Then ' Rotate 90 clockwise
             ' +90
            picDummy(iy, ImageWidth(ImageNum) - ix - 1) = picSmallDATA(ix, iy)
            
            If ImBPP(ImageNum) = 32 Then
               Select Case ImageNum
               Case 0
                  DATACULSRC(0, iy, ImageWidth(ImageNum) - ix - 1) = _
                     DATACUL0(0, ix, iy)
                  DATACULSRC(1, iy, ImageWidth(ImageNum) - ix - 1) = _
                     DATACUL0(1, ix, iy)
                  DATACULSRC(2, iy, ImageWidth(ImageNum) - ix - 1) = _
                     DATACUL0(2, ix, iy)
                  DATACULSRC(3, iy, ImageWidth(ImageNum) - ix - 1) = _
                     DATACUL0(3, ix, iy)
               Case 1
                  DATACULSRC(0, iy, ImageWidth(ImageNum) - ix - 1) = _
                     DATACUL1(0, ix, iy)
                  DATACULSRC(1, iy, ImageWidth(ImageNum) - ix - 1) = _
                     DATACUL1(1, ix, iy)
                  DATACULSRC(2, iy, ImageWidth(ImageNum) - ix - 1) = _
                     DATACUL1(2, ix, iy)
                  DATACULSRC(3, iy, ImageWidth(ImageNum) - ix - 1) = _
                     DATACUL1(3, ix, iy)
               Case 2
                  DATACULSRC(0, iy, ImageWidth(ImageNum) - ix - 1) = _
                     DATACUL2(0, ix, iy)
                  DATACULSRC(1, iy, ImageWidth(ImageNum) - ix - 1) = _
                     DATACUL2(1, ix, iy)
                  DATACULSRC(2, iy, ImageWidth(ImageNum) - ix - 1) = _
                     DATACUL2(2, ix, iy)
                  DATACULSRC(3, iy, ImageWidth(ImageNum) - ix - 1) = _
                     DATACUL2(3, ix, iy)
               End Select
            End If
         ElseIf Index = [Rot90ACW] Then  ' Rotate 90 anti-clockwise
            ' -90
            picDummy(ImageHeight(ImageNum) - iy - 1, ix) = picSmallDATA(ix, iy)
            
            If ImBPP(ImageNum) = 32 Then
               Select Case ImageNum
               Case 0
                  DATACULSRC(0, ImageHeight(ImageNum) - iy - 1, ix) = _
                     DATACUL0(0, ix, iy)
                  DATACULSRC(1, ImageHeight(ImageNum) - iy - 1, ix) = _
                     DATACUL0(1, ix, iy)
                  DATACULSRC(2, ImageHeight(ImageNum) - iy - 1, ix) = _
                     DATACUL0(2, ix, iy)
                  DATACULSRC(3, ImageHeight(ImageNum) - iy - 1, ix) = _
                     DATACUL0(3, ix, iy)
               Case 1
                  DATACULSRC(0, ImageHeight(ImageNum) - iy - 1, ix) = _
                     DATACUL1(0, ix, iy)
                  DATACULSRC(1, ImageHeight(ImageNum) - iy - 1, ix) = _
                     DATACUL1(1, ix, iy)
                  DATACULSRC(2, ImageHeight(ImageNum) - iy - 1, ix) = _
                     DATACUL1(2, ix, iy)
                  DATACULSRC(3, ImageHeight(ImageNum) - iy - 1, ix) = _
                     DATACUL1(3, ix, iy)
               Case 2
                  DATACULSRC(0, ImageHeight(ImageNum) - iy - 1, ix) = _
                     DATACUL2(0, ix, iy)
                  DATACULSRC(1, ImageHeight(ImageNum) - iy - 1, ix) = _
                     DATACUL2(1, ix, iy)
                  DATACULSRC(2, ImageHeight(ImageNum) - iy - 1, ix) = _
                     DATACUL2(2, ix, iy)
                  DATACULSRC(3, ImageHeight(ImageNum) - iy - 1, ix) = _
                     DATACUL2(3, ix, iy)
               End Select
            End If
         End If
      Next ix
      Next iy
      
      picSmallDATA() = picDummy()
      
      If ImBPP(ImageNum) = 32 Then
         Select Case ImageNum
         Case 0
            ReDim DATACUL0(0 To 3, 0 To ImageHeight(ImageNum) - 1, 0 To ImageWidth(ImageNum) - 1)
            FILL3D DATACUL0(), DATACULSRC()
         Case 1
            ReDim DATACUL1(0 To 3, 0 To ImageHeight(ImageNum) - 1, 0 To ImageWidth(ImageNum) - 1)
            FILL3D DATACUL1(), DATACULSRC()
         Case 2
            ReDim DATACUL2(0 To 3, 0 To ImageHeight(ImageNum) - 1, 0 To ImageWidth(ImageNum) - 1)
            FILL3D DATACUL2(), DATACULSRC()
         End Select
      End If
   
   Erase picDummy
   
   ' ImageHeight(ImageNum) & ImageWidth(ImageNum) gets swapped on return
End Sub


'#### EFFECTS #############

Public Sub Replace_Alpha_DATACUL(ix As Long, iy As Long, _
   AByt As Byte, ByVal RB As Byte, ByVal RG As Byte, ByVal RR As Byte, _
   Optional ChangeAlpha As Long = 0)
Dim ab As Byte

' ChangeAlpha
' 0  Blur ALL
' 1  Blur image only
' 2  Blur Alpha only

' Called by some Effects routines
   Select Case ImageNum
   Case 0
      
      Select Case ChangeAlpha
      Case 0
         DATACUL0(0, ix, iy) = RB
         DATACUL0(1, ix, iy) = RG
         DATACUL0(2, ix, iy) = RR
         DATACUL0(3, ix, iy) = AByt
      Case 1
         DATACUL0(0, ix, iy) = RB
         DATACUL0(1, ix, iy) = RG
         DATACUL0(2, ix, iy) = RR
         'DATACUL0(3, ix, iy) unchanged
      Case 2
         If AByt < 150 Then
            DATACUL0(3, ix, iy) = AByt
         End If
      End Select
   
   Case 1
      
      Select Case ChangeAlpha
      Case 0
         DATACUL1(0, ix, iy) = RB
         DATACUL1(1, ix, iy) = RG
         DATACUL1(2, ix, iy) = RR
         DATACUL1(3, ix, iy) = AByt
      Case 1
         DATACUL1(0, ix, iy) = RB
         DATACUL1(1, ix, iy) = RG
         DATACUL1(2, ix, iy) = RR
         'DATACUL1(3, ix, iy) unchanged
      Case 2
         If AByt < 150 Then
            DATACUL1(3, ix, iy) = AByt
         End If
      End Select
   
   Case 2
      
      Select Case ChangeAlpha
      Case 0
         DATACUL2(0, ix, iy) = RB
         DATACUL2(1, ix, iy) = RG
         DATACUL2(2, ix, iy) = RR
         DATACUL2(3, ix, iy) = AByt
      Case 1
         DATACUL2(0, ix, iy) = RB
         DATACUL2(1, ix, iy) = RG
         DATACUL2(2, ix, iy) = RR
         'DATACUL2(3, ix, iy) unchanged
      Case 2
         If AByt < 200 Then
            DATACUL2(3, ix, iy) = AByt
         End If
      End Select
   
   End Select


End Sub

Public Function Blur(picS As PictureBox, Optional ChangeAlpha As Long = 0) As Boolean
' picS = picSmall(ImageNum)
Dim Cul As Long, Cul2 As Long
Dim ix As Long, iy As Long
Dim RR As Long, RG As Long, RB As Long, ab As Long
Dim r2 As Long, g2 As Long, b2 As Long ', ab As Long
Dim n As Long
' Offsets for edges & corners
Dim pmx1 As Long, pmy1 As Long
Dim pmx2 As Long, pmy2 As Long

   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()

   ReDim picDATAORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   picDATAORG() = picDATAREG()
   
   If ImBPP(ImageNum) = 32 Then
      ReDim DATACULORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
      Select Case ImageNum
      Case 0: FILL3D DATACULORG(), DATACUL0()  ' dest, src
      Case 1: FILL3D DATACULORG(), DATACUL1()
      Case 2: FILL3D DATACULORG(), DATACUL2()
      End Select
   End If
   ' 0 1 0
   ' 1 8 1
   ' 0 1 0
   Blur = True
   
   For iy = LOY + 1 To HIY - 1
   For ix = LOX + 1 To HIX - 1
      If ImBPP(ImageNum) <> 32 Then
         Cul = RGB(picDATAORG(2, ix, iy), picDATAORG(1, ix, iy), picDATAORG(0, ix, iy))
         Cul = Cul And &HFFFFFF
         If Cul <> TColorBGR Then
            BlurPicDATAREG ix, iy   ' Blur picDATAREG()
         End If
      Else   ' IMBPP(ImageNum)=32
         ' ChangeAlpha
         ' 0  Blur ALL
         ' 1  Blur image only
         ' 2  Blur Alpha only
         ' Blur from DATACULORG() & fill DATACUL#() & picDATAREG() with new convolved colors
         BlurDATACUL ix, iy, ChangeAlpha
      End If
   Next ix
   Next iy
   
   ' Vert edges
   ix = LOX
   pmx1 = 0: pmy1 = 1
   For n = 0 To 1
      For iy = LOY To HIY - 1
         EdgeBlur ix, iy, pmx1, pmy1, ChangeAlpha
      Next iy
      ix = HIX
   Next n
   
   ' Horz Edges
   iy = LOY
   pmx1 = 1: pmy1 = 0
   For n = 0 To 1
      For ix = LOX To HIX - 1
         EdgeBlur ix, iy, pmx1, pmy1, ChangeAlpha
      Next ix
      iy = HIY
   Next n
   
If ImageWidth(ImageNum) = 1 Or ImageHeight(ImageNum) = 1 Then Exit Function

   ' Corners
   ' BL
   ix = LOX: iy = LOY
   pmx1 = 1: pmy1 = 0
   pmx2 = 0: pmy2 = 1
   CornerBlur ix, iy, pmx1, pmy1, pmx2, pmy2, ChangeAlpha
   ' TL
   ix = LOX: iy = HIY
   pmx1 = 1: pmy1 = 0
   pmx2 = 0: pmy2 = -1
   CornerBlur ix, iy, pmx1, pmy1, pmx2, pmy2, ChangeAlpha
   ' BR
   ix = HIX: iy = LOY
   pmx1 = -1: pmy1 = 0
   pmx2 = 0: pmy2 = 1
   CornerBlur ix, iy, pmx1, pmy1, pmx2, pmy2, ChangeAlpha
   ' TR
   ix = HIX: iy = HIY
   pmx1 = -1: pmy1 = 0
   pmx2 = 0: pmy2 = -1
   CornerBlur ix, iy, pmx1, pmy1, pmx2, pmy2, ChangeAlpha

'   Following done in Tools =  Blur on Form1
'   DisplayEffects Index  ie
'   SetDIBits picSmall(ImageNum).hDC, picSmall(ImageNum).Image, _
'      0, ImageHeight(ImageNum), picDATAREG(0, 0, 0), BMIH, 0
'   picSmall(ImageNum).Picture = picSmall(ImageNum).Image
End Function

Private Sub BlurPicDATAREG(ix As Long, iy As Long)
Dim RR As Long, RG As Long, RB As Long, ab As Long
            
   ' 0 1 0
   ' 1 8 1
   ' 0 1 0
   RB = 0: RG = 0: RR = 0: ab = 0
   
   RB = RB + 8 * picDATAORG(0, ix, iy)
   RG = RG + 8 * picDATAORG(1, ix, iy)
   RR = RR + 8 * picDATAORG(2, ix, iy)
   
   RB = RB + picDATAORG(0, ix, iy - 1)
   RG = RG + picDATAORG(1, ix, iy - 1)
   RR = RR + picDATAORG(2, ix, iy - 1)
   
   RB = RB + picDATAORG(0, ix, iy + 1)
   RG = RG + picDATAORG(1, ix, iy + 1)
   RR = RR + picDATAORG(2, ix, iy + 1)
   
   RB = RB + picDATAORG(0, ix - 1, iy)
   RG = RG + picDATAORG(1, ix - 1, iy)
   RR = RR + picDATAORG(2, ix - 1, iy)

   RB = RB + picDATAORG(0, ix + 1, iy)
   RG = RG + picDATAORG(1, ix + 1, iy)
   RR = RR + picDATAORG(2, ix + 1, iy)
   
   RB = RB \ 12 '10
   RG = RG \ 12 '10
   RR = RR \ 12 '10

   picDATAREG(0, ix, iy) = RB
   picDATAREG(1, ix, iy) = RG
   picDATAREG(2, ix, iy) = RR
End Sub

Private Sub BlurDATACUL(ix As Long, iy As Long, ChangeAlpha As Long)
Dim RR As Long, RG As Long, RB As Long, ab As Long

'Dim u1 As Long
'Dim u2 As Long
'Dim u3 As Long
'u1 = UBound(DATACULORG(), 1)
'u2 = UBound(DATACULORG(), 2)
'u3 = UBound(DATACULORG(), 3)
   ' 0 1 0
   ' 1 8 1
   ' 0 1 0
   
   ' ChangeAlpha
   ' 0  Blur ALL
   ' 1  Blur image only
   ' 2  Blur Alpha only
   RB = 0: RG = 0: RR = 0: ab = 0
   
   If DATACULORG(3, ix, iy) <> 0 Then
      RB = RB + 8 * DATACULORG(0, ix, iy)
      RG = RG + 8 * DATACULORG(1, ix, iy)
      RR = RR + 8 * DATACULORG(2, ix, iy)
      ab = ab + 1 * DATACULORG(3, ix, iy)
      
      RB = RB + DATACULORG(0, ix, iy - 1)
      RG = RG + DATACULORG(1, ix, iy - 1)
      RR = RR + DATACULORG(2, ix, iy - 1)
      ab = ab + DATACULORG(3, ix, iy - 1)
      
      RB = RB + DATACULORG(0, ix, iy + 1)
      RG = RG + DATACULORG(1, ix, iy + 1)
      RR = RR + DATACULORG(2, ix, iy + 1)
      ab = ab + DATACULORG(3, ix, iy + 1)
   
      RB = RB + DATACULORG(0, ix - 1, iy)
      RG = RG + DATACULORG(1, ix - 1, iy)
      RR = RR + DATACULORG(2, ix - 1, iy)
      ab = ab + DATACULORG(3, ix - 1, iy)
   
      RB = RB + DATACULORG(0, ix + 1, iy)
      RG = RG + DATACULORG(1, ix + 1, iy)
      RR = RR + DATACULORG(2, ix + 1, iy)
      ab = ab + DATACULORG(3, ix + 1, iy)
      
      RB = RB \ 12 '10
      RG = RG \ 12 '10
      RR = RR \ 12 '10
      ab = ab \ 5 '10
      Replace_Alpha_DATACUL ix, iy, CByte(ab), CByte(RB), CByte(RG), CByte(RR), ChangeAlpha
      DealWithSmallNA ix, iy   ' Fill picDATAREG from DATACUL#() convolved with Alpha
   End If
End Sub


Private Sub EdgeBlur(ix As Long, iy As Long, _
      pmx1 As Long, pmy1 As Long, Optional ChangeAlpha As Long = 0)

Dim RR As Long, RG As Long, RB As Long, ab As Long
Dim Cul As Long
      
' Vert Left edge
'   ix = LOX
'   pmx1 = 0: pmy1 = 1

' Vert Right edge
'   ix = HIX
'   pmx1 = 0: pmy1 = 1
      
' Top Horz Edges
'   iy = LOY
'   pmx1 = 1: pmy1 = 0

' Bottom Horz Edges
' iy = HIY
' pmx1 = 1: pmy1 = 0
      
   If ImBPP(ImageNum) <> 32 Then
         
      Cul = RGB(picDATAORG(2, ix, iy), picDATAORG(1, ix, iy), picDATAORG(0, ix, iy))
      Cul = Cul And &HFFFFFF
      If Cul <> TColorBGR Then
         RB = 0: RG = 0: RR = 0: ab = 0
         
         RB = RB + 6 * picDATAORG(0, ix, iy)
         RG = RG + 6 * picDATAORG(1, ix, iy)
         RR = RR + 6 * picDATAORG(2, ix, iy)
         
         RB = RB + picDATAORG(0, ix + pmx1, iy + pmy1)
         RG = RG + picDATAORG(1, ix + pmx1, iy + pmy1)
         RR = RR + picDATAORG(2, ix + pmx1, iy + pmy1)
         
         RB = RB \ 7
         RG = RG \ 7
         RR = RR \ 7
         ab = ab \ 7
         
         picDATAREG(0, ix, iy) = RB
         picDATAREG(1, ix, iy) = RG
         picDATAREG(2, ix, iy) = RR
      End If
   
   Else   ' ImBPP(ImageNum)=32
               
      RB = 0: RG = 0: RR = 0: ab = 0
      If DATACULORG(3, ix, iy) <> 0 Then
         RB = RB + 6 * DATACULORG(0, ix, iy)
         RG = RG + 6 * DATACULORG(1, ix, iy)
         RR = RR + 6 * DATACULORG(2, ix, iy)
         ab = ab + 6 * DATACULORG(3, ix, iy)
          
         RB = RB + DATACULORG(0, ix + pmx1, iy + pmy1)
         RG = RG + DATACULORG(1, ix + pmx1, iy + pmy1)
         RR = RR + DATACULORG(2, ix + pmx1, iy + pmy1)
         ab = ab + DATACULORG(3, ix + pmx1, iy + pmy1)
         RB = RB \ 7
         RG = RG \ 7
         RR = RR \ 7
         ab = ab \ 7
      
         Replace_Alpha_DATACUL ix, iy, CByte(ab), CByte(RB), CByte(RG), CByte(RR), ChangeAlpha
         DealWithSmallNA ix, iy
      End If
   End If

End Sub

Private Sub CornerBlur(ix As Long, iy As Long, _
      pmx1 As Long, pmy1 As Long, pmx2 As Long, pmy2 As Long, Optional ChangeAlpha As Long = 0)
      
Dim RR As Long, RG As Long, RB As Long, ab As Long
Dim Cul As Long

   
   If ImBPP(ImageNum) <> 32 Then
      
      Cul = RGB(picDATAORG(2, ix, iy), picDATAORG(1, ix, iy), picDATAORG(0, ix, iy))
      Cul = Cul And &HFFFFFF
      If Cul <> TColorBGR Then
         RB = 0: RG = 0: RR = 0: ab = 0
         RB = RB + 6 * picDATAORG(0, ix, iy)
         RG = RG + 6 * picDATAORG(1, ix, iy)
         RR = RR + 6 * picDATAORG(2, ix, iy)
         
         RB = RB + picDATAORG(0, ix + pmx1, iy + pmy1)
         RG = RG + picDATAORG(1, ix + pmx1, iy + pmy1)
         RR = RR + picDATAORG(2, ix + pmx1, iy + pmy1)
         
         RB = RB + picDATAORG(0, ix + pmx2, iy + pmy2)
         RG = RG + picDATAORG(1, ix + pmx2, iy + pmy2)
         RR = RR + picDATAORG(2, ix + pmx2, iy + pmy2)
         
         RB = RB \ 8
         RG = RG \ 8
         RR = RR \ 8
         ab = ab \ 8
      End If
   
   Else
      RB = 0: RG = 0: RR = 0: ab = 0
      If DATACULORG(3, ix, iy) <> 0 Then
         RB = RB + 6 * DATACULORG(0, ix, iy)
         RG = RG + 6 * DATACULORG(1, ix, iy)
         RR = RR + 6 * DATACULORG(2, ix, iy)
         ab = ab + 6 * DATACULORG(3, ix, iy)
         
         RB = RB + DATACULORG(0, ix + pmx1, iy + pmy1)
         RG = RG + DATACULORG(1, ix + pmx1, iy + pmy1)
         RR = RR + DATACULORG(2, ix + pmx1, iy + pmy1)
         ab = ab + DATACULORG(3, ix + pmx1, iy + pmy1)
         
         RB = RB + DATACULORG(0, ix + pmx2, iy + pmy2)
         RG = RG + DATACULORG(1, ix + pmx2, iy + pmy2)
         RR = RR + DATACULORG(2, ix + pmx2, iy + pmy2)
         ab = ab + DATACULORG(3, ix + pmx2, iy + pmy2)
      
         RB = RB \ 8
         RG = RG \ 8
         RR = RR \ 8
         ab = ab \ 8
         
         Replace_Alpha_DATACUL ix, iy, CByte(ab), CByte(RB), CByte(RG), CByte(RR), ChangeAlpha
         DealWithSmallNA ix, iy
      End If
   End If
End Sub

Public Function LeftColorRelief(picS As PictureBox) As Boolean    ' picS = picSmall(ImageNum)

'NO CHANGE in Alpha

Dim Cul As Long
Dim ix As Long, iy As Long
Dim RR As Long, RG As Long, RB As Long
Dim LR As Byte, LG As Byte, LB As Byte
'Dim MeanGrey As Long
Dim S As Long, SS As Byte
   
   
Dim ixm As Long, ixp As Long
Dim iym As Long, iyp As Long

   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()
   
   ReDim picDATAORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   picDATAORG() = picDATAREG()   ' Needed in case SELREC
   
   LngToRGB LColor, LR, LG, LB
'   MeanGrey = (0.3 * LR + 0.6 * LG + 0.1 * LB)
'   If MeanGrey > 255 Then MeanGrey = 128
   
   LeftColorRelief = False
   For iy = LOY To HIY
   For ix = LOX To HIX
      
      Cul = RGB(picDATAORG(2, ix, iy), picDATAORG(1, ix, iy), picDATAORG(0, ix, iy))
      Cul = Cul And &HFFFFFF
      If Cul <> TColorBGR Then
         LeftColorRelief = True
         RB = picDATAORG(0, ix, iy)
         RG = picDATAORG(1, ix, iy)
         RR = picDATAORG(2, ix, iy)
         S = (0.3 * RR + 0.6 * RG + 0.1 * RB) + 128 'CByte(MeanGrey) '32 ''+ 128 ' Good enough !?
         If S > 255 Then S = 255
         SS = CByte(S)
         picDATAREG(0, ix, iy) = SS
         picDATAREG(1, ix, iy) = SS
         picDATAREG(2, ix, iy) = SS
         If ImBPP(ImageNum) = 32 Then
            Select Case ImageNum
            Case 0
               DATACUL0(0, ix, iy) = SS
               DATACUL0(1, ix, iy) = SS
               DATACUL0(2, ix, iy) = SS
            Case 1
               DATACUL1(0, ix, iy) = SS
               DATACUL1(1, ix, iy) = SS
               DATACUL1(2, ix, iy) = SS
            Case 2
               DATACUL2(0, ix, iy) = SS
               DATACUL2(1, ix, iy) = SS
               DATACUL2(2, ix, iy) = SS
            End Select
         End If
      'Else  ' Cul = TColorBGR
      End If
   Next ix
   Next iy
   
   
' Relief after grey roughly a Disabled button
'+2 +1  0
'+1  0 -1
'0  -1 -2
   picDATAORG() = picDATAREG()
   
   For iy = LOY To HIY
   For ix = LOX To HIX
         Cul = RGB(picDATAORG(2, ix, iy), picDATAORG(1, ix, iy), picDATAORG(0, ix, iy))
         Cul = Cul And &HFFFFFF
         If Cul <> TColorBGR Then
            ixm = ix - 1
            ixp = ix + 1
            iym = iy - 1
            iyp = iy + 1
            If ixm < LOX Then ixm = LOX
            If ixp > HIX Then ixp = HIX
            If iym < LOY Then iym = LOY
            If iyp > HIY Then iyp = HIY
            RR = 0: RG = 0: RB = 0
            RB = RB + 2& * picDATAORG(0, ixm, iyp)
            RG = RG + 2& * picDATAORG(1, ixm, iyp)
            RR = RR + 2& * picDATAORG(2, ixm, iyp)
   
            RB = RB + picDATAORG(0, ix, iyp)
            RG = RG + picDATAORG(1, ix, iyp)
            RR = RR + picDATAORG(2, ix, iyp)
   
            RB = RB + picDATAORG(0, ixm, iy)
            RG = RG + picDATAORG(1, ixm, iy)
            RR = RR + picDATAORG(2, ixm, iy)
   
            RB = RB - 2& * picDATAORG(0, ixp, iym)
            RG = RG - 2& * picDATAORG(1, ixp, iym)
            RR = RR - 2& * picDATAORG(2, ixp, iym)
   
            RB = RB - picDATAORG(0, ix, iym)
            RG = RG - picDATAORG(1, ix, iym)
            RR = RR - picDATAORG(2, ix, iym)
   
            RB = RB - picDATAORG(0, ixp, iy)
            RG = RG - picDATAORG(1, ixp, iy)
            RR = RR - picDATAORG(2, ixp, iy)
   
            'Backcolor = Left Color
            RB = (RB + picDATAORG(0, ix, iy) + LB) \ 3
            RG = (RG + picDATAORG(1, ix, iy) + LG) \ 3
            RR = (RR + picDATAORG(2, ix, iy) + LR) \ 3
   
   '         'Backcolor = Left white
   '         RB = (RB + picDATAORG(0, ix, iy) + 255) \ 3
   '         RG = (RG + picDATAORG(1, ix, iy) + 255) \ 3
   '         RR = (RR + picDATAORG(2, ix, iy) + 255) \ 3
   
            If RB < 0 Then RB = 0
            If RB > 255 Then RB = 255
            If RG < 0 Then RG = 0
            If RG > 255 Then RG = 255
            If RR < 0 Then RR = 0
            If RR > 255 Then RR = 255
            
            picDATAREG(0, ix, iy) = RB
            picDATAREG(1, ix, iy) = RG
            picDATAREG(2, ix, iy) = RR
            
            
            If ImBPP(ImageNum) = 32 Then
               Select Case ImageNum
               Case 0
                  DATACUL0(0, ix, iy) = RB
                  DATACUL0(1, ix, iy) = RG
                  DATACUL0(2, ix, iy) = RR
               Case 1
                  DATACUL1(0, ix, iy) = RB
                  DATACUL1(1, ix, iy) = RG
                  DATACUL1(2, ix, iy) = RR
               Case 2
                  DATACUL2(0, ix, iy) = RB
                  DATACUL2(1, ix, iy) = RG
                  DATACUL2(2, ix, iy) = RR
               End Select
               DealWithSmallNA ix, iy
            End If
         'Else  ' Cul = TColorBGR
         End If   ' If Cul <> TColorBGR Then
   Next ix
   Next iy

End Function

Public Function Invert(picS As PictureBox) As Boolean     ' picS = picSmall(ImageNum)

'NO CHANGE in Alpha

Dim Cul As Long
Dim ix As Long, iy As Long
Dim LRed As Byte, LGreen As Byte, LBlue As Byte
   
   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()
   Invert = False
   For iy = LOY To HIY
   For ix = LOX To HIX
      
      Cul = RGB(picDATAREG(2, ix, iy), picDATAREG(1, ix, iy), picDATAREG(0, ix, iy))
      Cul = Cul And &HFFFFFF
      If Cul <> TColorBGR Then
         Invert = True
         picDATAREG(0, ix, iy) = Not picDATAREG(0, ix, iy)
         picDATAREG(1, ix, iy) = Not picDATAREG(1, ix, iy)
         picDATAREG(2, ix, iy) = Not picDATAREG(2, ix, iy)
         If ImBPP(ImageNum) = 32 Then
            LBlue = picDATAREG(0, ix, iy)
            LGreen = picDATAREG(1, ix, iy)
            LRed = picDATAREG(2, ix, iy)
            Select Case ImageNum
            Case 0
               DATACUL0(0, ix, iy) = Not DATACUL0(0, ix, iy)
               DATACUL0(1, ix, iy) = Not DATACUL0(1, ix, iy)
               DATACUL0(2, ix, iy) = Not DATACUL0(2, ix, iy)
            Case 1
               DATACUL1(0, ix, iy) = Not DATACUL1(0, ix, iy)
               DATACUL1(1, ix, iy) = Not DATACUL1(1, ix, iy)
               DATACUL1(2, ix, iy) = Not DATACUL1(2, ix, iy)
            Case 2
               DATACUL2(0, ix, iy) = Not DATACUL2(0, ix, iy)
               DATACUL2(1, ix, iy) = Not DATACUL2(1, ix, iy)
               DATACUL2(2, ix, iy) = Not DATACUL2(2, ix, iy)
            End Select
            DealWithSmallNA ix, iy  ' picDATAREG() from DATACUL#()
         End If
      Else  ' Cul = TColorBGR
      End If
   Next ix
   Next iy

End Function

Public Function BrighterDarker(picS As PictureBox) As Boolean

'NO CHANGE in Alpha


' Effects Brighter/Darker
' picS = picSmall(ImageNum)
' zFrac = 1.1 brighter or 0.9 darker
Dim ix As Long, iy As Long
Dim Cul As Long
Dim k As Long
Dim zFrac() As Single
Dim LRed As Byte, LGreen As Byte, LBlue As Byte
   
   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()
   
   ReDim picDATAORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   picDATAORG() = picDATAREG()   ' Needed in case SELREC

   ReDim zFrac(0 To 2)
   For k = 0 To 2
      zFrac(k) = 1
   Next k

   Select Case Tools
   Case [BrighterIm]
      zFrac(0) = 1.1
      zFrac(1) = 1.1
      zFrac(2) = 1.1
   Case [DarkerIm]
      zFrac(0) = 0.9
      zFrac(1) = 0.9
      zFrac(2) = 0.9
   End Select

   BrighterDarker = False
   For iy = LOY To HIY
   For ix = LOX To HIX
      Cul = RGB(picDATAORG(2, ix, iy), picDATAORG(1, ix, iy), picDATAORG(0, ix, iy))
      Cul = Cul And &HFFFFFF
      If Cul <> TColorBGR Then
         BrighterDarker = True
         If Cul = 0 Then   '(0,0,0)
            For k = 0 To 2
                  If zFrac(k) > 1 Then
                     Cul = (1& * picDATAORG(k, ix, iy) + 51) * zFrac(k)
                     ' NB 51 so that x 1.01 increase the long value
                     If Cul > 255 Then Cul = 255
                     picDATAREG(k, ix, iy) = CByte(Cul)
                  End If
            Next k
            If ImBPP(ImageNum) = 32 Then
               LBlue = picDATAREG(0, ix, iy)
               LGreen = picDATAREG(1, ix, iy)
               LRed = picDATAREG(2, ix, iy)
               Select Case ImageNum
               Case 0
                  DATACUL0(0, ix, iy) = LBlue
                  DATACUL0(1, ix, iy) = LGreen
                  DATACUL0(2, ix, iy) = LRed
               Case 1
                  DATACUL1(0, ix, iy) = LBlue
                  DATACUL1(1, ix, iy) = LGreen
                  DATACUL1(2, ix, iy) = LRed
               Case 2
                  DATACUL2(0, ix, iy) = LBlue
                  DATACUL2(1, ix, iy) = LGreen
                  DATACUL2(2, ix, iy) = LRed
               End Select
               DealWithSmallNA ix, iy    ' picDATAREG
            End If
            
         Else  '(B,0,0),(0,G,0),(0,0,R)
               '(B,G,0),(B,0,R),(0,G,R)
               '(B,G,R)
            For k = 0 To 2
                  If picDATAORG(k, ix, iy) > 0 Then
                     If zFrac(k) > 1 Then
                        Cul = (1& * picDATAORG(k, ix, iy) + 6) '* zFrac(k)
                        If Cul > 255 Then Cul = 255
                     Else
                        Cul = (1& * picDATAORG(k, ix, iy) - 6) '* zFrac(k)
                        If Cul < 0 Then Cul = 0
                     End If
                     picDATAREG(k, ix, iy) = CByte(Cul)
                  End If
            Next k
            If ImBPP(ImageNum) = 32 Then
               LBlue = picDATAREG(0, ix, iy)
               LGreen = picDATAREG(1, ix, iy)
               LRed = picDATAREG(2, ix, iy)
               Select Case ImageNum
               Case 0
                  DATACUL0(0, ix, iy) = LBlue
                  DATACUL0(1, ix, iy) = LGreen
                  DATACUL0(2, ix, iy) = LRed
               Case 1
                  DATACUL1(0, ix, iy) = LBlue
                  DATACUL1(1, ix, iy) = LGreen
                  DATACUL1(2, ix, iy) = LRed
               Case 2
                  DATACUL2(0, ix, iy) = LBlue
                  DATACUL2(1, ix, iy) = LGreen
                  DATACUL2(2, ix, iy) = LRed
                  DealWithSmallNA ix, iy    ' picDATAREG from DATACUL#() & Alpha
               End Select
            End If
         End If
      'Else  ' Cul = TColorBGR
      End If
   Next ix
   Next iy
   
End Function


Public Sub BrightDarkTool(picS As PictureBox, X As Single, Y As Single)
' & MoreLessRGBTools
' picS = picSmall(ImageNum)
Dim xs As Single, ys As Single
Dim Rad As Single, RadSq As Long
Dim i As Single, j As Single
Dim k As Long
Dim Cul As Long  ', Cul32 As Long
Dim ix As Long, iy As Long
Dim zFrac() As Single
Dim ABYTE As Byte

   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()
   ReDim picDATAORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   picDATAORG() = picDATAREG()   ' Needed in case SELREC
   
   ReDim zFrac(0 To 2)
   For k = 0 To 2
      zFrac(k) = 1
   Next k

   Select Case Tools
   Case [Tools2Bright]
      zFrac(0) = 1.01
      zFrac(1) = 1.01
      zFrac(2) = 1.01
   Case [Tools2Dark]
      zFrac(0) = 0.99
      zFrac(1) = 0.99
      zFrac(2) = 0.99
   Case [Tools2MoreBlue]
      zFrac(0) = 1.01
   Case [Tools2MoreGreen]
      zFrac(1) = 1.01
   Case [Tools2MoreRed]
      zFrac(2) = 1.01
   Case [Tools2LessBlue]
      zFrac(0) = 0.99
   Case [Tools2LessGreen]
      zFrac(1) = 0.99
   Case [Tools2LessRed]
      zFrac(2) = 0.99
   End Select
   
   xs = X \ GridMult - 1
   ys = ImageHeight(ImageNum) - Y \ GridMult
   Rad = 0.75  ' Maybe other values
   RadSq = Rad * Rad
   
   For j = -Rad To Rad
      iy = j + ys
      If iy >= 0 Then
      If iy <= ImageHeight(ImageNum) - 1 Then
            For i = -Rad To Rad
               ix = i + xs
               If ix >= 0 Then
               If ix <= ImageWidth(ImageNum) - 1 Then
                  
                  Cul = RGB(picDATAORG(2, ix, iy), picDATAORG(1, ix, iy), picDATAORG(0, ix, iy))
                  Cul = Cul And &HFFFFFF
                  If Cul <> TColorBGR Then
                     If Cul = 0 Then   '(0,0,0)
                        For k = 0 To 2
                              If zFrac(k) > 1 Then
                                 Cul = (1& * picDATAORG(k, ix, iy) + 51) * zFrac(k)
                                 ' NB 51 so that x 1.01 increase the long value
                                 If Cul > 255 Then Cul = 255
                                 picDATAREG(k, ix, iy) = CByte(Cul)
                                 If ImBPP(ImageNum) = 32 Then
                                    Select Case ImageNum
                                    Case 0
                                       DATACUL0(k, ix, iy) = CByte(Cul) 'picDATAREG(k, ix, iy)
                                    Case 1
                                       DATACUL1(k, ix, iy) = CByte(Cul) 'picDATAREG(k, ix, iy)
                                    Case 2
                                       DATACUL2(k, ix, iy) = CByte(Cul) 'picDATAREG(k, ix, iy)
                                    End Select
                                 End If
                              End If
                        Next k
                        ' k=3
                        If ImBPP(ImageNum) = 32 Then
                           Select Case ImageNum
                           Case 0: ABYTE = DATACUL0(3, ix, iy)
                           Case 1: ABYTE = DATACUL1(3, ix, iy)
                           Case 2: ABYTE = DATACUL2(3, ix, iy)
                           End Select
                           picDATAREG(3, ix, iy) = ABYTE
                        End If
                               
                     Else  ' Cul <> 0
                           '(B,0,0),(0,G,0),(0,0,R)
                           '(B,G,0),(B,0,R),(0,G,R)
                           '(B,G,R)
                        For k = 0 To 2
                           If picDATAORG(k, ix, iy) = 0 Then
                              If zFrac(k) > 1 Then
                                 ' NB 51 so that x 1.01 increase the long value
                                 Cul = (1& * picDATAORG(k, ix, iy) + 51) * zFrac(k)
                                 If Cul > 255 Then Cul = 255
                                 picDATAREG(k, ix, iy) = CByte(Cul)
                                 If ImBPP(ImageNum) = 32 Then
                                    Select Case ImageNum
                                    Case 0
                                       DATACUL0(k, ix, iy) = CByte(Cul) 'picDATAREG(k, ix, iy)
                                    Case 1
                                       DATACUL1(k, ix, iy) = CByte(Cul) 'picDATAREG(k, ix, iy)
                                    Case 2
                                       DATACUL2(k, ix, iy) = CByte(Cul) 'picDATAREG(k, ix, iy)
                                    End Select
                                 End If
                              End If
                           Else  'If picDATAORG(k, ix, iy) > 0 Then
                              Cul = 1& * picDATAORG(k, ix, iy)
                              ' NB 50 * 0.99 = 49.5 == 50 again lower limit
                              ' NB 50 * 1.01 = 50.05 == 50 again lower limit
                              ' NB 51 * 1.01 = 51.51 = 51
                              If Cul < 51 Then Cul = 51
                              Cul = Cul * zFrac(k)
                              If Cul > 255 Then Cul = 255
                              picDATAREG(k, ix, iy) = CByte(Cul)
                              If ImBPP(ImageNum) = 32 Then
                                 Select Case ImageNum
                                 Case 0
                                    DATACUL0(k, ix, iy) = CByte(Cul) 'picDATAREG(k, ix, iy)
                                 Case 1
                                    DATACUL1(k, ix, iy) = CByte(Cul) 'picDATAREG(k, ix, iy)
                                 Case 2
                                    DATACUL2(k, ix, iy) = CByte(Cul) 'picDATAREG(k, ix, iy)
                                 End Select
                              End If
                           End If
                        Next k
                        ' k=3
                        If ImBPP(ImageNum) = 32 Then
                           Select Case ImageNum
                           Case 0: ABYTE = DATACUL0(3, ix, iy)
                           Case 1: ABYTE = DATACUL1(3, ix, iy)
                           Case 2: ABYTE = DATACUL2(3, ix, iy)
                           End Select
                           picDATAREG(3, ix, iy) = ABYTE
                        End If
                        
                     End If
                  Else  ' Cul = TColorBGR
                     picDATAREG(3, ix, iy) = 0
                  End If
                  
                  DealWithSmallNA ix, iy  ' On picDATAREG()
               
               End If   ' If ix <= ImageWidth(ImageNum) - 1 Then
               End If   ' If ix >= 0 Then
            Next i   ' For i = -Rad To Rad
      End If   ' If iy <= ImageHeight(ImageNum) - 1 Then
      End If   ' If iy >= 0 Then
   
   Next j   ' For j = -Rad To Rad

   DisplayTools2 picS
   ' Does SetDIBits from picDATAREG()
End Sub

Public Sub ReplaceLbyRColor(PIC As PictureBox)  ' PIC = picSmall(ImageNum)
' PIC = picSmall(ImageNum)
Dim Cul As Long
Dim ix As Long, iy As Long
Dim RR As Byte, RG As Byte, RB As Byte
Dim ABYTE As Byte
   
   GetTheBitsBGR ImageNum, PIC   ' picSmall(ImageNum) -->> PICDATAREG()
   
   ReDim picDATAORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   picDATAORG() = picDATAREG()   ' Needed in case SELREC
   
   LngToRGB RColor, RR, RG, RB
   For iy = LOY To HIY
   For ix = LOX To HIX
      Cul = PIC.Point(ix, ImageHeight(ImageNum) - 1 - iy) ' HIY to 0  or HIY-LOY to 0
      If Cul <> TColorBGR Then
         If Cul >= 0 Then
            If Cul = LColor Then
               picDATAREG(0, ix, iy) = RB
               picDATAREG(1, ix, iy) = RG
               picDATAREG(2, ix, iy) = RR
               picDATAREG(3, ix, iy) = 255
               If ImBPP(ImageNum) = 32 Then
                  Select Case ImageNum
                  Case 0: ABYTE = DATACUL0(3, ix, iy)
                  Case 1: ABYTE = DATACUL1(3, ix, iy)
                  Case 2: ABYTE = DATACUL2(3, ix, iy)
                  End Select
                  CulTo32bppArrays RColor, ix, ImageHeight(ImageNum) - 1 - iy, ABYTE
                  picDATAREG(3, ix, iy) = ABYTE
                  DealWithSmallNA ix, iy  ' On picDATAREG()
               End If
            End If
         End If
      Else
      ix = ix
      End If
   Next ix
   Next iy
End Sub

Public Sub Border(picS As PictureBox)     ' picS = picSmall(ImageNum)
Dim ix As Long, iy As Long
Dim LRed As Byte, LGreen As Byte, LBlue As Byte  ' Left color
Dim RRed As Byte, RGreen As Byte, RBlue As Byte  ' Right color
   
   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()
   
   LngToRGB LColor, LRed, LGreen, LBlue
   LngToRGB RColor, RRed, RGreen, RBlue
   
   For ix = LOX + 1 To HIX - 1
      picDATAREG(0, ix, HIY) = RBlue
      picDATAREG(1, ix, HIY) = RGreen
      picDATAREG(2, ix, HIY) = RRed
      picDATAREG(0, ix, LOY) = LBlue
      picDATAREG(1, ix, LOY) = LGreen
      picDATAREG(2, ix, LOY) = LRed
      
      If ImBPP(ImageNum) = 32 Then
         Select Case ImageNum
         Case 0
            DATACUL0(0, ix, HIY) = RBlue
            DATACUL0(1, ix, HIY) = RGreen
            DATACUL0(2, ix, HIY) = RRed
            DATACUL0(3, ix, HIY) = 255
            
            DATACUL0(0, ix, LOY) = LBlue
            DATACUL0(1, ix, LOY) = LGreen
            DATACUL0(2, ix, LOY) = LRed
            DATACUL0(3, ix, LOY) = 255
         Case 1
            DATACUL1(0, ix, HIY) = RBlue
            DATACUL1(1, ix, HIY) = RGreen
            DATACUL1(2, ix, HIY) = RRed
            DATACUL1(3, ix, HIY) = 255
            
            DATACUL1(0, ix, LOY) = LBlue
            DATACUL1(1, ix, LOY) = LGreen
            DATACUL1(2, ix, LOY) = LRed
            DATACUL1(3, ix, LOY) = 255
         Case 2
            DATACUL2(0, ix, HIY) = RBlue
            DATACUL2(1, ix, HIY) = RGreen
            DATACUL2(2, ix, HIY) = RRed
            DATACUL2(3, ix, HIY) = 255
            
            DATACUL2(0, ix, LOY) = LBlue
            DATACUL2(1, ix, LOY) = LGreen
            DATACUL2(2, ix, LOY) = LRed
            DATACUL2(3, ix, LOY) = 255
         End Select
      End If
      
   Next ix
   For iy = LOY + 1 To HIY - 1
      picDATAREG(0, LOX, iy) = RBlue
      picDATAREG(1, LOX, iy) = RGreen
      picDATAREG(2, LOX, iy) = RRed
      picDATAREG(0, HIX, iy) = LBlue
      picDATAREG(1, HIX, iy) = LGreen
      picDATAREG(2, HIX, iy) = LRed
      If ImBPP(ImageNum) = 32 Then
         Select Case ImageNum
         Case 0
            DATACUL0(0, LOX, iy) = RBlue
            DATACUL0(1, LOX, iy) = RGreen
            DATACUL0(2, LOX, iy) = RRed
            DATACUL0(3, LOX, iy) = 255
            
            DATACUL0(0, HIX, iy) = LBlue
            DATACUL0(1, HIX, iy) = LGreen
            DATACUL0(2, HIX, iy) = LRed
            DATACUL0(3, HIX, iy) = 255
         Case 1
            DATACUL1(0, LOX, iy) = RBlue
            DATACUL1(1, LOX, iy) = RGreen
            DATACUL1(2, LOX, iy) = RRed
            DATACUL1(3, LOX, iy) = 255
            
            DATACUL1(0, HIX, iy) = LBlue
            DATACUL1(1, HIX, iy) = LGreen
            DATACUL1(2, HIX, iy) = LRed
            DATACUL1(3, HIX, iy) = 255
         Case 2
            DATACUL2(0, LOX, iy) = RBlue
            DATACUL2(1, LOX, iy) = RGreen
            DATACUL2(2, LOX, iy) = RRed
            DATACUL2(3, LOX, iy) = 255
            
            DATACUL2(0, HIX, iy) = LBlue
            DATACUL2(1, HIX, iy) = LGreen
            DATACUL2(2, HIX, iy) = LRed
            DATACUL2(3, HIX, iy) = 255
         End Select
      End If
   
   Next iy
   
   ' Transparent corners
   picDATAREG(0, LOX, LOY) = 197
   picDATAREG(1, LOX, LOY) = 195
   picDATAREG(2, LOX, LOY) = 194
   picDATAREG(3, LOX, LOY) = 0
   
   picDATAREG(0, LOX, HIY) = 197
   picDATAREG(1, LOX, HIY) = 195
   picDATAREG(2, LOX, HIY) = 194
   picDATAREG(3, LOX, HIY) = 0
   
   picDATAREG(0, HIX, LOY) = 197
   picDATAREG(1, HIX, LOY) = 195
   picDATAREG(2, HIX, LOY) = 194
   picDATAREG(3, HIX, LOY) = 0
   
   picDATAREG(0, HIX, HIY) = 197
   picDATAREG(1, HIX, HIY) = 195
   picDATAREG(2, HIX, HIY) = 194
   picDATAREG(3, HIX, HIY) = 0
   
   
   If ImBPP(ImageNum) = 32 Then
      Select Case ImageNum
      Case 0
         ' Transparent corners
         DATACUL0(0, LOX, LOY) = 197
         DATACUL0(1, LOX, LOY) = 195
         DATACUL0(2, LOX, LOY) = 194
         DATACUL0(3, LOX, LOY) = 0
         
         DATACUL0(0, LOX, HIY) = 197
         DATACUL0(1, LOX, HIY) = 195
         DATACUL0(2, LOX, HIY) = 194
         DATACUL0(3, LOX, HIY) = 0
         
         DATACUL0(0, HIX, LOY) = 197
         DATACUL0(1, HIX, LOY) = 195
         DATACUL0(2, HIX, LOY) = 194
         DATACUL0(3, HIX, LOY) = 0
         
         DATACUL0(0, HIX, HIY) = 197
         DATACUL0(1, HIX, HIY) = 195
         DATACUL0(2, HIX, HIY) = 194
         DATACUL0(3, HIX, HIY) = 0
         
         For ix = LOX + 1 To HIX - 1
            DATACUL0(3, ix, HIY) = 255
            DATACUL0(3, ix, LOY) = 255
         Next ix
         For iy = LOY + 1 To HIY - 1
            DATACUL0(3, LOX, iy) = 255
            DATACUL0(3, HIX, iy) = 255
         Next iy
      Case 1
         ' Transparent corners
         DATACUL1(0, LOX, LOY) = 197
         DATACUL1(1, LOX, LOY) = 195
         DATACUL1(2, LOX, LOY) = 194
         DATACUL1(3, LOX, LOY) = 0
         
         DATACUL1(0, LOX, HIY) = 197
         DATACUL1(1, LOX, HIY) = 195
         DATACUL1(2, LOX, HIY) = 194
         DATACUL1(3, LOX, HIY) = 0
         
         DATACUL1(0, HIX, LOY) = 197
         DATACUL1(1, HIX, LOY) = 195
         DATACUL1(2, HIX, LOY) = 194
         DATACUL1(3, HIX, LOY) = 0
         
         DATACUL1(0, HIX, HIY) = 197
         DATACUL1(1, HIX, HIY) = 195
         DATACUL1(2, HIX, HIY) = 194
         DATACUL1(3, HIX, HIY) = 0
      
         For ix = LOX + 1 To HIX - 1
            DATACUL1(3, ix, HIY) = 255
            DATACUL1(3, ix, LOY) = 255
         Next ix
         For iy = LOY + 1 To HIY - 1
            DATACUL1(3, LOX, iy) = 255
            DATACUL1(3, HIX, iy) = 255
         Next iy
      
      
      Case 2
         ' Transparent corners
         DATACUL2(0, LOX, LOY) = 197
         DATACUL2(1, LOX, LOY) = 195
         DATACUL2(2, LOX, LOY) = 194
         DATACUL2(3, LOX, LOY) = 0
         
         DATACUL2(0, LOX, HIY) = 197
         DATACUL2(1, LOX, HIY) = 195
         DATACUL2(2, LOX, HIY) = 194
         DATACUL2(3, LOX, HIY) = 0
         
         DATACUL2(0, HIX, LOY) = 197
         DATACUL2(1, HIX, LOY) = 195
         DATACUL2(2, HIX, LOY) = 194
         DATACUL2(3, HIX, LOY) = 0
         
         DATACUL2(0, HIX, HIY) = 197
         DATACUL2(1, HIX, HIY) = 195
         DATACUL2(2, HIX, HIY) = 194
         DATACUL2(3, HIX, HIY) = 0
      
         For ix = LOX + 1 To HIX - 1
            DATACUL2(3, ix, HIY) = 255
            DATACUL2(3, ix, LOY) = 255
         Next ix
         For iy = LOY + 1 To HIY - 1
            DATACUL2(3, LOX, iy) = 255
            DATACUL2(3, HIX, iy) = 255
         Next iy
      
      
      End Select
   End If
   
End Sub

'####  Tools2 Tools #####

Public Sub BlurTool(picS As PictureBox, X As Single, Y As Single)
' picS = picSmall(ImageNum)
' shpCirc in picPANEL
Dim xs As Single, ys As Single
Dim Rad As Single, RadSq As Long
Dim i As Single, j As Single
Dim Cul As Long
Dim LR As Long, LG As Long, LB As Long
Dim ix As Long, iy As Long
Dim ABYTE As Byte

   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()
   
   ReDim picDATAORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   picDATAORG() = picDATAREG()   ' Needed in case SELREC

   xs = X \ GridMult - 1
   ys = ImageHeight(ImageNum) - Y \ GridMult
   Rad = 0.75  ' Maybe other values
   RadSq = Rad * Rad
   
   For j = -Rad To Rad
      iy = j + ys
      If iy > 0 Then
      If iy < ImageHeight(ImageNum) - 1 Then
            For i = -Rad To Rad
               ix = i + xs
               If ix > 0 Then
               If ix < ImageWidth(ImageNum) - 1 Then
                  Cul = RGB(picDATAORG(2, ix, iy), picDATAORG(1, ix, iy), picDATAORG(0, ix, iy))
                  Cul = Cul And &HFFFFFF
                  If Cul <> TColorBGR Then
                     LB = 8 * picDATAORG(0, ix, iy)
                     LG = 8 * picDATAORG(1, ix, iy)
                     LR = 8 * picDATAORG(2, ix, iy)
                     
                     LB = LB + picDATAORG(0, ix - 1, iy - 1)
                     LG = LG + picDATAORG(1, ix - 1, iy - 1)
                     LR = LR + picDATAORG(2, ix - 1, iy - 1)
                     
                     LB = LB + picDATAORG(0, ix + 1, iy + 1)
                     LG = LG + picDATAORG(1, ix + 1, iy + 1)
                     LR = LR + picDATAORG(2, ix + 1, iy + 1)
                     
                     LB = LB + picDATAORG(0, ix - 1, iy + 1)
                     LG = LG + picDATAORG(1, ix - 1, iy + 1)
                     LR = LR + picDATAORG(2, ix - 1, iy + 1)

                     LB = LB + picDATAORG(0, ix + 1, iy - 1)
                     LG = LG + picDATAORG(1, ix + 1, iy - 1)
                     LR = LR + picDATAORG(2, ix + 1, iy - 1)
                     
                     LB = LB \ 12 '6 '5
                     LG = LG \ 12 '6 '5
                     LR = LR \ 12 '6 '5
                     
                     If LB < 0 Then LB = 0
                     If LB > 255 Then LB = 255
                     If LG < 0 Then LG = 0
                     If LG > 255 Then LG = 255
                     If LR < 0 Then LR = 0
                     If LR > 255 Then LR = 255
                     picDATAREG(0, ix, iy) = LB
                     picDATAREG(1, ix, iy) = LG
                     picDATAREG(2, ix, iy) = LR
                     If ImBPP(ImageNum) = 32 Then
                        Select Case ImageNum
                        Case 0: ABYTE = DATACUL0(3, ix, iy)
                        Case 1: ABYTE = DATACUL1(3, ix, iy)
                        Case 2: ABYTE = DATACUL2(3, ix, iy)
                        End Select
                        Replace_Alpha_DATACUL ix, iy, ABYTE, CByte(LB), CByte(LG), CByte(LR), 0
                        DealWithSmallNA ix, iy
                     End If
                  
                  End If
               End If
               End If
            Next i
      End If
      End If
   Next j

   DisplayTools2 picS
  ' Does SetDIBits from picDATAREG()
End Sub

Public Sub GreyTool(picS As PictureBox, X As Single, Y As Single)
' picS = picSmall(ImageNum)
Dim xs As Single, ys As Single
Dim Rad As Single, RadSq As Long
Dim i As Single, j As Single
Dim Cul As Long
Dim LR As Long, LG As Long, LB As Long
Dim ix As Long, iy As Long
Dim S As Long, SS As Byte
Dim ABYTE As Byte

   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()
   
   ReDim picDATAORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   picDATAORG() = picDATAREG()   ' Needed in case SELREC

   xs = X \ GridMult - 1
   ys = ImageHeight(ImageNum) - Y \ GridMult
   Rad = 0.75  ' Maybe other values
   RadSq = Rad * Rad
   
   For j = -Rad To Rad
      iy = j + ys
      If iy >= 0 Then
      If iy <= ImageHeight(ImageNum) - 1 Then
            For i = -Rad To Rad
                  ix = i + xs
                  If ix >= 0 Then
                  If ix <= ImageWidth(ImageNum) - 1 Then
                        Cul = RGB(picDATAORG(2, ix, iy), picDATAORG(1, ix, iy), picDATAORG(0, ix, iy))
                        Cul = Cul And &HFFFFFF
                        If Cul <> TColorBGR Then
                           LB = picDATAORG(0, ix, iy)
                           LG = picDATAORG(1, ix, iy)
                           LR = picDATAORG(2, ix, iy)
                           S = (0.3 * LR + 0.6 * LG + 0.1 * LB) '+ 128 ' Good enough !?
                           If S > 255 Then S = 255
                           SS = CByte(S)
                           picDATAREG(0, ix, iy) = SS
                           picDATAREG(1, ix, iy) = SS
                           picDATAREG(2, ix, iy) = SS
                           
                           If ImBPP(ImageNum) = 32 Then
                              Select Case ImageNum
                              Case 0: ABYTE = DATACUL0(3, ix, iy)
                              Case 1: ABYTE = DATACUL1(3, ix, iy)
                              Case 2: ABYTE = DATACUL2(3, ix, iy)
                              End Select
                              picDATAREG(3, ix, iy) = ABYTE
                              Replace_Alpha_DATACUL ix, iy, ABYTE, SS, SS, SS, 1
                              DealWithSmallNA ix, iy

                           End If
                        End If
                  End If
                  End If
            Next i
      End If
      End If
   Next j

   DisplayTools2 picS
End Sub

Public Sub InvertTool(picS As PictureBox, X As Single, Y As Single)
' picS = picSmall(ImageNum)
Dim xs As Single, ys As Single
Dim Rad As Single, RadSq As Long
Dim i As Single, j As Single
Dim Cul As Long
Dim LR As Byte, LG As Byte, LB As Byte
Dim ix As Long, iy As Long
Dim ABYTE As Byte

   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()
   
'   ReDim picDATAORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
'   picDATAORG() = picDATAREG()   ' Needed in case SELREC

   xs = X \ GridMult - 1
   ys = ImageHeight(ImageNum) - Y \ GridMult
   Rad = 0.75  ' Maybe other values
   RadSq = Rad * Rad
   
   For j = -Rad To Rad
      iy = j + ys
      If iy >= 0 Then
      If iy <= ImageHeight(ImageNum) - 1 Then
            For i = -Rad To Rad
                  ix = i + xs
                  If ix >= 0 Then
                  If ix <= ImageWidth(ImageNum) - 1 Then
                     Cul = RGB(picDATAREG(2, ix, iy), picDATAREG(1, ix, iy), picDATAREG(0, ix, iy))
                     Cul = Cul And &HFFFFFF
                     If Cul <> TColorBGR Then
                        picDATAREG(0, ix, iy) = Not picDATAREG(0, ix, iy)
                        picDATAREG(1, ix, iy) = Not picDATAREG(1, ix, iy)
                        picDATAREG(2, ix, iy) = Not picDATAREG(2, ix, iy)
                     
                        If ImBPP(ImageNum) = 32 Then
                           LB = picDATAREG(0, ix, iy)
                           LG = picDATAREG(1, ix, iy)
                           LR = picDATAREG(2, ix, iy)
                           Select Case ImageNum
                           Case 0: ABYTE = DATACUL0(3, ix, iy)
                           Case 1: ABYTE = DATACUL1(3, ix, iy)
                           Case 2: ABYTE = DATACUL2(3, ix, iy)
                           End Select
                           picDATAREG(3, ix, iy) = ABYTE
                           Replace_Alpha_DATACUL ix, iy, ABYTE, LB, LG, LR, 0 '1
                           DealWithSmallNA ix, iy
                        End If
                     End If
                  End If
                  End If
            Next i
      End If
      End If
   Next j

   DisplayTools2 picS
End Sub

Public Sub RandomTool(picS As PictureBox, X As Single, Y As Single)
' picS = picSmall(ImageNum)
Dim xs As Single, ys As Single
Dim Rad As Single, RadSq As Long
Dim i As Single, j As Single
Dim Cul As Long
'Dim LR As Long, LG As Long, LB As Long
Dim ix As Long, iy As Long
Dim S0 As Byte, S1 As Byte, s2 As Byte
Dim ABYTE As Byte


   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()
   
'   ReDim picDATAORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
'   picDATAORG() = picDATAREG()   ' Needed in case SELREC

   xs = X \ GridMult - 1
   ys = ImageHeight(ImageNum) - Y \ GridMult
   Rad = 0.75  ' Maybe other values
   RadSq = Rad * Rad
   
   For j = -Rad To Rad
      iy = j + ys
      If iy >= 0 Then
      If iy <= ImageHeight(ImageNum) - 1 Then
            For i = -Rad To Rad
                  ix = i + xs
                  If ix >= 0 Then
                  If ix <= ImageWidth(ImageNum) - 1 Then
                     Cul = RGB(picDATAREG(2, ix, iy), picDATAREG(1, ix, iy), picDATAREG(0, ix, iy))
                     Cul = Cul And &HFFFFFF
                     If Cul <> TColorBGR Then
                        S0 = 255 * Rnd
                        picDATAREG(0, ix, iy) = S0
                        S1 = 255 * Rnd
                        picDATAREG(1, ix, iy) = S1
                        s2 = 255 * Rnd
                        picDATAREG(2, ix, iy) = s2
                        If ImBPP(ImageNum) = 32 Then
                           Select Case ImageNum
                           Case 0: ABYTE = DATACUL0(3, ix, iy)
                           Case 1: ABYTE = DATACUL1(3, ix, iy)
                           Case 2: ABYTE = DATACUL2(3, ix, iy)
                           End Select
                           picDATAREG(3, ix, iy) = ABYTE
                           Replace_Alpha_DATACUL ix, iy, ABYTE, S0, S1, s2, 0 '1
                           DealWithSmallNA ix, iy
                        End If
                     End If
                  End If
                  End If
            Next i
      End If
      End If
   Next j

   DisplayTools2 picS
End Sub

Public Sub CheckerBoard(picS As PictureBox, X As Single, Y As Single, DCul As Long)
' picS = picSmall(ImageNum)
Dim xs As Single, ys As Single
Dim Rad As Single, RadSq As Long
Dim i As Single, j As Single
Dim BR As Byte, BG As Byte, BB As Byte
Dim ix As Long, iy As Long

   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()
   
   xs = X \ GridMult - 1
   ys = ImageHeight(ImageNum) - Y \ GridMult
   Rad = 0.5  ' Maybe other values
   RadSq = Rad * Rad
   
   'LngToRGB LColor, BR, BG, BB
   LngToRGB DCul, BR, BG, BB
   
   For j = -Rad To Rad
      iy = j + ys
      If iy >= 0 Then
      If iy <= ImageHeight(ImageNum) - 1 Then
            For i = -Rad To Rad
                  ix = i + xs
                  If ix >= 0 Then
                  If ix <= ImageWidth(ImageNum) - 1 Then
                     picDATAREG(0, ix, iy) = BB
                     picDATAREG(1, ix, iy) = BG
                     picDATAREG(2, ix, iy) = BR
                     If ImBPP(ImageNum) = 32 Then
                        Replace_Alpha_DATACUL ix, iy, 255, BB, BG, BR
                     End If
                  End If
                  End If
            Next i
      End If
      End If
   Next j

   DisplayTools2 picS
End Sub


Public Sub HorzLines(picS As PictureBox, X As Single, Y As Single, DCul As Long)
' picS = picSmall(ImageNum)
Dim xs As Single, ys As Single
Dim Rad As Single, RadSq As Long
Dim i As Single, j As Single
Dim BR As Byte, BG As Byte, BB As Byte
Dim ix As Long, iy As Long

   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()
   
   xs = X \ GridMult - 1
   ys = ImageHeight(ImageNum) - Y \ GridMult
   Rad = 1  ' Maybe other values
   RadSq = Rad * Rad
   
   LngToRGB DCul, BR, BG, BB
   
   For j = -Rad To Rad
      iy = j + ys
      If iy >= 0 Then
      If iy <= ImageHeight(ImageNum) - 1 Then
            If (iy And 2) = 0 Then
               For i = -Rad To Rad
                     ix = i + xs
                     If ix >= 0 Then
                     If ix <= ImageWidth(ImageNum) - 1 Then
                        picDATAREG(0, ix, iy) = BB
                        picDATAREG(1, ix, iy) = BG
                        picDATAREG(2, ix, iy) = BR
                        If ImBPP(ImageNum) = 32 Then
                           Replace_Alpha_DATACUL ix, iy, 255, BB, BG, BR
                        End If
                     End If
                     End If
               Next i
            End If
      End If
      End If
   Next j

   DisplayTools2 picS
End Sub

Public Sub VertLines(picS As PictureBox, X As Single, Y As Single, DCul As Long)
' picS = picSmall(ImageNum)
Dim xs As Single, ys As Single
Dim Rad As Single, RadSq As Long
Dim i As Single, j As Single
Dim BR As Byte, BG As Byte, BB As Byte
Dim ix As Long, iy As Long

   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()
   
   xs = X \ GridMult - 1
   ys = ImageHeight(ImageNum) - Y \ GridMult
   Rad = 1  ' Maybe other values
   RadSq = Rad * Rad
   
   LngToRGB DCul, BR, BG, BB
   
   For j = -Rad To Rad
      iy = j + ys
      If iy >= 0 Then
      If iy <= ImageHeight(ImageNum) - 1 Then
            For i = -Rad To Rad
                  ix = i + xs
                  If ix >= 0 Then
                  If ix <= ImageWidth(ImageNum) - 1 Then
                     If (ix And 2) = 0 Then
                        picDATAREG(0, ix, iy) = BB
                        picDATAREG(1, ix, iy) = BG
                        picDATAREG(2, ix, iy) = BR
                        If ImBPP(ImageNum) = 32 Then
                           Replace_Alpha_DATACUL ix, iy, 255, BB, BG, BR
                        End If
                     End If
                  End If
                  End If
            Next i
      End If
      End If
   Next j

   DisplayTools2 picS
End Sub


'#### MIRRORS ####

Public Sub Mirrors(picS As PictureBox)   ' picS = picSmall(ImageNum)
Dim ix As Long, iy As Long

   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()

   ReDim picDATAORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   picDATAORG() = picDATAREG()   ' Needed in case SELREC
      
   Select Case Tools
   Case [MirrorL] ' Left
      For iy = LOY To HIY
      For ix = LOX To (LOX + HIX) \ 2
         LeftRightMirrorData ix, iy, LOX, HIX
      Next ix
      Next iy
   Case [MirrorR] ' Right
      For iy = LOY To HIY
      For ix = HIX To (LOX + HIX) \ 2 Step -1
         LeftRightMirrorData ix, iy, LOX, HIX
      Next ix
      Next iy
   Case [MirrorT] ' Top
      For ix = LOX To HIX
      For iy = HIY To (HIY + LOY) \ 2 Step -1
         TopBottomMirrorData ix, iy, LOY, HIY
      Next iy
      Next ix
   Case [MirrorB] ' Bottom
      For ix = LOX To HIX
      For iy = LOY To (HIY + LOY) \ 2
         TopBottomMirrorData ix, iy, LOY, HIY
      Next iy
      Next ix
   End Select
   
   DisplayTools2 picS
End Sub

Private Sub LeftRightMirrorData(ix As Long, iy As Long, LOX As Long, HIX As Long)
   picDATAREG(0, HIX + LOX - ix, iy) = picDATAORG(0, ix, iy)
   picDATAREG(1, HIX + LOX - ix, iy) = picDATAORG(1, ix, iy)
   picDATAREG(2, HIX + LOX - ix, iy) = picDATAORG(2, ix, iy)
   If ImBPP(ImageNum) = 32 Then
      Select Case ImageNum
      Case 0
         DATACUL0(0, HIX + LOX - ix, iy) = DATACUL0(0, ix, iy)
         DATACUL0(1, HIX + LOX - ix, iy) = DATACUL0(1, ix, iy)
         DATACUL0(2, HIX + LOX - ix, iy) = DATACUL0(2, ix, iy)
         DATACUL0(3, HIX + LOX - ix, iy) = DATACUL0(3, ix, iy)
      Case 1
         DATACUL1(0, HIX + LOX - ix, iy) = DATACUL1(0, ix, iy)
         DATACUL1(1, HIX + LOX - ix, iy) = DATACUL1(1, ix, iy)
         DATACUL1(2, HIX + LOX - ix, iy) = DATACUL1(2, ix, iy)
         DATACUL1(3, HIX + LOX - ix, iy) = DATACUL1(3, ix, iy)
      Case 2
         DATACUL2(0, HIX + LOX - ix, iy) = DATACUL2(0, ix, iy)
         DATACUL2(1, HIX + LOX - ix, iy) = DATACUL2(1, ix, iy)
         DATACUL2(2, HIX + LOX - ix, iy) = DATACUL2(2, ix, iy)
         DATACUL2(3, HIX + LOX - ix, iy) = DATACUL2(3, ix, iy)
      End Select
   End If
End Sub

Private Sub TopBottomMirrorData(ix As Long, iy As Long, LOY As Long, HIY As Long)
   picDATAREG(0, ix, HIY + LOY - iy) = picDATAORG(0, ix, iy)
   picDATAREG(1, ix, HIY + LOY - iy) = picDATAORG(1, ix, iy)
   picDATAREG(2, ix, HIY + LOY - iy) = picDATAORG(2, ix, iy)
   If ImBPP(ImageNum) = 32 Then
      Select Case ImageNum
      Case 0
         DATACUL0(0, ix, HIY + LOY - iy) = DATACUL0(0, ix, iy)
         DATACUL0(1, ix, HIY + LOY - iy) = DATACUL0(1, ix, iy)
         DATACUL0(2, ix, HIY + LOY - iy) = DATACUL0(2, ix, iy)
         DATACUL0(3, ix, HIY + LOY - iy) = DATACUL0(3, ix, iy)
      Case 1
         DATACUL1(0, ix, HIY + LOY - iy) = DATACUL1(0, ix, iy)
         DATACUL1(1, ix, HIY + LOY - iy) = DATACUL1(1, ix, iy)
         DATACUL1(2, ix, HIY + LOY - iy) = DATACUL1(2, ix, iy)
         DATACUL1(3, ix, HIY + LOY - iy) = DATACUL1(3, ix, iy)
      Case 2
         DATACUL2(0, ix, HIY + LOY - iy) = DATACUL2(0, ix, iy)
         DATACUL2(1, ix, HIY + LOY - iy) = DATACUL2(1, ix, iy)
         DATACUL2(2, ix, HIY + LOY - iy) = DATACUL2(2, ix, iy)
         DATACUL2(3, ix, HIY + LOY - iy) = DATACUL2(3, ix, iy)
      End Select
   End If
End Sub


'#### REFLECTORS ####

Public Sub Reflectors(picS As PictureBox)   ' picS = picSmall(ImageNum)
' Only with a select rectangle
Dim ix As Long, iy As Long
Dim ixlr As Long, iyab As Long

   GetTheBitsBGR ImageNum, picS   ' picSmall(ImageNum) -->> PICDATAREG()

   ReDim picDATAORG(0 To 3, 0 To ImageWidth(ImageNum) - 1, 0 To ImageHeight(ImageNum) - 1)
   picDATAORG() = picDATAREG()   ' Needed in case SELREC
      
'   [ReflectLeft]
'   [ReflectRight]
'   [ReflectTop]
'   [ReflectBottom]
   Select Case Tools
   Case [ReflectLeft]   ' To Right
      For iy = LOY To HIY
      For ix = HIX To LOX Step -1
         ixlr = 2 * HIX - ix + 1
         If ixlr <= ImageWidth(ImageNum) - 1 Then
            LeftRightReflectData ix, iy, ixlr
         End If
      Next ix
      Next iy
   Case [ReflectRight]   ' To Left
      For iy = LOY To HIY
      For ix = LOX To HIX
         'ixLR = LOX - (ix - LOX) - 1
         ixlr = 2 * LOX - ix - 1
         If ixlr >= 0 Then
            LeftRightReflectData ix, iy, ixlr
         End If
      Next ix
      Next iy
   Case [ReflectTop]  ' To Below
      For ix = LOX To HIX
      For iy = LOY To HIY
         'iyAB = LOY - (iy - LOY) - 1
         iyab = 2 * LOY - iy - 1
         If iyab >= 0 Then
            AboveBelowReflectData ix, iy, iyab
         End If
      Next iy
      Next ix
   Case [ReflectBottom] ' To Above
      For ix = LOX To HIX
      For iy = HIY To LOY Step -1
         'iyAB = HIY - (iy - HIY) + 1
         iyab = 2 * HIY - iy + 1
         If iyab <= ImageHeight(ImageNum) - 1 Then
            AboveBelowReflectData ix, iy, iyab
         End If
      Next iy
      Next ix
   End Select
   DisplayTools2 picS
End Sub

Private Sub LeftRightReflectData(ix As Long, iy As Long, ixlr As Long)
   picDATAREG(0, ixlr, iy) = picDATAORG(0, ix, iy)
   picDATAREG(1, ixlr, iy) = picDATAORG(1, ix, iy)
   picDATAREG(2, ixlr, iy) = picDATAORG(2, ix, iy)
   If ImBPP(ImageNum) = 32 Then
      Select Case ImageNum
      Case 0
         DATACUL0(0, ixlr, iy) = DATACUL0(0, ix, iy)
         DATACUL0(1, ixlr, iy) = DATACUL0(1, ix, iy)
         DATACUL0(2, ixlr, iy) = DATACUL0(2, ix, iy)
         DATACUL0(3, ixlr, iy) = DATACUL0(3, ix, iy)
      Case 1
         DATACUL1(0, ixlr, iy) = DATACUL1(0, ix, iy)
         DATACUL1(1, ixlr, iy) = DATACUL1(1, ix, iy)
         DATACUL1(2, ixlr, iy) = DATACUL1(2, ix, iy)
         DATACUL1(3, ixlr, iy) = DATACUL1(3, ix, iy)
      Case 2
         DATACUL2(0, ixlr, iy) = DATACUL2(0, ix, iy)
         DATACUL2(1, ixlr, iy) = DATACUL2(1, ix, iy)
         DATACUL2(2, ixlr, iy) = DATACUL2(2, ix, iy)
         DATACUL2(3, ixlr, iy) = DATACUL2(3, ix, iy)
      End Select
   End If
End Sub

Private Sub AboveBelowReflectData(ix As Long, iy As Long, iyab As Long)
   picDATAREG(0, ix, iyab) = picDATAORG(0, ix, iy)
   picDATAREG(1, ix, iyab) = picDATAORG(1, ix, iy)
   picDATAREG(2, ix, iyab) = picDATAORG(2, ix, iy)
   If ImBPP(ImageNum) = 32 Then   ' Modfify picDATAREG() with Alpha ??
      Select Case ImageNum
      Case 0
         DATACUL0(0, ix, iyab) = DATACUL0(0, ix, iy)
         DATACUL0(1, ix, iyab) = DATACUL0(1, ix, iy)
         DATACUL0(2, ix, iyab) = DATACUL0(2, ix, iy)
         DATACUL0(3, ix, iyab) = DATACUL0(3, ix, iy)
      Case 1
         DATACUL1(0, ix, iyab) = DATACUL1(0, ix, iy)
         DATACUL1(1, ix, iyab) = DATACUL1(1, ix, iy)
         DATACUL1(2, ix, iyab) = DATACUL1(2, ix, iy)
         DATACUL1(2, ix, iyab) = DATACUL1(3, ix, iy)
      Case 2
         DATACUL2(0, ix, iyab) = DATACUL2(0, ix, iy)
         DATACUL2(1, ix, iyab) = DATACUL2(1, ix, iy)
         DATACUL2(2, ix, iyab) = DATACUL2(2, ix, iy)
         DATACUL2(3, ix, iyab) = DATACUL2(3, ix, iy)
      End Select
   End If
End Sub

Private Sub DisplayTools2(picS As PictureBox)
' picS = picSmall(ImageNum)
Dim BMIH As BITMAPINFOHEADER
   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = ImageWidth(ImageNum)
      .biHeight = ImageHeight(ImageNum)
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With
   SetDIBits picS.hdc, picS.Image, _
      0, ImageHeight(ImageNum), picDATAREG(0, 0, 0), BMIH, 0
   picS.Picture = picS.Image
   Form1.DrawGrid    ' also picSmall() to picPANEL
End Sub

Public Sub DealWithSmallNA(ix As Long, iy As Long)
' Adjust picDATAREG & DATACUL#() for spurious TColorBGR caused by
' small NA just > 0. Only called for BPP = 32.
Dim TR As Byte, TG As Byte, TB As Byte
Dim NR As Byte, NG As Byte, NB As Byte, NA As Byte
Dim NR0 As Byte, NG0 As Byte, NB0 As Byte
Dim Ealpha As Single
   LngToRGB TColorBGR, TR, TG, TB
   
   
'Dim U As Long
'U = UBound(DATACUL0, 2)
      
If ImBPP(ImageNum) = 32 Then
      Select Case ImageNum
      Case 0
         NA = DATACUL0(3, ix, iy)
      Case 1
         NA = DATACUL1(3, ix, iy)
      Case 2
         NA = DATACUL2(3, ix, iy)
      End Select
      
      If NA = 0 Then
         picDATAREG(0, ix, iy) = TB
         picDATAREG(1, ix, iy) = TG
         picDATAREG(2, ix, iy) = TR
         
      Else     ' Alpha>0 = NA
         
      
         Select Case ImageNum
         Case 0
            NB0 = DATACUL0(0, ix, iy)
            NG0 = DATACUL0(1, ix, iy)
            NR0 = DATACUL0(2, ix, iy)
         Case 1
            NB0 = DATACUL1(0, ix, iy)
            NG0 = DATACUL1(1, ix, iy)
            NR0 = DATACUL1(2, ix, iy)
         Case 2
            NB0 = DATACUL2(0, ix, iy)
            NG0 = DATACUL2(1, ix, iy)
            NR0 = DATACUL2(2, ix, iy)
         End Select
         
         Ealpha = (NA / 255)
         NB = TB * (1 - Ealpha) + NB0 * Ealpha
         NG = TG * (1 - Ealpha) + NG0 * Ealpha
         NR = TR * (1 - Ealpha) + NR0 * Ealpha
      
         ' Adjust for tiny Ealpha where NB=TB,, when NA<>0
         ' to avoid spurious transparency
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
         picDATAREG(0, ix, iy) = NB
         picDATAREG(1, ix, iy) = NG
         picDATAREG(2, ix, iy) = NR
      
      End If
End If
End Sub

'Public Sub ReconcileForAlpha(ImNum As Long)
'' Only If ImBPP(ImeNum) = 32 After Blur
'Dim ix As Long, iy As Long
'Dim TR As Byte, TG As Byte, TB As Byte
'Dim NR As Byte, NG As Byte, NB As Byte, NA As Byte
'Dim NR0 As Byte, NG0 As Byte, NB0 As Byte
'Dim Ealpha As Single
'
'   LngToRGB TColorBGR, TR, TG, TB
'
'   For iy = LOY To HIY
'   For ix = LOX To HIX
'      If picDATAREG(3, ix, ImageHeight(ImNum) - 1 - iy) = 0 Then
'
'         picDATAREG(0, ix, ImageHeight(ImNum) - 1 - iy) = TB
'         picDATAREG(1, ix, ImageHeight(ImNum) - 1 - iy) = TG
'         picDATAREG(2, ix, ImageHeight(ImNum) - 1 - iy) = TR
'
'
'         If ImBPP(ImNum) = 32 Then
'            Select Case ImNum
'            Case 0
'               DATACUL0(0, ix, ImageHeight(ImNum) - 1 - iy) = TB
'               DATACUL0(1, ix, ImageHeight(ImNum) - 1 - iy) = TG
'               DATACUL0(2, ix, ImageHeight(ImNum) - 1 - iy) = TR
'            Case 1
'               DATACUL1(0, ix, ImageHeight(ImNum) - 1 - iy) = TB
'               DATACUL1(1, ix, ImageHeight(ImNum) - 1 - iy) = TG
'               DATACUL1(2, ix, ImageHeight(ImNum) - 1 - iy) = TR
'            Case 2
'               DATACUL2(0, ix, ImageHeight(ImNum) - 1 - iy) = TB
'               DATACUL2(1, ix, ImageHeight(ImNum) - 1 - iy) = TG
'               DATACUL2(2, ix, ImageHeight(ImNum) - 1 - iy) = TR
'            End Select
'         End If
'      Else     ' Alpha>0 = NA
'
'         NB0 = picDATAREG(0, ix, ImageHeight(ImNum) - 1 - iy)
'         NG0 = picDATAREG(1, ix, ImageHeight(ImNum) - 1 - iy)
'         NR0 = picDATAREG(2, ix, ImageHeight(ImNum) - 1 - iy)
'         NA = picDATAREG(3, ix, ImageHeight(ImNum) - 1 - iy)
'         Ealpha = (NA / 255)
'         NB = TB * (1 - Ealpha) + NB0 * Ealpha
'         NG = TG * (1 - Ealpha) + NG0 * Ealpha
'         NR = TR * (1 - Ealpha) + NR0 * Ealpha
'
'         ' Adjust for tiny Ealpha where NB=TB when NA<>0
'         If NA <> 0 Then
'            If NB = TB Then
'            If NG = TG Then
'            If NR = TR Then
'               If NB0 > TB Then
'                  NB = TB + 1
'               Else
'                  NB = TB - 1
'               End If
'            End If
'            End If
'            End If
'         End If
'         picDATAREG(0, ix, ImageHeight(ImNum) - 1 - iy) = NB
'         picDATAREG(1, ix, ImageHeight(ImNum) - 1 - iy) = NG
'         picDATAREG(2, ix, ImageHeight(ImNum) - 1 - iy) = NR
'
'
'      End If
'   Next ix
'   Next iy
'
'End Sub

Public Sub Reconcile(PIC As PictureBox, ImNum As Long)
' From mnuRotator  Only If ImBPP(ImeNum) = 32 Then
Dim ix As Long, iy As Long
Dim Cul As Long
      
   For iy = LOY To HIY
   For ix = LOX To HIX
      ' Rotator
      Cul = Form1.picSmall(ImNum).Point(ix, ImageHeight(ImNum) - 1 - iy)
      
      Cul = PIC.Point(ix, ImageHeight(ImNum) - 1 - iy)
      If Cul = TColorBGR Then  ' Outside Mask
         Select Case ImNum
         Case 0
            If DATACUL0(3, ix, iy) <> 0 Then
               DATACUL0(3, ix, iy) = 0
            End If
         Case 1
            If DATACUL1(3, ix, iy) <> 0 Then
               DATACUL1(3, ix, iy) = 0
            End If
         Case 2
            If DATACUL2(3, ix, iy) <> 0 Then
               DATACUL2(3, ix, iy) = 0
            End If
         End Select
      Else  ' Inside mask
         Select Case ImNum
         Case 0
            If DATACUL0(3, ix, iy) = 0 Then
               DATACUL0(3, ix, iy) = 1
            End If
         Case 1
            If DATACUL1(3, ix, iy) = 0 Then
               DATACUL1(3, ix, iy) = 1
            End If
         Case 2
            If DATACUL2(3, ix, iy) = 0 Then
               DATACUL2(3, ix, iy) = 1
            End If
         End Select
      End If
   Next ix
   Next iy

End Sub

'Public Sub FillTransparentAreas(Optional Cul As Long = 0)
'' Fill transparent areas of ImageNum with color = Cul
'Dim ix As Long, iy As Long
'Dim Trans As Long
'
'   For iy = 0 To ImageHeight(ImageNum) - 1
'   For ix = 0 To ImageWidth(ImageNum) - 1
'      Trans = 0
'      Select Case ImageNum
'      Case 0
'         If DATACUL0(3, ix, ImageHeight(ImageNum) - 1 - iy) = 0 Then Trans = 1
'      Case 1
'         If DATACUL1(3, ix, ImageHeight(ImageNum) - 1 - iy) = 0 Then Trans = 1
'      Case 2
'         If DATACUL2(3, ix, ImageHeight(ImageNum) - 1 - iy) = 0 Then Trans = 1
'      End Select
'      If Trans = 1 Then
'         CulTo32bppArrays Cul, ix, iy, 0 'AlphaValue=0
'      End If
'   Next ix
'   Next iy
'End Sub
'

Public Sub SwapPictures(k1 As Integer, k2 As Integer, _
   InDATACUL() As Byte, OutDATACUL() As Byte)
' From mnuSwapImages, only called if one or both of IMbpp()s = 32
'Dim AlphaSRC2() As Byte
Dim DATACULSRC2() As Byte
Dim Inbpp As Long, Outbpp As Long

Inbpp = ImBPP(k1)
Outbpp = ImBPP(k2)
   
   If Inbpp = 32 And Outbpp = 32 Then
      ReDim DATACULSRC(0 To 3, ImageWidth(k1) - 1, ImageHeight(k1) - 1)
      ReDim DATACULSRC2(0 To 3, ImageWidth(k2) - 1, ImageHeight(k2) - 1)
      DATACULSRC() = InDATACUL()
      DATACULSRC2() = OutDATACUL()
      
      ReDim OutDATACUL(0 To 3, ImageWidth(k1) - 1, ImageHeight(k1) - 1)
      FILL3D OutDATACUL(), DATACULSRC()
      ReDim InDATACUL(0 To 3, ImageWidth(k2) - 1, ImageHeight(k2) - 1)
      FILL3D InDATACUL(), DATACULSRC2()
   ElseIf Inbpp = 32 Then ' bpp2<>32
      ' move In to Out and cancel In, bpps swapped on return
      ' Redim Out = In
      ReDim OutDATACUL(0 To 3, ImageWidth(k1) - 1, ImageHeight(k1) - 1)
      FILL3D OutDATACUL(), InDATACUL()
   Else ' bpp<>32 bpp2=32
      ' move Out to In cancel Out bpps swapped on return
      ' Redim In = Out
      ReDim InDATACUL(0 To 3, ImageWidth(k2) - 1, ImageHeight(k2) - 1)
      FILL3D InDATACUL(), OutDATACUL()
   End If
End Sub


Public Sub FILL3D(dataDEST() As Byte, dataSRC() As Byte)
Dim ix As Long, iy As Long ', k As Long


'Dim W As Long, H As Long
Dim sw As Long, sh As Long
sw = UBound(dataSRC(), 2)
sh = UBound(dataSRC(), 3)

For iy = 0 To UBound(dataDEST(), 3)
For ix = 0 To UBound(dataDEST(), 2)
'If ix = 22 And iy = 0 Then Stop
'For k = 0 To 3
'   dataDEST(k, ix, iy) = dataSRC(k, ix, iy)
'Next k

   dataDEST(0, ix, iy) = dataSRC(0, ix, iy)
   dataDEST(1, ix, iy) = dataSRC(1, ix, iy)
   dataDEST(2, ix, iy) = dataSRC(2, ix, iy)
   dataDEST(3, ix, iy) = dataSRC(3, ix, iy)


Next ix
Next iy
'   W = UBound(dataDEST(), 2) + 1
'   H = UBound(dataDEST(), 3) + 1
'   CopyMemory dataDEST(0, 0, 0), dataSRC(0, 0, 0), 4 * W * H
'   DoEvents
End Sub

Public Sub Resize3DByteArray(OutArr() As Byte, InArr() As Byte)
' Only called from Adapter Stretch/Shrink
Dim inW As Single, inH As Single
Dim outW As Single, outH As Single
Dim sfH As Single, sfW As Single
Dim ixOut As Long, iyOut As Long
Dim xIn As Long, yIn As Long
   
   ' Actually W-1 & H-1
   inW = UBound(InArr(), 2)
   inH = UBound(InArr(), 3)
   outW = UBound(OutArr(), 2)
   outH = UBound(OutArr(), 3)
   sfH = (inH + 1) / (outH + 1)
   sfW = (inW + 1) / (outW + 1)
   For iyOut = 0 To outH
      yIn = Int(iyOut * sfH)  ' Int is essential
      If yIn <= inH Then
         For ixOut = 0 To outW
            xIn = Int(ixOut * sfW)
            If xIn <= inW Then
               OutArr(0, ixOut, iyOut) = InArr(0, xIn, yIn)
               OutArr(1, ixOut, iyOut) = InArr(1, xIn, yIn)
               OutArr(2, ixOut, iyOut) = InArr(2, xIn, yIn)
               OutArr(3, ixOut, iyOut) = InArr(3, xIn, yIn)
            End If
         Next ixOut
      End If
   Next iyOut
End Sub

Public Sub ScrollLeft32(ImN As Long, ByVal iy As Long)
' Only called for ImBPP()=32
Dim TempCUL(0 To 3)
   
   Select Case ImN
      Case 0
         CopyMemory TempCUL(0), DATACUL0(0, LOX, iy), 4
         CopyMemory DATACUL0(0, LOX, iy), DATACUL0(0, LOX + 1, iy), 4 * (HIX - LOX)
         CopyMemory DATACUL0(0, HIX, iy), TempCUL(0), 4
      Case 1
         CopyMemory TempCUL(0), DATACUL1(0, LOX, iy), 4
         CopyMemory DATACUL1(0, LOX, iy), DATACUL1(0, LOX + 1, iy), 4 * (HIX - LOX)
         CopyMemory DATACUL1(0, HIX, iy), TempCUL(0), 4
      Case 2
         CopyMemory TempCUL(0), DATACUL2(0, LOX, iy), 4
         CopyMemory DATACUL2(0, LOX, iy), DATACUL2(0, LOX + 1, iy), 4 * (HIX - LOX)
         CopyMemory DATACUL2(0, HIX, iy), TempCUL(0), 4
      End Select
End Sub

Public Sub ScrollRight32(ImN As Long, ByVal iy As Long)
' Only called for ImBPP()=32
Dim ix As Long
Dim TempB As Byte
   
   Select Case ImN
   Case 0
      TempB = DATACUL0(0, HIX, iy)
      For ix = HIX - 1 To LOX Step -1
         DATACUL0(0, ix + 1, iy) = DATACUL0(0, ix, iy)
      Next ix
      DATACUL0(0, LOX, iy) = TempB ' To left column
      
      TempB = DATACUL0(1, HIX, iy)
      For ix = HIX - 1 To LOX Step -1
         DATACUL0(1, ix + 1, iy) = DATACUL0(1, ix, iy)
      Next ix
      DATACUL0(1, LOX, iy) = TempB ' To left column
      
      TempB = DATACUL0(2, HIX, iy)
      For ix = HIX - 1 To LOX Step -1
         DATACUL0(2, ix + 1, iy) = DATACUL0(2, ix, iy)
      Next ix
      DATACUL0(2, LOX, iy) = TempB ' To left column
      
      TempB = DATACUL0(3, HIX, iy)
      For ix = HIX - 1 To LOX Step -1
         DATACUL0(3, ix + 1, iy) = DATACUL0(3, ix, iy)
      Next ix
      DATACUL0(3, LOX, iy) = TempB ' To left column
      '----------------------------------------------------
   Case 1
      TempB = DATACUL1(0, HIX, iy)
      For ix = HIX - 1 To LOX Step -1
         DATACUL1(0, ix + 1, iy) = DATACUL1(0, ix, iy)
      Next ix
      DATACUL1(0, LOX, iy) = TempB ' To left column
      
      TempB = DATACUL1(1, HIX, iy)
      For ix = HIX - 1 To LOX Step -1
         DATACUL1(1, ix + 1, iy) = DATACUL1(1, ix, iy)
      Next ix
      DATACUL1(1, LOX, iy) = TempB ' To left column
      
      TempB = DATACUL1(2, HIX, iy)
      For ix = HIX - 1 To LOX Step -1
         DATACUL1(2, ix + 1, iy) = DATACUL1(2, ix, iy)
      Next ix
      DATACUL1(2, LOX, iy) = TempB ' To left column
      
      TempB = DATACUL1(3, HIX, iy)
      For ix = HIX - 1 To LOX Step -1
         DATACUL1(3, ix + 1, iy) = DATACUL1(3, ix, iy)
      Next ix
      DATACUL1(3, LOX, iy) = TempB ' To left column
      '----------------------------------------------------
   Case 2
      TempB = DATACUL2(0, HIX, iy)
      For ix = HIX - 1 To LOX Step -1
         DATACUL2(0, ix + 1, iy) = DATACUL2(0, ix, iy)
      Next ix
      DATACUL2(0, LOX, iy) = TempB ' To left column
      
      TempB = DATACUL2(1, HIX, iy)
      For ix = HIX - 1 To LOX Step -1
         DATACUL2(1, ix + 1, iy) = DATACUL2(1, ix, iy)
      Next ix
      DATACUL2(1, LOX, iy) = TempB ' To left column
      
      TempB = DATACUL2(2, HIX, iy)
      For ix = HIX - 1 To LOX Step -1
         DATACUL2(2, ix + 1, iy) = DATACUL2(2, ix, iy)
      Next ix
      DATACUL2(2, LOX, iy) = TempB ' To left column
      
      TempB = DATACUL2(3, HIX, iy)
      For ix = HIX - 1 To LOX Step -1
         DATACUL2(3, ix + 1, iy) = DATACUL2(3, ix, iy)
      Next ix
      DATACUL2(3, LOX, iy) = TempB ' To left column
      '----------------------------------------------------
   End Select
End Sub

Public Sub ScrollUp32(ImN As Long)
' Only called for ImBPP()=32
Dim iy As Long
   ReDim TempB(0 To ImageWidth(ImN) - 1) As Byte
   ReDim TempCUL(0 To 4 * ImageWidth(ImN) - 1) As Byte

   Select Case ImN
   Case 0
      CopyMemory TempCUL(0), DATACUL0(0, LOX, HIY), 4 * (HIX - LOX + 1) ' Top row
      For iy = HIY To LOY + 1 Step -1
         CopyMemory DATACUL0(0, LOX, iy), DATACUL0(0, LOX, iy - 1), 4 * (HIX - LOX + 1)
      Next iy
      CopyMemory DATACUL0(0, LOX, LOY), TempCUL(0), 4 * (HIX - LOX + 1) ' To Bottom row
   Case 1
      CopyMemory TempCUL(0), DATACUL1(0, LOX, HIY), 4 * (HIX - LOX + 1) ' Top row
      For iy = HIY To LOY + 1 Step -1
         CopyMemory DATACUL1(0, LOX, iy), DATACUL1(0, LOX, iy - 1), 4 * (HIX - LOX + 1)
      Next iy
      CopyMemory DATACUL1(0, LOX, LOY), TempCUL(0), 4 * (HIX - LOX + 1) ' To Bottom row
   Case 2
      CopyMemory TempCUL(0), DATACUL2(0, LOX, HIY), 4 * (HIX - LOX + 1) ' Top row
      For iy = HIY To LOY + 1 Step -1
         CopyMemory DATACUL2(0, LOX, iy), DATACUL2(0, LOX, iy - 1), 4 * (HIX - LOX + 1)
      Next iy
      CopyMemory DATACUL2(0, LOX, LOY), TempCUL(0), 4 * (HIX - LOX + 1) ' To Bottom row
   End Select

End Sub

Public Sub ScrollDown32(ImN As Long)
' Only called for ImBPP()=32
Dim iy As Long
   ReDim TempB(0 To ImageWidth(ImN) - 1) As Byte
   ReDim TempCUL(0 To 4 * ImageWidth(ImN) - 1) As Byte
   Select Case ImN
   Case 0
      CopyMemory TempCUL(0), DATACUL0(0, LOX, LOY), 4 * (HIX - LOX + 1) ' Bottom row
      For iy = LOY To HIY - 1
         CopyMemory DATACUL0(0, LOX, iy), DATACUL0(0, LOX, iy + 1), 4 * (HIX - LOX + 1)
      Next iy
      CopyMemory DATACUL0(0, LOX, HIY), TempCUL(0), 4 * (HIX - LOX + 1) ' To Top row
   Case 1
      CopyMemory TempCUL(0), DATACUL1(0, LOX, LOY), 4 * (HIX - LOX + 1) ' Bottom row
      For iy = LOY To HIY - 1
         CopyMemory DATACUL1(0, LOX, iy), DATACUL1(0, LOX, iy + 1), 4 * (HIX - LOX + 1)
      Next iy
      CopyMemory DATACUL1(0, LOX, HIY), TempCUL(0), 4 * (HIX - LOX + 1) ' To Top row
   Case 2
      CopyMemory TempCUL(0), DATACUL2(0, LOX, LOY), 4 * (HIX - LOX + 1) ' Bottom row
      For iy = LOY To HIY - 1
         CopyMemory DATACUL2(0, LOX, iy), DATACUL2(0, LOX, iy + 1), 4 * (HIX - LOX + 1)
      Next iy
      CopyMemory DATACUL2(0, LOX, HIY), TempCUL(0), 4 * (HIX - LOX + 1) ' To Top row
   End Select

End Sub

Public Sub HorzFlip(ImN As Long)
' Only called for ImBPP()=32
Dim ix As Long, iy As Long
   ReDim DATACULSRC(0 To 3, 0 To ImageWidth(ImN) - 1, 0 To ImageHeight(ImageNum) - 1)
   Select Case ImN
   Case 0
      FILL3D DATACULSRC(), DATACUL0()
   Case 1
      FILL3D DATACULSRC(), DATACUL1()
   Case 2
      FILL3D DATACULSRC(), DATACUL2()
   End Select
   For iy = LOY To HIY
   For ix = LOX To HIX
      Select Case ImN
      Case 0
         DATACUL0(0, HIX + LOX - ix, iy) = DATACULSRC(0, ix, iy)
         DATACUL0(1, HIX + LOX - ix, iy) = DATACULSRC(1, ix, iy)
         DATACUL0(2, HIX + LOX - ix, iy) = DATACULSRC(2, ix, iy)
         DATACUL0(3, HIX + LOX - ix, iy) = DATACULSRC(3, ix, iy)
      Case 1
         DATACUL1(0, HIX + LOX - ix, iy) = DATACULSRC(0, ix, iy)
         DATACUL1(1, HIX + LOX - ix, iy) = DATACULSRC(1, ix, iy)
         DATACUL1(2, HIX + LOX - ix, iy) = DATACULSRC(2, ix, iy)
         DATACUL1(3, HIX + LOX - ix, iy) = DATACULSRC(3, ix, iy)
      Case 2
         DATACUL2(0, HIX + LOX - ix, iy) = DATACULSRC(0, ix, iy)
         DATACUL2(1, HIX + LOX - ix, iy) = DATACULSRC(1, ix, iy)
         DATACUL2(2, HIX + LOX - ix, iy) = DATACULSRC(2, ix, iy)
         DATACUL2(3, HIX + LOX - ix, iy) = DATACULSRC(3, ix, iy)
      End Select
   Next ix
   Next iy
End Sub

Public Sub VertFlip(ImN As Long)
' Only called for ImBPP()=32
Dim ix As Long, iy As Long
   ReDim DATACULSRC(0 To 3, 0 To ImageWidth(ImN) - 1, 0 To ImageHeight(ImageNum) - 1)
   Select Case ImN
   Case 0
      FILL3D DATACULSRC(), DATACUL0()
   Case 1
      FILL3D DATACULSRC(), DATACUL1()
   Case 2
      FILL3D DATACULSRC(), DATACUL2()
   End Select
   For iy = LOY To HIY
   For ix = LOX To HIX
      Select Case ImN
      Case 0
         DATACUL0(0, ix, HIY + LOY - iy) = DATACULSRC(0, ix, iy)
         DATACUL0(1, ix, HIY + LOY - iy) = DATACULSRC(1, ix, iy)
         DATACUL0(2, ix, HIY + LOY - iy) = DATACULSRC(2, ix, iy)
         DATACUL0(3, ix, HIY + LOY - iy) = DATACULSRC(3, ix, iy)
      Case 1
         DATACUL1(0, ix, HIY + LOY - iy) = DATACULSRC(0, ix, iy)
         DATACUL1(1, ix, HIY + LOY - iy) = DATACULSRC(1, ix, iy)
         DATACUL1(2, ix, HIY + LOY - iy) = DATACULSRC(2, ix, iy)
         DATACUL1(3, ix, HIY + LOY - iy) = DATACULSRC(3, ix, iy)
      Case 2
         DATACUL2(0, ix, HIY + LOY - iy) = DATACULSRC(0, ix, iy)
         DATACUL2(1, ix, HIY + LOY - iy) = DATACULSRC(1, ix, iy)
         DATACUL2(2, ix, HIY + LOY - iy) = DATACULSRC(2, ix, iy)
         DATACUL2(3, ix, HIY + LOY - iy) = DATACULSRC(3, ix, iy)
      End Select
   Next ix
   Next iy
End Sub

Public Sub SetAlphas(ix As Long, iy As Long, ABYTE As Byte)
   Select Case ImageNum
   Case 0
      DATACUL0(3, ix, ImageHeight(ImageNum) - 1 - iy) = ABYTE
   Case 1
      DATACUL1(3, ix, ImageHeight(ImageNum) - 1 - iy) = ABYTE
   Case 2
      DATACUL2(3, ix, ImageHeight(ImageNum) - 1 - iy) = ABYTE
   End Select
End Sub

Public Sub CulTo32bppArrays(Cul As Long, ix As Long, iy As Long, ABYTE As Byte)
Dim DR As Byte, DG As Byte, DB As Byte
   LngToRGB Cul, DR, DG, DB
   Select Case ImageNum
   Case 0
      DATACUL0(0, ix, ImageHeight(ImageNum) - 1 - iy) = DB
      DATACUL0(1, ix, ImageHeight(ImageNum) - 1 - iy) = DG
      DATACUL0(2, ix, ImageHeight(ImageNum) - 1 - iy) = DR
      DATACUL0(3, ix, ImageHeight(ImageNum) - 1 - iy) = ABYTE
   Case 1
      DATACUL1(0, ix, ImageHeight(ImageNum) - 1 - iy) = DB
      DATACUL1(1, ix, ImageHeight(ImageNum) - 1 - iy) = DG
      DATACUL1(2, ix, ImageHeight(ImageNum) - 1 - iy) = DR
      DATACUL1(3, ix, ImageHeight(ImageNum) - 1 - iy) = ABYTE
   Case 2
      DATACUL2(0, ix, ImageHeight(ImageNum) - 1 - iy) = DB
      DATACUL2(1, ix, ImageHeight(ImageNum) - 1 - iy) = DG
      DATACUL2(2, ix, ImageHeight(ImageNum) - 1 - iy) = DR
      DATACUL2(3, ix, ImageHeight(ImageNum) - 1 - iy) = ABYTE
   End Select
End Sub

Public Sub StartOnLine()
      If xStart Mod GridMult >= GridMult \ 2 Then
         xStart = (xStart \ GridMult) * GridMult + GridMult
      Else
         xStart = (xStart \ GridMult) * GridMult
      End If
      If yStart Mod GridMult >= GridMult \ 2 Then
         yStart = (yStart \ GridMult) * GridMult + GridMult
      Else
         yStart = (yStart \ GridMult) * GridMult
      End If
      xend = xStart
      yend = yStart
End Sub

Public Sub DrawBox1(frm As Form, PIC As PictureBox, X As Single, Y As Single, Optional MainAlpha As Long = 0)
' Draw shape box
' From picPANEL_MouseMove
' xStart & yStart from picPANEL_MouseDown
' xebd & yend form picPANEL_MouseMove
Dim sw As Long
Dim sh As Long
Dim xs As Long, ys As Long
Dim xe As Long, ye As Long
   
   xs = xStart
   ys = yStart
   xe = X
   ye = Y
      
   If xe > PIC.Width - 1 Then
      xe = PIC.Width
   End If
      
   If xe < xs Then   ' Swap ends
      If xe < 0 Then xe = 0
      X = xe
      xe = xs
      xs = X
      xs = (xs \ GridMult) * GridMult
   Else
      xe = (xe \ GridMult) * GridMult
   End If

   If ye > PIC.Height - 1 Then
      ye = PIC.Height
   End If
   If ye < ys Then   ' Swap ends
      If ye < 0 Then ye = 0
      Y = ye
      ye = ys
      ys = Y
      ys = (ys \ GridMult) * GridMult
   Else
      ye = (ye \ GridMult) * GridMult
   End If
   sw = Abs(xe - xs) '+ GridMult \ 2
   sh = Abs(ye - ys) '+ GridMult \ 2
   If MainAlpha <> 0 Then  ' Alpha draw box
      sw = sw + GridMult '\ 2
      sh = sh + GridMult '\ 2
   End If
   With frm
      .Box1(0).Move xs, ys, sw, sh
      .Box1(1).Move .Box1(0).Left, .Box1(0).Top, .Box1(0).Width, .Box1(0).Height
   End With
End Sub

Private Function Checksize(InSize As Long, iw As Long, ih As Long, ibpp) As Boolean
' InSize = BIH(40)+ num bytes in icon (Pal + color bytes * height)
Dim XW As Long, AW As Long ' XORSize, ANDSize
Dim PalS As Long
Dim CalcSize As Long
Dim k As Long
   Checksize = False
   Select Case ibpp
   Case 1
      XW = ((iw + 7) \ 8 + 3) And &HFFFFFFFC
      PalS = 8
   Case 4
      XW = ((iw + 1) \ 2 + 3) And &HFFFFFFFC
      PalS = 64
   Case 8
      XW = (iw + 3) And &HFFFFFFFC
      PalS = 1024
   Case 24
      k = (3 * iw + 3) And &HFFFFFFFC
      k = k - 3 * iw
      XW = 3 * iw + k
      PalS = 0
   Case 32
      XW = 4 * iw
      PalS = 0
   End Select

   AW = ((iw + 7) \ 8 + 3) And &HFFFFFFFC
   CalcSize = 40 + PalS + (XW + AW) * ih
   If InSize <> CalcSize Then
      Exit Function
   Else
      Checksize = True
   End If
End Function

