Attribute VB_Name = "ModColors"
' ModColors.bas

Option Explicit

' 3 vertical picboxes height = 256  (picSlide)
' 3 imageboxes (imSlide)

' Returns RGBSlideVal(Index)

'Public imDown As Boolean
'Public Yim As Single
'Public RGBSlideVal() As Long

'Public pLeftColor As Long
'Public pRightColor As Long
'Public aFillpicRGB As Boolean
Public PALIndex As Integer
Public PalARR() As Byte
Private palRed() As Byte, palGreen() As Byte, palBlue() As Byte
Private Pow2(7) As Byte


Public Sub CenteredPAL(PIC As PictureBox, Optional HV As Long = 0)
' Method from Stefan Casier Paint256
Dim r1 As Long, g1 As Long, b1 As Long
Dim r2 As Long, g2 As Long, b2 As Long
Dim zR As Single, zG As Single, zB As Single
Dim k As Long, k2 As Long, k3 As Long
Dim j As Long
   ReDim PalARR(0 To 2, 0 To 255) As Byte
   ReDim palRed(0 To 16), palGreen(0 To 16), palBlue(0 To 16)
   
   palRed(0) = 0:    palGreen(0) = 0:    palBlue(0) = 254
   palRed(1) = 86:   palGreen(1) = 254:  palBlue(1) = 86
   palRed(2) = 254:  palGreen(2) = 168:  palBlue(2) = 86
   palRed(3) = 92:   palGreen(3) = 0:    palBlue(3) = 0
   palRed(4) = 254:  palGreen(4) = 254:  palBlue(4) = 0
   palRed(5) = 0:    palGreen(5) = 112:  palBlue(5) = 142
   palRed(6) = 254:  palGreen(6) = 254:  palBlue(6) = 254
   palRed(7) = 174:  palGreen(7) = 174:  palBlue(7) = 174
   palRed(8) = 138:  palGreen(8) = 110:  palBlue(8) = 233
   palRed(9) = 0:    palGreen(9) = 101:  palBlue(9) = 0
   palRed(10) = 0:   palGreen(10) = 254: palBlue(10) = 254
   palRed(11) = 254: palGreen(11) = 0:   palBlue(11) = 0
   palRed(12) = 254: palGreen(12) = 254: palBlue(12) = 0
   palRed(13) = 178: palGreen(13) = 0:   palBlue(13) = 178
   palRed(14) = 254: palGreen(14) = 254: palBlue(14) = 254
   palRed(15) = 90:  palGreen(15) = 90:  palBlue(15) = 90
   

   k3 = 0
   For k = 0 To 15
      k2 = k + 1
      If k = 15 Then k2 = 0
      zR = (1& * palRed(k) - palRed(k2)) / 16
      zG = (1& * palGreen(k) - palGreen(k2)) / 16
      zB = (1& * palBlue(k) - palBlue(k2)) / 16
      r2 = palRed(k)
      g2 = palGreen(k)
      b2 = palBlue(k)
      For j = 0 To 14
         r1 = r2 - zR
         g1 = g2 - zG
         b1 = b2 - zB
         If r1 < 0 Then r1 = 255
         If r1 > 255 Then r1 = 0
         r2 = r1
         If g1 < 0 Then g1 = 255
         If g1 > 255 Then g1 = 0
         g2 = g1
         If b1 < 0 Then b1 = 255
         If b1 > 255 Then b1 = 0
         b2 = b1
         PalARR(2, k3) = r1
         PalARR(1, k3) = g1
         PalARR(0, k3) = b1
         k3 = k3 + 1
         If k3 > 255 Then Exit For
      Next j
      If k3 > 255 Then Exit For
   Next k
   For k = 248 To 255
      PalARR(2, k) = 255
      PalARR(1, k) = 255
      PalARR(0, k) = 255
   Next k
   pRGB2PalARR PIC, HV
End Sub

Public Sub ShortBandedPAL(PIC As PictureBox, Optional HV As Long = 0)
' pic = picScreenPAL on frmColors
Dim k As Long
Dim RR As Byte
Dim GG As Byte
Dim BB As Byte
   ReDim PalARR(0 To 2, 0 To 255) As Byte
   ' Banded palette
   RR = 248: GG = 0: BB = 0
   For k = 0 To 31
      PalARR(2, k) = RR
      If RR > 7 Then RR = RR - 8
   Next k
   RR = 0: GG = 244: BB = 0
   For k = 32 To 63
      PalARR(1, k) = GG
      If GG > 7 Then GG = GG - 8
   Next k
   RR = 0: GG = 0: BB = 248
   For k = 64 To 95
      PalARR(0, k) = BB
      If BB > 7 Then BB = BB - 8
   Next k
   RR = 255: GG = 255: BB = 0
   For k = 96 To 127
      PalARR(2, k) = RR
      PalARR(1, k) = GG
      If RR > 7 Then RR = RR - 8
      If GG > 7 Then GG = GG - 8
   Next k
   RR = 248: GG = 0: BB = 248
   For k = 128 To 159
      PalARR(2, k) = RR
      PalARR(0, k) = BB
      If RR > 7 Then RR = RR - 8
      If BB > 7 Then BB = BB - 8
   Next k
   RR = 0: GG = 240: BB = 240
   For k = 160 To 191
      PalARR(1, k) = GG
      PalARR(0, k) = BB
      If GG > 7 Then GG = GG - 8
      If BB > 7 Then BB = BB - 8
   Next k
   RR = 240: GG = 244: BB = 240
   For k = 192 To 223
      PalARR(2, k) = GG
      PalARR(1, k) = GG
      PalARR(0, k) = BB
      If RR > 7 Then RR = RR - 8
      If GG > 7 Then GG = GG - 8
      If BB > 7 Then BB = BB - 8
   Next k
   RR = 248: GG = 248: BB = 248
   For k = 224 To 255
      PalARR(2, k) = RR
      PalARR(1, k) = GG
      PalARR(0, k) = BB
      If GG > 7 Then GG = GG - 8
      If BB > 7 Then BB = BB - 8
   Next k
   For k = 240 To 247
      PalARR(2, k) = 0
      PalARR(1, k) = 0
      PalARR(0, k) = 0
   Next k
   For k = 248 To 255
      PalARR(2, k) = 255
      PalARR(1, k) = 255
      PalARR(0, k) = 255
   Next k
   pRGB2PalARR PIC, HV
End Sub

Public Sub LongBandedPAL(PIC As PictureBox, Optional HV As Long = 0)
Dim k As Long
   ReDim PalARR(0 To 2, 0 To 255) As Byte
   ReDim palRed(0 To 255), palGreen(0 To 255), palBlue(0 To 255)
   ' Banded palette
   For k = 0 To 39
      PalARR(2, k) = 60 + 5 * k
   Next k
   For k = 40 To 79
      PalARR(1, k) = 60 + 5 * (k - 40)
   Next k
   For k = 80 To 119
      PalARR(0, k) = 60 + 5 * (k - 80)
   Next k
   For k = 120 To 159
      PalARR(1, k) = 60 + 5 * (k - 120)
      PalARR(0, k) = 60 + 5 * (k - 120)
   Next k
   
   For k = 160 To 199
      PalARR(2, k) = 60 + 5 * (k - 160)
      PalARR(0, k) = 60 + 5 * (k - 160)
   Next k
   
   For k = 200 To 239
      PalARR(2, k) = 60 + 5 * (k - 200)
      PalARR(1, k) = 60 + 5 * (k - 200)
      PalARR(0, k) = 5 * (k - 200)
   Next k
   For k = 240 To 247
      PalARR(2, k) = 0
      PalARR(1, k) = 0
      PalARR(0, k) = 0
   Next k
   For k = 248 To 255
      PalARR(2, k) = 255
      PalARR(1, k) = 255
      PalARR(0, k) = 255
   Next k
   pRGB2PalARR PIC, HV
End Sub

Public Sub GreyPAL(PIC As PictureBox, Optional HV As Long = 0)
Dim k As Long
   ReDim PalARR(0 To 2, 0 To 255) As Byte
   ' Greyed palette
   For k = 0 To 255
      PalARR(2, k) = k
      PalARR(1, k) = k
      PalARR(0, k) = k
   Next k
   pRGB2PalARR PIC, HV
End Sub

Public Sub QBColors(PIC As PictureBox, Optional HV As Long = 0)
' PIC= picPAL (HV=0)
' HV =0 Horz, 1,2 Vert picbox
Dim k As Long, j As Long
Dim n As Long
   ReDim PalARR(0 To 2, 0 To 255) As Byte
   ' QB palette
   n = 0
   For k = 0 To 255 - 15 Step 16
   For j = k To k + 15
      LngToRGB QBColor(n), PalARR(2, j), PalARR(1, j), PalARR(0, j)
   Next j
   n = n + 1
   Next k
   If HV = 0 Then
      pRGB2PalARR PIC
   Else
      pRGB2PalARRVert PIC, HV
   End If
End Sub

Public Sub pRGB2PalARR(PIC As PictureBox, Optional HV As Long = 0)
' PIC=picPAL(Main palette), picLColor(Text)
Dim k As Long
Dim Mul As Single
   Select Case HV
   Case 0
      Mul = 1.5 ' Widen > 256 = 384
      If PIC.Width < 300 And PIC.Width > 40 Then Mul = 1#  ' For Text palette
      For k = 0 To 255
         PIC.Line (k * Mul, 0)-((k + 1) * Mul, PIC.Height - 1), RGB(PalARR(2, k), PalARR(1, k), PalARR(0, k)), BF
      Next k
   Case 1   ' Vert palette for Alpha edit
      For k = 0 To 255
         PIC.Line (0, k)-(PIC.Width, k), RGB(PalARR(2, k), PalARR(1, k), PalARR(0, k))
      Next k
   Case 2
      For k = 0 To 255
         PIC.Line (0, k / 2)-(PIC.Width, k / 2), RGB(PalARR(2, k), PalARR(1, k), PalARR(0, k))
      Next k
   End Select
   PIC.Refresh
End Sub

Public Sub pRGB2PalARRVert(PIC As PictureBox, Optional HV As Long = 0)
' PIC=picGrid(Grid vert), picVisColor(Visibilty Vert)
Dim k As Long
Dim QR As Long, QG As Long, QB As Long
   If HV = 1 Then
   
      For k = 0 To 255
         'PIC.Line (k, 0)-(k, PIC.Height - 1), RGB(PalARR(2, k), PalARR(1, k), PalARR(0, k))
         PIC.Line (0, k \ 2)-(PIC.Width, k \ 2), RGB(PalARR(2, k), PalARR(1, k), PalARR(0, k))
      Next k
   
   ElseIf HV = 2 Then ' HV=2 Grid colors
      For k = 0 To 255
         QR = PalARR(2, k)
         QG = PalARR(1, k)
         QB = PalARR(0, k)
   
         If QR = 0 And QG = 0 And QB = 0 Then   ' Black
            PIC.Line (0, k \ 2)-(PIC.Width, k \ 2), RGB(64, 64, 64)
         Else
            PIC.Line (0, k \ 2)-(PIC.Width, k \ 2), RGB(QR, QG, QB)
         End If
      Next k
   Else ' HV=3
      For k = 0 To 255
         'PIC.Line (k, 0)-(k, PIC.Height - 1), RGB(PalARR(2, k), PalARR(1, k), PalARR(0, k))
         PIC.Line (0, k)-(PIC.Width, k), RGB(PalARR(2, k), PalARR(1, k), PalARR(0, k))
      Next k
   End If
   PIC.Refresh
End Sub


Public Sub LngToRGB(LCul As Long, r As Byte, g As Byte, B As Byte)
   r = LCul And &HFF&
   g = (LCul And &HFF00&) \ &H100&
   B = (LCul And &HFF0000) \ &H10000
End Sub

Public Function ColorCount(bArr() As Long) As Long
' eg bARR() = picSmallDATA(0 to ImageWidth(ImageNum) - 1, 0 to ImageHeight(ImageNum) - 1)
Dim iy As Long, ix As Long
Dim LCul As Long
Dim BytePos As Long
Dim BitPos As Long
Dim BW As Long ' Width
Dim BH As Long ' Height

On Error GoTo ccError

   BW = UBound(bArr(), 1)   ' = W - 1
   BH = UBound(bArr(), 2)   ' = H - 1
   
   ReDim bColorTable(0 To 16777216 \ 8) As Byte ' 16/8 = 2 MB
   Pow2Tab
   ColorCount = 0
   For iy = 0 To BH
   For ix = 0 To BW
      LCul = bArr(ix, iy) And &HFFFFFF   ' To avoid any alpha
      BytePos = LCul \ 8
      BitPos = LCul And 7
      If (bColorTable(BytePos) And Pow2(BitPos)) = 0 Then
         bColorTable(BytePos) = bColorTable(BytePos) Or Pow2(BitPos)
         ColorCount = ColorCount + 1
      End If
   Next ix
   Next iy
   Erase bColorTable()
   Exit Function
ccError:
   Erase bColorTable()

End Function

Public Sub Pow2Tab()
Dim k As Long
   For k = 0 To 7
      Pow2(k) = 2 ^ k
   Next k
End Sub

