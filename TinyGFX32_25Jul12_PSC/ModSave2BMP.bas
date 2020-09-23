Attribute VB_Name = "ModSave2BMP"
' ModSave2BMP.bas

Option Explicit

Private Type BITMAPFILEHEADER   ' For 1bpp
    bType             As Integer  ' BM        2
    bSize             As Long     ' 54+8 + (iw\8 + Mod(4))*ih  FileSize B
    bReserved1        As Integer  ' 0         2
    bReserved2        As Integer  ' 0         2
    bOffBits          As Long     ' 54+8      4
    bHeaderSize       As Long     ' 40        4
    bWidth            As Long     ' iw         4
    bHeight           As Long     ' ih         4
    bNumPlanes        As Integer  ' 1         2
    bBPP              As Integer  ' 1         2
    bCompress         As Long     ' 0         4
    bBytesInImage     As Long     ' (iw\8 + Mod(4))*ih  4  Image size B
    bHRES             As Long     ' 0 ignore  4
    bVRES             As Long     ' 0 ignore  4
    bUsedIndexes      As Long     ' 0 ignore  4
    bImportantIndexes As Long     ' 0 ignore  4  Total = 62
End Type

'Private Const BI_RLE4 As Long = 2&
'Private Const BI_RLE8 As Long = 1&

Private bARRIndexes() As Byte
Private Pal() As Long

Public Function SaveBMP2(FSpec$, bArr() As Byte, bWidth As Long, bHeight As Long, _
                Optional Color1 As Long = vbBlack, Optional Color2 As Long = vbWhite) As Boolean
' Save 2 color BMP
' EG bArr() = Public PaperData(0 To 3, 0 To iw - 1, 0 To ih - 1)

Dim BFH As BITMAPFILEHEADER ' 54 bytes
Dim fnum As Integer
Dim BytesPerScanLine As Long
Dim ix As Long, iy As Long

Dim B As Long, n As Long

Dim Cul As Long
Dim CR As Byte, CG As Byte, CB As Byte

On Error GoTo SaveBMPError

'   ' Expand to 4B boundary
   BytesPerScanLine = ((bWidth + 7) \ 8 + 3) And &HFFFFFFFC
   
   With BFH
      .bType = &H4D42    ' BM
      .bWidth = bWidth
      .bHeight = bHeight
      .bSize = 54 + 8 + BytesPerScanLine * Abs(bHeight)
      .bOffBits = 54 + 8
      .bHeaderSize = 40
      .bNumPlanes = 1
      .bBPP = 1
      .bCompress = 0
      .bBytesInImage = BytesPerScanLine * Abs(bHeight)
   End With
   ' 1bpp image map.  Fill from 32bpp image map.
   ReDim bARRIndexes(0 To BytesPerScanLine - 1, 0 To Abs(bHeight) - 1)
   ' all zero ie Color1
   ReDim Pal(0 To 1)
   ' Switch RGB to BGR for palette
   LngToRGB Color1, CR, CG, CB
   Pal(0) = RGB(CB, CG, CR)
   LngToRGB Color2, CR, CG, CB
   Pal(1) = RGB(CB, CG, CR)

   'eg  0 0 0 0 255 255 255 255  Index 0(Black) & 1(White)
   
   For iy = 0 To bHeight - 1
   n = 0
   B = 0
   For ix = 0 To bWidth - 1
      Cul = RGB(bArr(2, ix, iy), bArr(1, ix, iy), bArr(0, ix, iy))
      If Cul = Color2 Then
      'If bArr(0, ix, iy) > 250 Then
         bARRIndexes(n, iy) = bARRIndexes(n, iy) Or 1
         If B < 7 Then bARRIndexes(n, iy) = bARRIndexes(n, iy) * 2
      Else
         If B < 7 Then bARRIndexes(n, iy) = bARRIndexes(n, iy) * 2
      End If
      B = B + 1
      If B = 8 Then
         n = n + 1
         B = 0
      End If
   Next ix
   If B <> 0 Then
      bARRIndexes(n, iy) = bARRIndexes(n, iy) * 2 ^ (7 - B)
   End If
   Next iy
   
   '-- Kill previous
   If FileExists(FSpec$) Then
      Kill FSpec$
   End If
   
   ' bARRIndexes() & PAL() could be Input to GIF save
   ' width = BytesPerScanLine, height = bHeight +/- ?
   
   fnum = FreeFile
   Open FSpec$ For Binary As fnum
   Put #fnum, , BFH
   Put #fnum, , Pal()
   Put #fnum, , bARRIndexes()
   Close #fnum
   Erase bARRIndexes()
   SaveBMP2 = True
   On Error GoTo 0
   Exit Function
'=======
SaveBMPError:
   Close
   SaveBMP2 = False
End Function


