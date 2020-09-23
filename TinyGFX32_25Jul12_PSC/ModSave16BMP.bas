Attribute VB_Name = "ModSave16BMP"
' ModSave2BMP.bas

Option Explicit

Private Type BITMAPFILEHEADER   ' For 4bpp
    bType             As Integer  ' BM        2
    bSize             As Long     ' 54+64 + (iw\2 + Mod(4))*ih  FileSize B
    bReserved1        As Integer  ' 0         2
    bReserved2        As Integer  ' 0         2
    bOffBits          As Long     ' 54+64     4
    bHeaderSize       As Long     ' 40        4
    bWidth            As Long     ' iw         4
    bHeight           As Long     ' ih         4
    bNumPlanes        As Integer  ' 1         2
    bBPP              As Integer  ' 4         2
    bCompress         As Long     ' 0         4
    bBytesInImage     As Long     ' (iw\2 + Mod(4))*ih  4  Image size B
    bHRES             As Long     ' 0 ignore  4
    bVRES             As Long     ' 0 ignore  4
    bUsedIndexes      As Long     ' 0 ignore  4
    bImportantIndexes As Long     ' 0 ignore  4  Total = 62
End Type

'Private Const BI_RLE4 As Long = 2&
'Private Const BI_RLE8 As Long = 1&

Private bARRIndexes() As Byte

Public Function SaveBMP16(FSpec$, bArr() As Byte, bWidth As Long, bHeight As Long, _
                gPAL() As Long) As Boolean
' Save 2 color BMP
' EG bArr() = Public picDATAREG(0 To 3, 0 To iw - 1, 0 To ih - 1)

Dim BFH As BITMAPFILEHEADER ' 54 bytes
Dim fnum As Integer
Dim BytesPerScanLine As Long
Dim ix As Long, iy As Long
Dim Cul1 As Long, Cul2 As Long
Dim k1 As Long, k2 As Long
Dim Pal16(0 To 15) As Long
Dim n As Long
'On Error GoTo SaveBMPError16

   BytesPerScanLine = ((bWidth + 1) \ 2 + 3) And &HFFFFFFFC
   
   With BFH
      .bType = &H4D42    ' BM
      .bWidth = bWidth
      .bHeight = bHeight
      .bSize = 54 + 64 + BytesPerScanLine * Abs(bHeight)
      .bOffBits = 54 + 64
      .bHeaderSize = 40
      .bNumPlanes = 1
      .bBPP = 4
      .bCompress = 0
      .bBytesInImage = BytesPerScanLine * Abs(bHeight)
   End With
   ' 4bpp image map.  Fill from 32bpp image map.
   ReDim bARRIndexes(0 To BytesPerScanLine - 1, 0 To Abs(bHeight) - 1)
   
   ' gPal() & bARR()  are BGRA
   
   For iy = 0 To bHeight - 1
   n = 0
   For ix = 0 To bWidth - 1 Step 2
      Cul1 = RGB(bArr(0, ix, iy), bArr(1, ix, iy), bArr(2, ix, iy))
      ' Find Index to Cul1 in gPAL()
      For k1 = 0 To 15
         If Cul1 = gPAL(k1) Then Exit For
      Next k1
      If k1 = 16 Then k1 = 0 ' Color not found
      
      If ix + 1 < bWidth Then
         Cul2 = RGB(bArr(0, ix + 1, iy), bArr(1, ix + 1, iy), bArr(2, ix + 1, iy))
         ' Find Index to Cul2 in gPAL()
         For k2 = 0 To 15
            If Cul2 = gPAL(k2) Then Exit For
         Next k2
         If k2 = 16 Then k2 = 0 ' Color not found
      Else
         k2 = 0
      End If
      ' k1 = 0-15  k2= 0-15
      bARRIndexes(n, iy) = 16 * CByte(k1) + CByte(k2)
      n = n + 1
   Next ix
   Next iy
   
   For k1 = 0 To 15
      Pal16(k1) = gPAL(k1)
   Next k1
   
   '-- Kill any previous
   If FileExists(FSpec$) Then
      Kill FSpec$
   End If
   
   ' Save file
   fnum = FreeFile
   Open FSpec$ For Binary As fnum
   Put #fnum, , BFH
   Put #fnum, , Pal16()
   Put #fnum, , bARRIndexes()
   Close #fnum
   Erase bARRIndexes()
   SaveBMP16 = True
   On Error GoTo 0
   Exit Function
'=======
SaveBMPError16:
   Close
   SaveBMP16 = False
End Function


