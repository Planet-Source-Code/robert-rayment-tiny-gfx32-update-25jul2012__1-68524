Attribute VB_Name = "ModSave256BMP"
' ModSave256BMP.bas

Option Explicit

Private Type BITMAPFILEHEADER   ' For 8bpp
    bType             As Integer  ' BM        2
    bSize             As Long     ' 54+1024 + (iw + Mod(4))*ih  FileSize B
    bReserved1        As Integer  ' 0         2
    bReserved2        As Integer  ' 0         2
    bOffBits          As Long     ' 54+1024      4
    bHeaderSize       As Long     ' 40        4
    bWidth            As Long     ' iw         4
    bHeight           As Long     ' ih         4
    bNumPlanes        As Integer  ' 1         2
    bBPP              As Integer  ' 8         2
    bCompress         As Long     ' 0         4
    bBytesInImage     As Long     ' (iw + Mod(4))*ih  4  Image size B
    bHRES             As Long     ' 0 ignore  4
    bVRES             As Long     ' 0 ignore  4
    bUsedIndexes      As Long     ' 0 ignore  4
    bImportantIndexes As Long     ' 0 ignore  4  Total = 62
End Type

'Private Const BI_RLE4 As Long = 2&
'Private Const BI_RLE8 As Long = 1&

Private bARRIndexes() As Byte

Public Function SaveBMP256(FSpec$, bArr() As Byte, bWidth As Long, bHeight As Long, _
                gPAL() As Long) As Boolean
' Save 2 color BMP
' EG bArr() = Public picDATAREG(0 To 3, 0 To iw - 1, 0 To ih - 1)

Dim BFH As BITMAPFILEHEADER ' 54 bytes
Dim fnum As Integer
Dim BytesPerScanLine As Long
Dim ix As Long, iy As Long
Dim Cul1 As Long
Dim k1 As Long

On Error GoTo SaveBMPError256

   BytesPerScanLine = (bWidth + 3) And &HFFFFFFFC
   
   With BFH
      .bType = &H4D42    ' BM
      .bWidth = bWidth
      .bHeight = bHeight
      .bSize = 54 + 1024 + BytesPerScanLine * Abs(bHeight)
      .bOffBits = 54 + 1024
      .bHeaderSize = 40
      .bNumPlanes = 1
      .bBPP = 8
      .bCompress = 0
      .bBytesInImage = BytesPerScanLine * Abs(bHeight)
   End With
   ' 8bpp image map.  Fill from 32bpp image map.
   ReDim bARRIndexes(0 To BytesPerScanLine - 1, 0 To Abs(bHeight) - 1)
   
   ' gPal() & bARR()  are BGRA
   
   For iy = 0 To bHeight - 1
   For ix = 0 To bWidth - 1
      Cul1 = RGB(bArr(0, ix, iy), bArr(1, ix, iy), bArr(2, ix, iy))
      ' Find Index to Cul1 in gPAL()
      For k1 = 0 To 255
         If Cul1 = gPAL(k1) Then Exit For
      Next k1
      If k1 = 256 Then k1 = 0 ' Color not found
      bARRIndexes(ix, iy) = CByte(k1)
   
   Next ix
   Next iy
   
   '-- Kill any previous
   If FileExists(FSpec$) Then
      Kill FSpec$
   End If
   
   ' Save File
   fnum = FreeFile
   Open FSpec$ For Binary As fnum
   Put #fnum, , BFH
   Put #fnum, , gPAL()
   Put #fnum, , bARRIndexes()
   Close #fnum
   Erase bARRIndexes()
   SaveBMP256 = True
   On Error GoTo 0
   Exit Function
'=======
SaveBMPError256:
   Close
   SaveBMP256 = False
End Function


