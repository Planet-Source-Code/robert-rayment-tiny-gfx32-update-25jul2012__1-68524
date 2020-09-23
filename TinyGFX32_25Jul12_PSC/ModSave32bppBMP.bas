Attribute VB_Name = "ModSave32bppBMP"
'ModSave32bppBMP.bas

Option Explicit

Private Type BITMAPFILEHEADER   ' For 32bpp
    bType             As Integer  ' BM        2
    bSize             As Long     ' 54+4*iw*ih
    bReserved1        As Integer  ' 0         2
    bReserved2        As Integer  ' 0         2
    bOffBits          As Long     ' 54        4
    bHeaderSize       As Long     ' 40        4
    bWidth            As Long     ' iw        4
    bHeight           As Long     ' ih        4
    bNumPlanes        As Integer  ' 1         2
    bBPP              As Integer  ' 32        2
    bCompress         As Long     ' 0         4
    bBytesInImage     As Long     ' (4*iw*ih) 4  Image size B
    bHRES             As Long     ' 0 ignore  4
    bVRES             As Long     ' 0 ignore  4
    bUsedIndexes      As Long     ' 0 ignore  4
    bImportantIndexes As Long     ' 0 ignore  4  Total = 62
End Type

Public Function Save32bppBMP(FSpec$, iw As Long, ih As Long) As Boolean
'picDATAREG(0 To 3, iw - 1, ih - 1) contains the BGRA

Dim BFH As BITMAPFILEHEADER ' 54 bytes
Dim fnum As Integer

   On Error GoTo Save32bppBMPError

   With BFH
      .bType = &H4D42    ' BM
      .bWidth = iw
      .bHeight = ih
      .bSize = 54 + iw * Abs(ih)
      .bOffBits = 54
      .bHeaderSize = 40
      .bNumPlanes = 1
      .bBPP = 32
      .bCompress = 0
      .bBytesInImage = iw * Abs(ih)
   End With

   '-- Kill any previous
   If FileExists(FSpec$) Then
      Kill FSpec$
   End If
   
'Dim ix As Long, iy As Long
'Dim k As Long
'Dim BB0 As Byte, BB1 As Byte, BB2 As Byte, BB3 As Byte
'For iy = 0 To ih - 1
'For ix = 0 To 4 * iw - 1 Step 4
'   BB0 = picDATAREG(0, ix, iy)
'   BB1 = picDATAREG(1, ix, iy)
'   BB2 = picDATAREG(2, ix, iy)
'   BB3 = picDATAREG(3, ix, iy)
'Next ix
'Next iy

   
   ' Save File
   fnum = FreeFile
   Open FSpec$ For Binary As fnum
   Put #fnum, , BFH
   Select Case ImageNum
   Case 0: Put #fnum, , DATACUL0()
   Case 1: Put #fnum, , DATACUL1()
   Case 2: Put #fnum, , DATACUL2()
   End Select
   'Put #fnum, , picDATAREG()
   Close #fnum
   DoEvents
   Save32bppBMP = True
   On Error GoTo 0
   Exit Function
'=======
Save32bppBMPError:
   Close
   Save32bppBMP = False
End Function


