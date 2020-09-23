Attribute VB_Name = "Mod256Colors"
' Mod256Colors.bas

'   From cPalette.Cls
'   Copyright © 1999 Steve McMahon
'
'   The Octree Colour Quantisation Code (CreateOptimal) was written by
'   Brian Schimpf Copyright © 1999 Brian Schimpf

' Also eTools2cted & adapted from 'Ordered dither test"
' by Carles P V (PSC CodeId=49875) which also
' contains full references, caviats and options.

' This is a simplified & modified version by RR

Option Explicit

Private Type PALETTEENTRY
    peR     As Byte
    peG     As Byte
    peB     As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE256
    palVersion       As Integer
    palNumEntries    As Integer
    palPalEntry(0 To 255) As PALETTEENTRY
End Type
Private logpal256 As LOGPALETTE256

Private Type tNode    'OCT-TREE node struct.
   bIsLeaf            As Boolean ' Leaf flag
   bAddedReduce       As Boolean ' Linked list flag
   vR                 As Long    ' Red Value
   vG                 As Long    ' Green Value
   vB                 As Long    ' Blue Value
   cClrs              As Long    ' Pixel count
   iChildren(0 To 1, 0 To 1, 0 To 1) As Long   ' Child pointers
   iNext              As Long    ' Next reducable node
End Type

' Octree variables
Private aNodes()   As tNode
Private cNodes     As Long
Private nDepth     As Byte
Private TopGarbage As Long
Private cClr       As Long
Private aReduce()  As Long

' For logical palette
Private NumEntries  As Integer
Private hPal     As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
 (Destination As Any, Source As Any, ByVal Length As Long)

' For building RGB4096Inv LUT
Private Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As Long, ByVal crColor As Long) As Long
' Look Up Tables
Private RGB4096Inv(0 To &HF, 0 To &HF, 0 To &HF) As Byte ' RGB4096 palette Inverse index LUT
Private RGB4096Trn() As Long ' RGB4096 Translation LUT
Private m_ODM_O(0 To 7, 0 To 7) As Long ' Ordered dither matrix (Bayer 8x8)


Public gifPAL() As Long
Public bA8() As Byte

Public Sub GetPalette(bA32() As Byte)
' bA32() = picDATAREG(0 To 3, 0 To ImageWidth(ImageNum), ImageHeight(ImageNum))
' TO  Public gifPAL(0 To 255) As Long
Dim gRED(0 To 255) As Byte
Dim gGREEN(0 To 255) As Byte
Dim gBLUE(0 To 255) As Byte
Dim NumColors As Long
Dim NumLevels As Long
  
Dim gWidth As Long
Dim gHeight As Long
Dim ix As Long
Dim iy As Long
Dim k As Long

   'Screen.MousePointer = vbHourglass
   
   ReDim gifPAL(0 To 255)

   gWidth = UBound(bA32, 2)      ' FullW-1
   gHeight = UBound(bA32, 3)     ' FullH-1

   NumColors = 256
   NumLevels = 8

   '-- Allocates initial storage
   ReDim aNodes(1 To 50) As tNode
   ReDim aReduce(1 To 8) As Long
   nDepth = NumLevels
   cNodes = 1
   TopGarbage = 0
   cClr = 0

   For iy = 0 To gHeight
      For ix = 0 To gWidth
         '-- Adds the current pixel to the color octree
         Call pvAddClr(1, 1, 0, 255, 0, 255, 0, 255, _
                       bA32(0, ix, iy), bA32(1, ix, iy), bA32(2, ix, iy))
         '-- Combine the levels to get down to desired palette size
         Do While (cClr > NumColors)
             If (pvCombineNodes = 0) Then Exit Do
         Loop
      Next ix
   Next iy
   '-- Go through octree and eTools2ct colors
   k = 0
   For iy = 1 To UBound(aNodes)
      If (aNodes(iy).bIsLeaf) Then
         With aNodes(iy)
            gRED(k) = .vR / .cClrs
            gGREEN(k) = .vG / .cClrs
            gBLUE(k) = .vB / .cClrs
            k = k + 1
         End With
      End If
   Next iy
   NumEntries = k
   ' Transfer colors to gifPAL(0-255) As Long
   'Public gifPAL(0 To 255) As Long
   For k = 0 To 255
      gifPAL(k) = RGB(gRED(k), gGREEN(k), gBLUE(k))
   Next k
   
   'Screen.MousePointer = vbDefault
End Sub

Public Sub Get_Pal_Indexes(bA32() As Byte, NumColors As Long, Optional aTransfer As Boolean = False)
' EG
' Get_Pal_Indexes picDATAREG()
' bA32() = picDATA()
' aTransfer = False convert bA32()
'           = True  indexes to bA8()

' Output Public gifPAL(0 to 255) Longs                            ' 1024 palette
'        Public bA8(0 To UBound(bA32, 2), 0 To UBound(bA32, 3))   ' Indexes(bytes) to palette


Dim gRED(0 To 255) As Byte
Dim gGREEN(0 To 255) As Byte
Dim gBLUE(0 To 255) As Byte

Dim gINDEX As Long
Dim RB As Byte, GB As Byte, BB As Byte

Dim ix As Long
Dim iy As Long
Dim k As Long

' For making RGB4096Trn LUT
Dim aBPP As Byte     ' Color depth
Dim LDW  As Single   ' Dither weight
Dim TARR As Variant  ' Bayer array
Dim LIdx As Long
Dim LOff As Long
   
' For building RGB4096Inv LUT
Dim r As Long
Dim g As Long
Dim B As Long
Dim CulR As Long
   
' Can be inputted
Dim NColors As Long
Dim NumLevels As Long
  
Dim gWidth As Long
Dim gHeight As Long

' Ordered Dither
Dim lODx   As Long
Dim lODy   As Long
Dim lODInc As Long


   'Screen.MousePointer = vbHourglass
   
   ReDim gifPAL(0 To 255)

   gWidth = UBound(bA32, 2)      ' ImageWidth(ImageNum) - 1
   gHeight = UBound(bA32, 3)     ' ImageHeight(ImageNum) - 1

   'NumColors = 256 '128,64,32,16
   NColors = NumColors - 1
   NumLevels = 16 '8

   '-- Allocates initial storage
   ReDim aNodes(1 To 50) As tNode
   'ReDim aReduce(1 To 8) As Long
   ReDim aReduce(1 To NumLevels) As Long
   nDepth = NumLevels
   cNodes = 1
   TopGarbage = 0
   cClr = 0

   'StartProgress (gHeight)
   For iy = 0 To gHeight
      For ix = 0 To gWidth
         '-- Adds the current pixel to the color octree
         Call pvAddClr(1, 1, 0, 255, 0, 255, 0, 255, _
                       bA32(0, ix, iy), bA32(1, ix, iy), bA32(2, ix, iy))
         '-- Combine the levels to get down to desired palette size
         Do While (cClr > NColors)
             If (pvCombineNodes = 0) Then Exit Do
         Loop
      Next ix
   Next iy
   '-- Go through octree and eTools2ct colors
   k = 0
   For iy = 1 To UBound(aNodes)
      If (aNodes(iy).bIsLeaf) Then
         With aNodes(iy)
            gRED(k) = .vR / .cClrs
            gGREEN(k) = .vG / .cClrs
            gBLUE(k) = .vB / .cClrs
            k = k + 1
         End With
      End If
   Next iy
   NumEntries = k
   ' Transfer colors to Public gifPAL(0-255) As Long
'   If aICO Then
'      For k = 0 To 255
'         gifPAL(k) = RGB(gBLUE(k), gGREEN(k), gRED(k))
'      Next k
'   End If
   For k = 0 To 255
      gifPAL(k) = RGB(gRED(k), gGREEN(k), gBLUE(k))
   Next k
   '--------------------------------------------------------------------------
   ' RR mod
    If NumEntries <= 256 Then ' No dithering necessary, in fact makes a worse image
      ' Make logical palette
         If hPal <> 0 Then
            DeleteObject hPal
            hPal = 0
         End If
         ' Force 256
         NumEntries = 256
         With logpal256
            .palNumEntries = NumEntries
            .palVersion = &H300
            CopyMemory .palPalEntry(0), gifPAL(0), 1024
         End With
         hPal = CreatePalette(logpal256)
         
         ReDim bA8(0 To gWidth, 0 To gHeight)
         For iy = 0 To gHeight
         For ix = 0 To gWidth
            CulR = RGB(picDATAREG(0, ix, iy), picDATAREG(1, ix, iy), picDATAREG(2, ix, iy))
            bA8(ix, iy) = CByte(GetNearestPaletteIndex(hPal, CulR))
         Next ix
         Next iy
         ' Delete logical palette
         If hPal <> 0 Then
            DeleteObject hPal
            hPal = 0
         End If
         'Screen.MousePointer = vbDefault
         Exit Sub
    End If
   '--------------------------------------------------------------------------
   ' Make RGB4096Trn LUT
   aBPP = 8
   '--------------------------------------------------------------------------
   LDW = (&H11 * (9 - aBPP)) / &H32       ' LDW=17/50 = .34
   
   LOff = 25 * LDW + 1                    ' Loff=9.5 = 10
   ReDim RGB4096Trn(-LOff To &HFF + LOff) '-10 To 265
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '-- Ordered dither matrix (Bayer 8x8. [-25, 25])
   LIdx = 0
   TARR = Array(0, 38, 9, 47, 2, 40, 11, 50, 25, 12, 35, 22, 27, 15, 37, 24, 6, _
   44, 3, 41, 8, 47, 5, 43, 31, 19, 28, 15, 34, 21, 31, 18, 1, 39, _
   11, 49, 0, 39, 10, 48, 27, 14, 36, 23, 26, 13, 35, 23, 7, 46, 4, _
   43, 7, 45, 3, 42, 33, 20, 30, 17, 32, 19, 29, 16)
   For ix = 0 To 7
   For iy = 0 To 7
      m_ODM_O(ix, iy) = LDW * (TARR(LIdx) - 25): LIdx = LIdx + 1
   Next iy
   Next ix
   
   For LIdx = -LOff To &HFF + LOff
      RGB4096Trn(LIdx) = (LIdx + &H8) \ &H11
      If (RGB4096Trn(LIdx) < &H0) Then RGB4096Trn(LIdx) = &H0
      If (RGB4096Trn(LIdx) > &HF) Then RGB4096Trn(LIdx) = &HF
   Next LIdx
   
   ' Make logical palette
   If hPal <> 0 Then
      DeleteObject hPal
      hPal = 0
   End If
   ' Force 256
   NumEntries = 256
   With logpal256
      .palNumEntries = NumEntries
      .palVersion = &H300
      CopyMemory .palPalEntry(0), gifPAL(0), 1024
   End With
   hPal = CreatePalette(logpal256)
   
   ' Build_RGB4096Inv LUT
   '-- Build 4096-colors palette inverse indexes LUT
   For r = 0 To &HF
   For g = 0 To &HF
   For B = 0 To &HF
      ' OK
      CulR = (B + 256& * g + 65536 * r) * &H11
      RGB4096Inv(r, g, B) = CByte(GetNearestPaletteIndex(hPal, CulR))
   Next B
   Next g
   Next r
   '------------------------------------------------------------

   If Not aTransfer Then
      
      For iy = 0 To gHeight
         For ix = 0 To gWidth
            '--  Get index
            lODInc = m_ODM_O(lODx, lODy)
            gINDEX = _
            RGB4096Inv(RGB4096Trn(bA32(2, ix, iy) + lODInc), _
            RGB4096Trn(bA32(1, ix, iy) + lODInc), _
            RGB4096Trn(bA32(0, ix, iy) + lODInc))
            
            LngToRGB gifPAL(gINDEX), BB, GB, RB   ' OK
            
            bA32(0, ix, iy) = BB
            bA32(1, ix, iy) = GB
            bA32(2, ix, iy) = RB
            '-- Inc. ord. matrix column
            lODx = lODx + 1
            If (lODx = 8) Then lODx = 0
         Next ix
         '-- Inc. ord. matrix row
         lODx = 0
         lODy = lODy + 1
         If (lODy = 8) Then lODy = 0
      Next iy
   
   Else      ' Transfer indexes to bA8()
      ReDim bA8(0 To gWidth, 0 To gHeight)
      For iy = 0 To gHeight
         For ix = 0 To gWidth
            '--  Get index
            lODInc = m_ODM_O(lODx, lODy)
            bA8(ix, iy) = _
               RGB4096Inv(RGB4096Trn(bA32(2, ix, iy) + lODInc), _
               RGB4096Trn(bA32(1, ix, iy) + lODInc), _
               RGB4096Trn(bA32(0, ix, iy) + lODInc))
            '-- Inc. ord. matrix column
            lODx = lODx + 1
            If (lODx = 8) Then lODx = 0
         Next ix
         '-- Inc. ord. matrix row
         lODx = 0
         lODy = lODy + 1
         If (lODy = 8) Then lODy = 0
         'DrawProgress
      Next iy
   End If
   
   ' Delete logical palette
   If hPal <> 0 Then
      DeleteObject hPal
      hPal = 0
   End If
End Sub


'========================================================================================
' Private   From Carles P V
'========================================================================================

Private Sub pvAddClr(ByVal iBranch As Long, ByVal nLevel As Long, _
                     ByVal vMinR As Byte, ByVal vMaxR As Byte, _
                     ByVal vMinG As Byte, ByVal vMaxG As Byte, _
                     ByVal vMinB As Byte, ByVal vMaxB As Byte, _
                     ByVal r As Byte, ByVal g As Byte, ByVal B As Byte)

' <Recursive>
' Adds a color to the OctTree palette.
' Will call itself if not in correct level.
'
' Inputs:
'  - iBranch (Branch to look down)
'  - nLevel (Current level (depth) in tree)
'  - vMin(R, G, B) (The minimum branch value)
'  - vMax(R, G, B) (The maximum branch value)
'  - R, G, B (The Red, Green, and Blue color components)
  
  Dim iR As Byte, iG As Byte, iB As Byte
  Dim vMid As Byte, iIndex As Long
    
    '-- Find mid values for colors and decide which path to take.
    '   Also update max and min values for later call to self.
    
    vMid = vMinR / 2 + vMaxR / 2
    If (r > vMid) Then iR = 1: vMinR = vMid Else iR = 0: vMaxR = vMid

    vMid = vMinG / 2 + vMaxG / 2
    If (g > vMid) Then iG = 1: vMinG = vMid Else iG = 0: vMaxG = vMid

    vMid = vMinB / 2 + vMaxB / 2
    If (B > vMid) Then iB = 1: vMinB = vMid Else iB = 0: vMaxB = vMid
    
    '-- If no child here then...
    If (aNodes(iBranch).iChildren(iR, iG, iB) = 0) Then
        '-- Get a new node index
        iIndex = pvGetFreeNode
        aNodes(iBranch).iChildren(iR, iG, iB) = iIndex
        aNodes(iBranch).cClrs = aNodes(iBranch).cClrs + 1
        '-- Clear/set data
        With aNodes(iIndex)
            .bIsLeaf = (nLevel = nDepth)
            .iNext = 0
            .cClrs = 0
            .vR = 0
            .vG = 0
            .vB = 0
        End With
      Else
        '-- Has a child here
        iIndex = aNodes(iBranch).iChildren(iR, iG, iB)
    End If
 
    '-- If it is a leaf
    If (aNodes(iIndex).bIsLeaf) Then
        With aNodes(iIndex)
            If (.cClrs = 0) Then cClr = cClr + 1
            .cClrs = .cClrs + 1
            .vR = .vR + r
            .vG = .vG + g
            .vB = .vB + B
        End With
    Else
        With aNodes(iIndex)
            '-- If 2 or more colors, add to reducable aNodes list
            If (.bAddedReduce = 0) Then
                .iNext = aReduce(nLevel)
                 aReduce(nLevel) = iIndex
                .bAddedReduce = -1
            End If
        End With
        '-- Search a level deeper
        Call pvAddClr(iIndex, nLevel + 1, vMinR, vMaxR, vMinG, vMaxG, vMinB, vMaxB, r, g, B)
    End If
End Sub

Private Function pvCombineNodes() As Boolean

' Combines octree aNodes to reduce the count of colors.
' Combines all children of a leaf into itself.
  
  Dim i As Long, iIndex As Long
  Dim iR As Byte, iG As Byte, iB As Byte
  Dim NR As Long, NG As Long, NB As Long, nPixel As Long

    '-- Find deepest reducable level
    For i = nDepth To 1 Step -1
        If (aReduce(i) <> 0) Then Exit For
    Next i

    If (i = 0) Then Exit Function
    iIndex = aReduce(i)
    aReduce(i) = aNodes(iIndex).iNext

    For i = 0 To 7
        If (i And 1) = 1 Then iR = 1 Else iR = 0
        If (i And 2) = 2 Then iG = 1 Else iG = 0
        If (i And 4) = 4 Then iB = 1 Else iB = 0
        
        '-- If there is a child
        If (aNodes(iIndex).iChildren(iR, iG, iB) <> 0) Then
            With aNodes(aNodes(iIndex).iChildren(iR, iG, iB))
                '-- Add red, green, blue, and pixel count to running total
                NR = NR + .vR
                NG = NG + .vG
                NB = NB + .vB
                nPixel = nPixel + .cClrs
                '-- Free the node
                Call pvFreeNode(aNodes(iIndex).iChildren(iR, iG, iB))
                cClr = cClr - 1
            End With
            '-- Clear the link
            aNodes(iIndex).iChildren(iR, iG, iB) = 0
        End If
    Next i
    cClr = cClr + 1

    '-- Set the new node data
    With aNodes(iIndex)
        .cClrs = nPixel
        .bIsLeaf = -1
        .vR = NR
        .vG = NG
        .vB = NB
    End With
    pvCombineNodes = -1
End Function

Private Sub pvFreeNode(ByVal iNode As Long)

' Puts a node on the top of the garbage list.
' Inputs:
'  - iNode
'  - Index of node to free
    
    aNodes(iNode).iNext = TopGarbage
    TopGarbage = iNode
    aNodes(iNode).bIsLeaf = 0 ' Necessary for final loop through
    aNodes(iNode).bAddedReduce = 0
    cNodes = cNodes - 1
End Sub

Private Function pvGetFreeNode() As Long

' pvGetFreeNode: Gets a new node index from the trash list or the
' end of the list. Clears child pointers.
' Outputs:
'  - Node index
  
  Dim i  As Long
  Dim iR As Byte
  Dim iG As Byte
  Dim iB As Byte
  
    cNodes = cNodes + 1
    If (TopGarbage = 0) Then
        If (cNodes > UBound(aNodes)) Then
            i = cNodes * 1.1
            ReDim Preserve aNodes(1 To i)
        End If
        pvGetFreeNode = cNodes
      Else
        pvGetFreeNode = TopGarbage
        TopGarbage = aNodes(TopGarbage).iNext
        For i = 0 To 7
            If (i And 1) = 1 Then iR = 1 Else iR = 0
            If (i And 2) = 2 Then iG = 1 Else iG = 0
            If (i And 4) = 4 Then iB = 1 Else iB = 0
            aNodes(pvGetFreeNode).iChildren(iR, iG, iB) = 0
        Next i
    End If
End Function
'========================================================================================

