Attribute VB_Name = "ModText"
' ModText.bas
Option Explicit

' Text

Public Type TEXTMETRIC
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
End Type

Public Declare Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsA" _
   (ByVal hdc As Long, ByRef lpMetrics As TEXTMETRIC) As Long

Public tm As TEXTMETRIC
Public tLeading As Long

Public Const LF_FACESIZE = 32
Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
'        lfFaceName(LF_FACESIZE - 1) As Byte
        lfFaceName As String * LF_FACESIZE
End Type

'Public Const FE_FONTSMOOTHINGCLEARTYPE As Long = &H2
Public Const NONANTIALIASED_QUALITY As Byte = 3
'Public Const ANTIALIASED_QUALITY As Byte = 4
Public Const CLEARTYPE_QUALITY As Byte = 5 '6

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" _
(lpLogFont As LOGFONT) As Long
Public LFONT As LOGFONT

Public Declare Function SelectObject Lib "gdi32" _
(ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" _
(ByVal hObject As Long) As Long

Public aText As Boolean
Public aTextOK As Boolean
Public TFontname As String
Public TFontsize As Long
Public TFontBold As Boolean
Public TFontItalic As Boolean
Public aClearType  As Boolean
Public TextLine$
Public minx As Long, miny As Long, maxx As Long, maxy As Long

'-------------------------------------------------
' For Text

Public Sub ShowText(PIC As PictureBox, TextLine$)
' PIC picSmallText
Dim rfont As Long
Dim TheCurfont As Long

   LFONT.lfEscapement = 0
   With LFONT
      If PIC.FontBold Then .lfWeight = 1200 / 2 Else .lfWeight = 400 / 2
      If PIC.FontItalic Then .lfItalic = 1 Else .lfItalic = 0
      .lfStrikeOut = 0
      .lfUnderline = 0
      .lfCharSet = 1   '0,1
      .lfFaceName = PIC.FontName & Chr$(0)
      .lfHeight = (PIC.FontSize * -20) / STY
      .lfWidth = 0 '(PIC.FontSize) * -7 / STX
      '.lfQuality = FE_FONTSMOOTHINGCLEARTYPE ' ClearType XP
      '.lfQuality = ANTIALIASED_QUALITY
      '.lfQuality = NONANTIALIASED_QUALITY  ' Temp disable any ClearType
      If aClearType Then
         .lfQuality = CLEARTYPE_QUALITY
      Else
         .lfQuality = NONANTIALIASED_QUALITY
      End If

   End With
   '------------------------------------
   rfont = CreateFontIndirect(LFONT)
   TheCurfont = SelectObject(PIC.hdc, rfont)
   PIC.Print TextLine$;
   PIC.Refresh
   'Restore CurFont
   SelectObject PIC.hdc, TheCurfont
   DeleteObject rfont
   '------------------------------------
End Sub

Public Sub FindMaxMins(PIC As PictureBox, BackCul As Long)
' PIC = picSmallText from frmText
'Public minx,miny,maxx,maxy
Dim ix As Long, iy As Long

   minx = 1000
   miny = 1000
   maxx = -1000
   maxy = -1000
   
   For iy = 0 To PIC.Height - 1
   For ix = 0 To PIC.Width - 1
      If PIC.Point(ix, iy) <> BackCul Then
         If ix < minx Then minx = ix
         If iy < miny Then miny = iy
         If ix > maxx Then maxx = ix
         If iy > maxy Then maxy = iy
      End If
   Next
   Next
End Sub



