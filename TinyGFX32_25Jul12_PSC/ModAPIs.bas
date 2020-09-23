Attribute VB_Name = "ModAPI"
' ModAPIs.bas

' NB Not all used but apparently will not go into EXE if not used

Option Explicit
' frmCAP, frmCAP2
Public Declare Function BringWindowToTop Lib "USER32" (ByVal hwnd As Long) As Long

' For finding Clipbrd.exe
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" _
   (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' For Fillers
Public Const FLOODFILLBORDER = 0 ' Fill until crColor& color encountered.
Public Const FLOODFILLSURFACE = 1 ' Fill surface until crColor& color not encountered.
Public Declare Function ExtFloodFill Lib "gdi32.dll" _
   (ByVal hdc As Long, _
   ByVal X As Long, ByVal Y As Long, _
   ByVal crColor As Long, _
   ByVal wFillType As Long) As Long

Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Declare Sub SetCursorPos Lib "USER32" (ByVal ix As Long, ByVal iy As Long)
Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Public Type POINTAPI
    kx As Long
    ky As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    bottom As Long
End Type

Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long

' On top
'Public Declare Function SetWindowPos Lib "USER32" (ByVal hwnd As Long, _
'   ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
'   ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long
'
'Public Const hWndInsertAfter = -1
'Public Const SWP_NOMOVE = 2
'Public Const SWP_NOSIZE = 1
'''Public Const SWP_SHOWWINDOW As Long = &H40
'''Public Const SWP_NOACTIVATE As Long = &H10
''Public Const wFlags = SWP_NOMOVE Or SWP_NOSIZE    ' New
''Public Const wFlags = &H40 Or &H2   ' Colors
''Public Const HWND_TOPMOST = -1
'Public wFlags As Long

Public Declare Function GetDC Lib "USER32" (ByVal hwnd As Long) As Long

' For saving selection
Public Declare Function BitBlt Lib "gdi32.dll" _
 (ByVal hDestDC As Long, _
 ByVal xDest As Long, ByVal yDest As Long, _
 ByVal nWidth As Long, ByVal nHeight As Long, _
 ByVal hSrcDC As Long, _
 ByVal xSrc As Long, ByVal ySrc As Long, _
 ByVal dwRop As Long) As Long

Public Type BITMAPINFOHEADER ' 40 bytes
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
Public Declare Function GetDIBits Lib "gdi32.dll" _
   (ByVal aHDC As Long, ByVal hBitmap As Long, _
   ByVal nStartScan As Long, ByVal nNumScans As Long, _
   ByRef lpBits As Any, _
   ByRef BInfo As BITMAPINFOHEADER, _
   ByVal wUsage As Long) As Long

Public Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type
  
Public Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
   
' For display & zooming
Public Declare Function SetStretchBltMode Lib "gdi32.dll" _
(ByVal hdc As Long, ByVal nStretchMode As Long) As Long

Public Declare Function GetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long) As Long

Public Const HALFTONE As Long = 4
Public Const COLORONCOLOR As Long = 3

Public Declare Function StretchDIBits Lib "gdi32.dll" _
   (ByVal hdc As Long, _
   ByVal X As Long, ByVal Y As Long, _
   ByVal dx As Long, ByVal DY As Long, _
   ByVal SrcX As Long, ByVal SrcY As Long, _
   ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, _
   ByRef lpBits As Any, _
   ByRef BInfo As BITMAPINFOHEADER, _
   ByVal wUsage As Long, _
   ByVal dwRop As Long) As Long
   
'Public Declare Function SetDIBitsToDevice Lib "gdi32" _
'   (ByVal hDC As Long, _
'   ByVal x As Long, ByVal y As Long, _
'   ByVal dx As Long, ByVal DY As Long, _
'   ByVal SrcX As Long, ByVal SrcY As Long, _
'   ByVal Scan As Long, ByVal NumScans As Long, _
'   lpBits As Any, _
'   BInfo As BITMAPINFOHEADER, _
'   ByVal wUsage As Long) As Long
   
Public Declare Function SetDIBits Lib "gdi32.dll" _
   (ByVal hdc As Long, ByVal hBitmap As Long, _
    ByVal nStartScan As Long, ByVal nNumScans As Long, _
    ByRef lpBits As Any, ByRef lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
   

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
(ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

'Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" _
'(ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Public Declare Function StretchBlt Lib "gdi32.dll" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
 ByVal nWidth As Long, ByVal nHeight As Long, _
 ByVal hSrcDC As Long, _
 ByVal xSrc As Long, ByVal ySrc As Long, _
 ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
 ByVal dwRop As Long) As Long

Public Declare Function SetPixelV Lib "gdi32.dll" _
   (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Declare Function GetPixel Lib "gdi32.dll" _
   (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

' wUsage
'Public Const DIB_RGB_COLORS As Long = 0
'Public Const DIB_PAL_COLORS As Long = 1
' eg dwRop
' vbSrcCopy = &H00CC0020

