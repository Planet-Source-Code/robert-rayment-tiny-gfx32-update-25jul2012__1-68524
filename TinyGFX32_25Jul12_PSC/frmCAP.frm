VERSION 5.00
Begin VB.Form frmCAP 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "cc"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCAP.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMAG 
      AutoRedraw      =   -1  'True
      Height          =   1050
      Left            =   1200
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   0
      Top             =   285
      Width           =   1050
   End
End
Attribute VB_Name = "frmCAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long


Private Declare Function BitBlt Lib "gdi32.dll" _
(ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
 ByVal nWidth As Long, ByVal nHeight As Long, _
 ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
 ByVal dwRop As Long) As Long
 
Private Declare Function GetAsyncKeyState Lib "USER32" _
   (ByVal vKey As KeyCodeConstants) As Long

'Public MagCAP As Long
' & all the ImageNum variables
Private aMouseDown As Boolean

Private HalfW As Long, HalfH As Long


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim PT As POINTAPI
   aMouseDown = False
   GetCursorPos PT
   If KeyCode = vbKeySpace Then
      Form1.picSmall(ImageNum).Picture = LoadPicture
      SetStretchBltMode Form1.picSmall(ImageNum).hdc, HALFTONE 'COLORONCOLOR
      BitBlt Form1.picSmall(ImageNum).hdc, 0, 0, ImageWidth(ImageNum), ImageHeight(ImageNum), _
         frmCAP.hdc, PT.kx - HalfW, PT.ky - HalfH, vbSrcCopy

      Form1.picSmall(ImageNum).Picture = Form1.picSmall(ImageNum).Image
      aCAP = True
      ImBPP(ImageNum) = 0
      Unload frmCAP
      Form1.SetFocus
   ElseIf KeyCode = vbKeyEscape Then
      aCAP = False
      Unload frmCAP
      Form1.SetFocus
   End If
End Sub

Private Sub Form_Load()
Dim k As Long
      
   frmCAP.WindowState = vbMaximized
   frmCAP.Move 0, 0, Screen.Width, Screen.Height
   
   BringWindowToTop frmCAP.hwnd
   
'   k = GetDC(0)
'   BitBlt frmCAP.hdc, 0, 0, Screen.Width, Screen.Height, _
'   k, 0, 0, vbSrcCopy
'   frmCAP.Picture = frmCAP.Image
'   DoEvents
'   ReleaseDC 0&, k
'
'   HalfW = ImageWidth(ImageNum) \ 2
'   HalfH = ImageHeight(ImageNum) \ 2
'
'   picMAG.Move Screen.Width \ (4 * STX), Screen.Width \ (4 * STY), _
'            MagCAP * ImageWidth(ImageNum) + 2, MagCAP * ImageHeight(ImageNum) + 2
   
   frmCAP.KeyPreview = True
   
   MsgBox "Move cursor over screen to  " & vbCrLf & "show captured area, then" & vbCrLf & vbCrLf & _
          "Spacebar to Capture  " & vbCrLf & "Esc to Cancel  " & vbCrLf & vbCrLf & _
          "(Note that the Window can be moved " & vbCrLf & "   out of the way with the mouse)", _
          vbInformation + vbSystemModal, "Capture image"
   DoEvents  ' To ensure MsgBox closes
   
   k = GetDC(0)
   BitBlt frmCAP.hdc, 0, 0, Screen.Width, Screen.Height, _
   k, 0, 0, vbSrcCopy
   frmCAP.Picture = frmCAP.Image
   DoEvents
   ReleaseDC 0&, k
   
   HalfW = ImageWidth(ImageNum) \ 2
   HalfH = ImageHeight(ImageNum) \ 2
   
   picMAG.Move Screen.Width \ (4 * STX), Screen.Width \ (4 * STY), _
            MagCAP * ImageWidth(ImageNum) + 2, MagCAP * ImageHeight(ImageNum) + 2

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
   i = Button
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PT As POINTAPI
Dim PLeft As Long, PTop As Long
Dim xDest As Long, yDest As Long

   ' Check for Quit
   If GetAsyncKeyState(vbKeyEscape) And &H8000 Then
      aCAP = False
      Unload frmCAP
      Form1.SetFocus
   End If

   GetCursorPos PT
   PLeft = PT.kx - HalfW
   PTop = PT.ky - HalfH
   picMAG.Picture = LoadPicture
   xDest = 0
   yDest = 0
   ' Avoid top & left reflections
   If PTop < 0 Then
      yDest = -MagCAP * PTop
      PTop = 0
   End If
   If PLeft < 0 Then
      xDest = -MagCAP * PLeft
      PLeft = 0
   End If
   StretchBlt picMAG.hdc, xDest - 2, yDest - 2, picMAG.Width, picMAG.Height, _
   frmCAP.hdc, PLeft, PTop, ImageWidth(ImageNum), ImageHeight(ImageNum), vbSrcCopy
   picMAG.Picture = picMAG.Image
End Sub

Private Sub picMAG_KeyDown(KeyCode As Integer, Shift As Integer)
Dim PT As POINTAPI
   aMouseDown = False
   GetCursorPos PT
   If KeyCode = vbKeySpace Then
      Form1.picSmall(ImageNum).Picture = LoadPicture
      SetStretchBltMode Form1.picSmall(ImageNum).hdc, 3
      BitBlt Form1.picSmall(ImageNum).hdc, 0, 0, ImageWidth(ImageNum), ImageHeight(ImageNum), _
         frmCAP.hdc, PT.kx - HalfW, PT.ky - HalfH, vbSrcCopy
      
      Form1.picSmall(ImageNum).Picture = Form1.picSmall(ImageNum).Image
      aCAP = True
      Unload frmCAP
      Form1.SetFocus
   ElseIf KeyCode = vbKeyEscape Then
      aCAP = False
      Unload frmCAP
      Form1.SetFocus
   End If

End Sub

Private Sub picMAG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   aMouseDown = True
End Sub

Private Sub picMAG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aMouseDown Then
      ReleaseCapture
      SendMessage picMAG.hwnd, WM_NCLBUTTONDOWN, 2, 0&
   End If
End Sub

Private Sub picMAG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   aMouseDown = False
End Sub

