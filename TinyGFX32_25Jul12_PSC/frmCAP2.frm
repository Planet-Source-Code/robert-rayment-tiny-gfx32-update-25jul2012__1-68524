VERSION 5.00
Begin VB.Form frmCAP2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawMode        =   7  'Invert
   Icon            =   "frmCAP2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmCAP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'CAPTURE FROM DRAWN RECTANGLE

Option Explicit

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long

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
Private aMouseUp As Boolean
Private sx As Long, sy As Long
Private ex As Long, ey As Long


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' For aspect ratio for Drawn rectangle Capture and Clipboard pasting
Dim CCW As Long, CCH As Long
Dim CCpw As Long
Dim CCph As Long
Dim zAspectCC As Single

   'Swap end points if nec
   If ex < sx Then
      CCW = ex
      ex = sx
      sx = CCW
   End If
   If ey < sy Then
      CCH = ey
      ey = sy
      sy = CCH
   End If
   
   If KeyCode = vbKeySpace Then
      
      '------------------------------
      If AspectNumber = 6 Then   ' Keep aspect ratio
         CCW = Abs(ex - sx)
         CCH = Abs(ey - sy)

         If CCW < 1 Or CCH < 1 Then
            MsgBox " Nothing drawn", vbInformation, "Capture"
            Exit Sub
         End If
         
         zAspectCC = CCW / CCH
         If CCW >= CCH Then  ' zASpectCC >= 1
            CCpw = ImageWidth(ImageNum)
            CCph = ImageWidth(ImageNum) / zAspectCC
         Else  ' W < H   ' zAspect < 1
            CCph = ImageHeight(ImageNum)
            CCpw = ImageHeight(ImageNum) * zAspectCC
         End If
         If CCpw < 8 Or CCph < 8 Then
            MsgBox "Keeping aspect makes width or height < 8", vbInformation, "Clipboard"
            Exit Sub
         End If
         With Form1
            .picSmallBU.Picture = LoadPicture
            .picSmall(ImageNum).Picture = LoadPicture
            .picSmallBU.Width = CCpw
            .picSmallBU.Height = CCph
            .picSmall(ImageNum).Width = CCpw
            .picSmall(ImageNum).Height = CCph
         End With
         ImageWidth(ImageNum) = CCpw
         ImageHeight(ImageNum) = CCph
      End If
      '------------------------------
      
      Form1.picSmall(ImageNum).Picture = LoadPicture
      SetStretchBltMode Form1.picSmall(ImageNum).hdc, HALFTONE 'COLORONCOLOR
      StretchBlt Form1.picSmall(ImageNum).hdc, 0, 0, ImageWidth(ImageNum), ImageHeight(ImageNum), _
         frmCAP2.hdc, sx + 1, sy + 1, Abs(ex - sx) - 1, Abs(ey - sy) - 1, vbSrcCopy
      Form1.picSmall(ImageNum).Picture = Form1.picSmall(ImageNum).Image
      aCAP = True
      ImBPP(ImageNum) = 0
      Unload frmCAP2
      Form1.SetFocus
   ElseIf KeyCode = vbKeyEscape Then
      aCAP = False
      Unload frmCAP2
      Form1.SetFocus
   End If
End Sub

Private Sub Form_Load()
Dim k As Long
      
   frmCAP2.WindowState = vbMaximized
   frmCAP2.Move 0, 0, Screen.Width, Screen.Height
   
   BringWindowToTop frmCAP2.hwnd

   frmCAP2.DrawMode = 7
   aMouseUp = False
   
   frmCAP2.KeyPreview = True
         
   MsgBox "Draw a rectangle with the mouse, then" & vbCrLf & vbCrLf & _
          "Spacebar to Capture  " & vbCrLf & "Esc to Cancel", _
          vbInformation + vbSystemModal, "Capture rectangle"
   DoEvents  ' To ensure MsgBox closes

   ' Here to prevent part of screen blanking out!
   k = GetDC(0)
   BitBlt frmCAP2.hdc, 0, 0, Screen.Width, Screen.Height, _
   k, 0, 0, vbSrcCopy
   frmCAP2.Picture = frmCAP2.Image
   DoEvents
   ReleaseDC 0&, k

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   aMouseDown = True
   
   If aMouseUp Then
      frmCAP2.Line (sx, sy)-(ex, ey), vbWhite, B   ' Clear prev rect
      aMouseUp = False
   End If
   sx = X
   sy = Y
   ex = X
   ey = Y
   frmCAP2.Line (sx, sy)-(ex, ey), vbWhite, B   ' Draw

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   ' Check for Quit
   If GetAsyncKeyState(vbKeyEscape) And &H8000 Then
      aCAP = False
      Unload frmCAP2
      Form1.SetFocus
   End If

   If aMouseDown Then
      frmCAP2.Line (sx, sy)-(ex, ey), vbWhite, B   ' Clear
      ex = X
      ey = Y
      frmCAP2.Line (sx, sy)-(ex, ey), vbWhite, B   ' Draw
   End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   aMouseDown = False
   aMouseUp = True
End Sub
