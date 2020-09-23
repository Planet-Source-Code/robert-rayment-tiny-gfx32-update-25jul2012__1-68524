Attribute VB_Name = "ModTestCursor"
' ModTestCursor.bas

Option Explicit

Private Declare Function SetClassLong Lib "USER32" Alias "SetClassLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetCursor Lib "USER32" () As Long

Private Declare Function LoadCursorFromFile Lib "USER32" Alias "LoadCursorFromFileA" _
   (ByVal lpFileName As String) As Long

Private Declare Function DestroyCursor Lib "USER32" (ByVal hCursor As Long) As Long

Private Declare Function CopyImage Lib "USER32" _
   (ByVal handle As Long, _
   ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Const GCL_HCURSOR As Long = -12
Private Const LR_COPYDELETEORG = &H8
Private Const IMAGE_CURSOR = 2
'Private Const OCR_NORMAL As Long = 32512


Public currenthcurs As Long
Public SYStempcurs  As Long
Public tempcurs1 As Long
Public tempcurs2 As Long
Public newhcurs As Long
Public aTestCursor As Boolean
Public TestCursNum As Long

Public Sub SaveSystemCursor()
    currenthcurs = GetCursor()
    SYStempcurs = CopyImage(currenthcurs, IMAGE_CURSOR, 0, 0, LR_COPYDELETEORG)
End Sub

'' WILL RESTORE AN ANIMATED CURSOR BUT TEST ANI WILL ONLY SHOW
'' OVER Nominated controls.  NB this will do all picboxes on Form frm
Public Sub ShowNewCursor(frm As Form, PIC As PictureBox, FilePath As String)
    newhcurs = LoadCursorFromFile(FilePath)
    tempcurs1 = SetClassLong(frm.hwnd, GCL_HCURSOR, newhcurs)
    tempcurs2 = SetClassLong(PIC.hwnd, GCL_HCURSOR, newhcurs)
End Sub
'
Public Sub RestoreOldCursor(frm As Form, PIC As PictureBox)
    SetClassLong frm.hwnd, GCL_HCURSOR, tempcurs1
    SetClassLong PIC.hwnd, GCL_HCURSOR, tempcurs2
    DestroyCursor tempcurs1
    DestroyCursor tempcurs2
End Sub

