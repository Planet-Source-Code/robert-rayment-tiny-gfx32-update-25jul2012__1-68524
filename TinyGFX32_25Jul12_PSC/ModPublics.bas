Attribute VB_Name = "ModPublics"
'ModPublic.bas

'All sorts!

Option Explicit

Public STX As Long, STY As Long  ' ScreenTwipsPerPixelX/Y

Public aDIBError As Boolean

' Image
Public ImageNum As Long
Public ImageWidth() As Long
Public ImageHeight() As Long

' For shading
Public zRL As Single
Public zGL As Single
Public zBL As Single
Public zRR As Single
Public zGR As Single
Public zBR As Single
Public zRIncr As Single
Public zGIncr As Single
Public zBIncr As Single
Public NSteps As Long

' Preferences                        Menu Num
Public OptimizeNumber As Integer   ' 0,1     Opt or Non-opt saving
Public ToneNumber As Long          ' 3,4     Halftone or ColorOnColor mode
Public AspectNumber As Long        ' 6,7     Keep or ignore aspect
Public MarkPixels As Long          ' 9,10,11 Marked/Marked/Unmarked pixels
Public ExtractorBackups As Long    '13,14    All,Last saved from EXtractor
Public aAlphaRestricted As Byte    '16       True/False Alpha restricted or not
Public aQueryOC As Boolean         '18       True/False Query original colors on saving

' Grid
Public GridMult As Long
Public GridLinesCul As Long

' For Shape Lines & Boxes
' in picPANEL(Main) & picEDIT(Alpha)
Public xStart As Single
Public yStart  As Single
Public xend As Single
Public yend As Single

Public MaxWidth As Long
Public MaxHeight As Long

Public aShow As Boolean

' Capture
Public aCAP As Boolean
Public MagCAP As Long

' Rotator transferred
Public aTransfer As Boolean
Public frmRotateLeft As Long
Public frmRotateTop  As Long

' For cmdSelect button NOT USED
Public acmdSelect As Boolean

' Help file
Public HelpSpec$

' Cursor
Public CurSpec$, CPath$
' Icons & BMPs
Public SaveSpec$, SavePath$
' Input
Public FileSpec$
' Saved file specs
Public NameSpec$(0 To 2)
Public ImFileSpec$(0 To 2)
' General
Public AppPathSpec$, CurrPath$



'#### FILE STUFF ####

Public Sub FixExtension(FSpec$, Ext$)
' In: FileSpec$ & Ext$ (".xxx")
Dim p As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   p = InStr(1, FSpec$, ".")
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      FSpec$ = Mid$(FSpec$, 1, p - 1) & Ext$
   End If
End Sub

Public Function GetFileName(FSpec$) As String
' VB5 also
Dim L As Long
Dim k As Long
   GetFileName = ""
   L = Len(FSpec$)
   If L < 1 Then Exit Function
   For k = L To 1 Step -1
      If Mid$(FSpec$, k, 1) = "\" Then Exit For
   Next k
   If k = 0 Then
      GetFileName = FSpec$
   Else
      GetFileName = Right$(FSpec$, L - k)
   End If
End Function

Public Function FileExists(FSpec$) As Boolean
On Error Resume Next   ' Needed if CD, Zip etc. disk removed
   'If Dir(FSpec$) <> "" Then FileExists = True
   FileExists = LenB(Dir$(FSpec$))
End Function

Public Function GetPath(FSpec$) As String
' VB5 also
Dim L As Long
Dim k As Long
   GetPath = ""
   L = Len(FSpec$)
   If L < 1 Then Exit Function
   For k = L To 1 Step -1
      If Mid$(FSpec$, k, 1) = "\" Then Exit For
   Next k
   If k <> 0 Then
      GetPath = Left$(FSpec$, k)  ' NB includes last \
   End If
End Function

Public Function FindExtension$(FSpec$)
Dim p As Long
   p = FindLastCharPos(FSpec$, ".")
   If p = 0 Then
      FindExtension$ = ""
   Else
      FindExtension$ = Mid$(FSpec$, p + 1)
   End If
End Function

Public Function FindLastCharPos(InString$, SerChar$) As Long
' Also VB5
Dim p As Long
    For p = Len(InString$) To 1 Step -1
      If Mid$(InString$, p, 1) = SerChar$ Then Exit For
    Next p
    If p < 1 Then p = 0
    FindLastCharPos = p
End Function

Public Function ShortenString$(n$, L As Long)
' API for this ?
Dim LName As Long
   'N$ = FindName$(FSpec$)
   LName = Len(n$)
   If LName >= L Then
      ShortenString$ = Left$(n$, L \ 2 - 2) & ".." & Right$(n$, L \ 2)
   Else
      ShortenString$ = n$
   End If
End Function

' Locate

Public Sub LocateForm(frm As Form, FH As Long, FW As Long, fLeft As Long, fTop As Long)
' FH, FW NU
   frm.Left = fLeft
   If frm.Left < 10 Then
      frm.Left = Form1.Left + 30
   End If
   If frm.Left + frm.Width / 2 > Screen.Width Then
      frm.Left = Screen.Width - frm.Width / 2
   End If
   frm.Top = fTop
   If frm.Top < 10 Then
      frm.Top = Form1.Top + 1400
   End If
   If frm.Top > Screen.Height - 1400 Then
      frm.Top = Form1.Top + 1400
   End If
'   wFlags = &H40 Or &H2   ' Colors, Filters, etc
'   SetWindowPos frm.hwnd, hWndInsertAfter, 0, 0, _
'      FW \ STX, FH \ STY, wFlags
End Sub



' Frame mover

'Public Sub fraMOVER(frm As Form, fra As Frame, Button As Integer, _
'   X As Single, Y As Single, Xfra As Single, Yfra As Single)
'Dim fraLeft As Long
'Dim fraTop As Long
'
'   If Button = vbLeftButton Then
'
'      fraLeft = fra.Left + (X - Xfra) \ STX
'      If fraLeft < 0 Then fraLeft = 0
'      If fraLeft + fra.Width > frm.Width \ STX Then
'         fraLeft = frm.Width \ STX - fra.Width - 4
'      End If
'      fra.Left = fraLeft
'
'      fraTop = fra.Top + (Y - Yfra) \ STY
'      If fraTop < 0 Then fraTop = 0
'      If fraTop + fra.Height > frm.Height \ STY Then
'         fraTop = frm.Height \ STY - fra.Height - 8
'      End If
'      fra.Top = fraTop
'
'   End If
'End Sub

