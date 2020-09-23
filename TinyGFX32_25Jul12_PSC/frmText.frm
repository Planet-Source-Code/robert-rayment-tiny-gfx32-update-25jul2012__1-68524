VERSION 5.00
Begin VB.Form frmText 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Text"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   188
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkClearType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ClearType"
      Height          =   225
      Left            =   1170
      TabIndex        =   14
      Top             =   105
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Delete text"
      Height          =   300
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   60
      Width           =   945
   End
   Begin VB.CommandButton cmdFontTable 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Font Table ->"
      Height          =   300
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   60
      Width           =   1140
   End
   Begin VB.PictureBox picLColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   10
      ToolTipText     =   " Change Left color "
      Top             =   2025
      Width           =   3870
   End
   Begin VB.PictureBox picSmallCopy 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   3510
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   9
      Top             =   780
      Width           =   990
   End
   Begin VB.PictureBox picSmallText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1950
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   8
      Top             =   780
      Width           =   990
   End
   Begin VB.CommandButton cmdACCCAN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2385
      Width           =   1245
   End
   Begin VB.CommandButton cmdACCCAN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Accept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   705
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2385
      Width           =   1245
   End
   Begin VB.CommandButton cmdLRUD 
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Index           =   3
      Left            =   1665
      Picture         =   "frmText.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1125
      Width           =   195
   End
   Begin VB.CommandButton cmdLRUD 
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Index           =   2
      Left            =   1665
      Picture         =   "frmText.frx":0092
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   750
      Width           =   195
   End
   Begin VB.CommandButton cmdLRUD 
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   1
      Left            =   2280
      Picture         =   "frmText.frx":0124
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   510
      Width           =   345
   End
   Begin VB.CommandButton cmdLRUD 
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   0
      Left            =   1905
      Picture         =   "frmText.frx":0196
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   510
      Width           =   345
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   180
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   630
      Width           =   1065
   End
   Begin VB.CommandButton cmdSelectFont 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select Font"
      Height          =   300
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   1050
   End
   Begin VB.Line Line1 
      X1              =   1
      X2              =   320
      Y1              =   26
      Y2              =   26
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Left color"
      Height          =   450
      Left            =   4155
      TabIndex        =   15
      Top             =   1995
      Width           =   435
   End
   Begin VB.Label LabImageNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      ToolTipText     =   " Image Number "
      Top             =   780
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   1650
      Picture         =   "frmText.frx":0208
      ToolTipText     =   " Scrollers - or use arrow keys "
      Top             =   510
      Width           =   210
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmText.frm

Option Explicit

Private Temp() As Long

Dim CurFont As StdFont

Private Sub Form_Load()
   LocateForm frmText, 3210, 4905, 1000, 1000

   picSmallText.Height = ImageHeight(ImageNum) '+ 2
   picSmallText.Width = ImageWidth(ImageNum) '+ 2
   picSmallCopy.Height = ImageHeight(ImageNum) '+ 2
   picSmallCopy.Width = ImageWidth(ImageNum) '+ 2
   
   picSmallText.BackColor = vbWhite
   picSmallText.ForeColor = vbBlack
   Text1.BackColor = vbWhite
   Text1.ForeColor = LColor
   If LColor = vbWhite Then
      Text1.BackColor = vbBlack
   End If
   Text1.Refresh
   Set CurFont = New StdFont
   Text1.Text = ""
   With Text1
      .FontName = TFontname
      .FontSize = TFontsize
      .FontBold = TFontBold
      .FontItalic = TFontItalic
   End With
   With picSmallText
      .FontName = TFontname
      .FontSize = TFontsize
      .FontBold = TFontBold
      .FontItalic = TFontItalic
   End With
   With picSmallCopy
      .FontName = TFontname
      .FontSize = TFontsize
      .FontBold = TFontBold
      .FontItalic = TFontItalic
   End With
   Caption = "Text - " & TFontname$ & " Size " & Str$(TFontsize) & " Regular"
   aTextOK = False
   
   picLColor.Width = 258
   Select Case PALIndex
   Case 0: QBColors picLColor
   Case 1: ShortBandedPAL picLColor
   Case 2: LongBandedPAL picLColor
   Case 3: GreyPAL picLColor
   Case 4: CenteredPAL picLColor
   End Select

   LabImageNumber.Caption = ImageNum + 1
   
   picSmallCopy.Cls
   
   BitBlt picSmallCopy.hdc, 0, 0, picSmallCopy.ScaleWidth, picSmallCopy.ScaleHeight, _
    Form1.picSmall(ImageNum).hdc, 0, 0, vbSrcCopy
   picSmallCopy.Picture = picSmallCopy.Image
   
   If chkClearType.Value = Checked Then
      aClearType = True
   Else
      aClearType = False
   End If

   minx = 0
   miny = 0
End Sub

Private Sub chkClearType_Click()
   If chkClearType.Value = Checked Then
      aClearType = True
   Else
      aClearType = False
   End If
   minx = 0
   miny = 0
   Text1_Change
   Text1.SetFocus
   Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub cmdACCCan_Click(Index As Integer)
Dim ix As Long, iy As Long
Dim Cul As Long
Dim CR As Byte, CG As Byte, CB As Byte

   aTextOK = False
   If Len(TextLine$) <> 0 Then
   If Index = 0 Then     ' Accept
      If CurrentGen(ImageNum) = 0 Then Form1.BackUp ImageNum
      BitBlt Form1.picSmall(ImageNum).hdc, 0, 0, picSmallCopy.ScaleWidth, picSmallCopy.ScaleHeight, _
         picSmallCopy.hdc, 0, 0, vbSrcCopy
      Form1.picSmall(ImageNum).Picture = Form1.picSmall(ImageNum).Image
      ' alpha picSmalltext where = not white make alpha=255
      If ImBPP(ImageNum) = 32 Then
         LngToRGB LColor, CR, CG, CB
         For iy = 0 To ImageHeight(ImageNum) - 1
         For ix = 0 To ImageWidth(ImageNum) - 1
            Cul = picSmallText.Point(ix, iy)
            If Cul <> vbWhite Then
               Select Case ImageNum
               Case 0:
                  DATACUL0(0, ix, ImageHeight(ImageNum) - 1 - iy) = CB
                  DATACUL0(1, ix, ImageHeight(ImageNum) - 1 - iy) = CG
                  DATACUL0(2, ix, ImageHeight(ImageNum) - 1 - iy) = CR
                  DATACUL0(3, ix, ImageHeight(ImageNum) - 1 - iy) = 255
               Case 1:
                  DATACUL1(0, ix, ImageHeight(ImageNum) - 1 - iy) = CB
                  DATACUL1(1, ix, ImageHeight(ImageNum) - 1 - iy) = CG
                  DATACUL1(2, ix, ImageHeight(ImageNum) - 1 - iy) = CR
                  DATACUL1(3, ix, ImageHeight(ImageNum) - 1 - iy) = 255
               Case 2:
                  DATACUL2(0, ix, ImageHeight(ImageNum) - 1 - iy) = CB
                  DATACUL2(1, ix, ImageHeight(ImageNum) - 1 - iy) = CG
                  DATACUL2(2, ix, ImageHeight(ImageNum) - 1 - iy) = CR
                  DATACUL2(3, ix, ImageHeight(ImageNum) - 1 - iy) = 255
               End Select
            End If
         Next ix
         Next iy
      End If
      aTextOK = True
   End If
   End If
   Unload Me
End Sub

Private Sub cmdClear_Click()
   TextLine$ = ""
   Text1.SetFocus
   Text1.Text = ""
   picSmallText.Picture = LoadPicture
   picSmallText.Cls  ' Set CurrentX,Y = 0
End Sub

Private Sub cmdFontTable_Click()
   frmFont.Show vbModal
   Text1.SetFocus
   Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub cmdLRUD_Click(Index As Integer)
'Private Temp() As Long
Dim BMIH As BITMAPINFOHEADER
Dim ix As Long, iy As Long

   If TextLine$ = "" Then Exit Sub
   
   GetTextMetrics frmText.picSmallText.hdc, tm
   tLeading = tm.tmExternalLeading 'tm.tmInternalLeading '+ tm.tmExternalLeading

   With BMIH
      .biSize = 40
      .biPlanes = 1
      .biWidth = ImageWidth(ImageNum)
      .biHeight = ImageHeight(ImageNum)
      .biBitCount = 32
      '.biSizeImage = 4 * W * H
   End With

   ' Transfer pixel colors to picSmallDATA() from picSmallText
   '===========================================================
   GetTheBitsLong picSmallText
   '===========================================================
    
   If aDIBError Then Exit Sub
  
   ReDim Temp(0 To ImageWidth(ImageNum) - 1)
   ' Public minx,miny,maxx,maxy
   ' Find minx,miny,maxx,maxy from picSmallText
   ' and then check minx > 0 maxx < picSmallText.Width-1
   '                miny > 0 maxy < picSmallText.Height-1
 
   FindMaxMins picSmallText, vbWhite
   
   picSmallCopy.CurrentX = picSmallText.CurrentX
   picSmallCopy.CurrentY = picSmallText.CurrentY
   
   Select Case Index
   Case 0   ' Move picSmallText Left
      If minx > 0 Then
         For iy = 0 To ImageHeight(ImageNum) - 1
            Temp(0) = picSmallDATA(0, iy)   ' Left column
            CopyMemory picSmallDATA(0, iy), picSmallDATA(1, iy), 4 * (ImageWidth(ImageNum) - 1)
            picSmallDATA(ImageWidth(ImageNum) - 1, iy) = Temp(0)  ' To right column
         Next iy
         minx = minx - 1
         picSmallText.CurrentX = picSmallText.CurrentX - 1
      End If
      
   Case 1   ' Move picSmallText Right
      If maxx < picSmallText.Width - 1 Then
         For iy = 0 To ImageHeight(ImageNum) - 1
            Temp(0) = picSmallDATA(ImageWidth(ImageNum) - 1, iy) ' Right column
            For ix = ImageWidth(ImageNum) - 2 To 0 Step -1
               picSmallDATA(ix + 1, iy) = picSmallDATA(ix, iy)
            Next ix
            picSmallDATA(0, iy) = Temp(0)  ' To left column
         Next iy
         maxx = maxx + 1
         picSmallText.CurrentX = picSmallText.CurrentX + 1
      End If
   Case 2   ' Move picSmallText Up
      If miny > tLeading Then
         CopyMemory Temp(0), picSmallDATA(0, ImageHeight(ImageNum) - 1), 4 * ImageWidth(ImageNum) ' Top row
         For iy = ImageHeight(ImageNum) - 1 To 1 Step -1
            CopyMemory picSmallDATA(0, iy), picSmallDATA(0, iy - 1), 4 * ImageWidth(ImageNum)
         Next iy
         CopyMemory picSmallDATA(0, 0), Temp(0), 4 * ImageWidth(ImageNum)  ' To Bottom row
         miny = miny - 1
         picSmallText.CurrentY = picSmallText.CurrentY - 1
      End If
   Case 3   ' Move picSmallText Down
      If maxy < picSmallText.Height - 1 Then
         CopyMemory Temp(0), picSmallDATA(0, 0), 4 * ImageWidth(ImageNum) ' Bottom row
         For iy = 0 To ImageHeight(ImageNum) - 2
            CopyMemory picSmallDATA(0, iy), picSmallDATA(0, iy + 1), 4 * ImageWidth(ImageNum)
         Next iy
         CopyMemory picSmallDATA(0, ImageHeight(ImageNum) - 1), Temp(0), 4 * ImageWidth(ImageNum) ' To Top row
         maxy = maxy + 1
         picSmallText.CurrentY = picSmallText.CurrentY + 1
      End If
   
   End Select
   
   ' Transfer
   picSmallText.Picture = LoadPicture
   SetDIBits picSmallText.hdc, picSmallText.Image, _
       0, ImageHeight(ImageNum), picSmallDATA(0, 0), BMIH, 0
   picSmallText.Picture = picSmallText.Image
   
   Erase Temp(), picSmallDATA()
   
   ' Refresh image
   picSmallCopy.Picture = LoadPicture
   BitBlt picSmallCopy.hdc, 0, 0, picSmallCopy.ScaleWidth, picSmallCopy.ScaleHeight, _
    Form1.picSmall(ImageNum).hdc, 0, 0, vbSrcCopy
   picSmallCopy.Picture = picSmallCopy.Image
   
   picSmallCopy.CurrentX = picSmallText.CurrentX
   picSmallCopy.CurrentY = picSmallText.CurrentY
   ShowText picSmallCopy, TextLine$
End Sub

Private Sub cmdSelectFont_Click()
Dim a$
Dim cc As CFDialog
   Set cc = New CFDialog
   CurFont.Name = TFontname$
   If cc.VBChooseFont(CurFont, , Me.hwnd) Then
      With Text1
         .FontName = CurFont.Name
         .FontSize = CurFont.Size
         .FontItalic = CurFont.Italic
         .FontBold = CurFont.Bold
         '.ForeColor = TextColor
      End With
      With picSmallText
         .FontName = CurFont.Name
         .FontSize = CurFont.Size
         .FontItalic = CurFont.Italic
         .FontBold = CurFont.Bold
         .ForeColor = LColor
         '.FontStrikethru = False
         '.FontUnderline = False
      End With
      
      With picSmallCopy
         .FontName = CurFont.Name
         .FontSize = CurFont.Size
         .FontItalic = CurFont.Italic
         .FontBold = CurFont.Bold
         .ForeColor = LColor
         '.FontStrikethru = False
         '.FontUnderline = False
      End With
      
      ' Save selection
      TFontname$ = CurFont.Name
      TFontsize = CurFont.Size
      TFontBold = CurFont.Bold
      TFontItalic = CurFont.Italic
   End If
   
   Set cc = Nothing
   
   a$ = "Text - " & TFontname$ & " Size " & Str$(TFontsize)
   If TFontBold Then a$ = a$ & " Bold"
   If TFontItalic Then a$ = a$ & " Italic"
   If Not TFontBold And Not TFontItalic Then a$ = a$ & " Regular"
   Caption = a$
   Text1_Change
   Text1.SetFocus
End Sub


Private Sub picLColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
   Cul = picLColor.Point(X, Y)
   If Cul >= 0 Then
      If Button = vbLeftButton Then LColor = Cul
      Form1.LabColor(0).BackColor = LColor
      Text1.ForeColor = LColor
      picSmallCopy.ForeColor = vbBlack 'LColor
      
      If LColor = vbWhite Then
         Text1.BackColor = 0
      End If
      
      Text1_Change
      Text1.SetFocus
   End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
' Comment out if L/R arrow keys wanted for editting instead
   Select Case KeyCode
   ' Arrows & Keypad
   Case 37, 100: cmdLRUD_Click 0: KeyCode = 0  ' LeftA, 0
   Case 38, 104: cmdLRUD_Click 2: KeyCode = 0  ' UpA, 2
   Case 39, 102: cmdLRUD_Click 1: KeyCode = 0  ' RightA, 1
   Case 40, 98:  cmdLRUD_Click 3: KeyCode = 0  ' DownA, 3
   End Select
End Sub

Public Sub Text1_Change()
Dim ix As Long, iy As Long
   TextLine$ = Text1.Text
   picSmallText.Picture = LoadPicture
   picSmallText.Cls  ' Sets CurrentX.Y = 0
   picSmallText.BackColor = vbWhite
   picSmallText.ForeColor = vbBlack 'LColor
   Text1.BackColor = vbWhite
   If LColor = vbWhite Then
      Text1.BackColor = vbBlack
   End If
   
   ix = picSmallText.CurrentX
   iy = picSmallText.CurrentY
   ShowText picSmallText, TextLine$
   picSmallText.CurrentX = ix
   picSmallText.CurrentY = iy
   
   ' Refresh image from Form1.picSmall()
   picSmallCopy.Picture = LoadPicture
   picSmallCopy.ForeColor = LColor
   picSmallCopy.Cls   ' Sets CurrentX.Y = 0
   BitBlt picSmallCopy.hdc, 0, 0, picSmallCopy.ScaleWidth, picSmallCopy.ScaleHeight, _
    Form1.picSmall(ImageNum).hdc, 0, 0, vbSrcCopy
   picSmallCopy.Picture = picSmallCopy.Image
   
   picSmallCopy.CurrentX = ix
   picSmallCopy.CurrentY = iy
   ShowText picSmallCopy, TextLine$
   picSmallCopy.CurrentX = picSmallText.CurrentX
   picSmallCopy.CurrentY = picSmallText.CurrentY
End Sub

Private Sub Form_Unload(Cancel As Integer)
   aText = False
End Sub

'Private Sub Image1_Click()
'   If LColor <> vbWhite Then
'      FindMaxMins picSmallText, vbWhite
'   Else
'      FindMaxMins picSmallText, vbBlack
'   End If
'   Label2 = Str$(minx) & " ," & Str$(miny) & " :" & Str$(maxx) & " ," & Str$(maxy)
'
'End Sub

