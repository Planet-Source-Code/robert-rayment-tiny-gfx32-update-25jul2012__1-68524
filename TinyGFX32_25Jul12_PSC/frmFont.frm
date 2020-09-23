VERSION 5.00
Begin VB.Form frmFont 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Courier New"
   ClientHeight    =   7455
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   6030
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFont.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4710
      TabIndex        =   7
      Top             =   150
      Width           =   990
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Font Fixed"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   408
      Width           =   936
   End
   Begin VB.PictureBox PICF 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1665
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   4
      Top             =   132
      Width           =   570
   End
   Begin VB.CommandButton cmdFont 
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   2400
      TabIndex        =   3
      Top             =   36
      Width           =   936
   End
   Begin VB.PictureBox PIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   240
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   363
      TabIndex        =   0
      Top             =   1092
      Width           =   5505
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Click on Table to transfer characters to Text form then Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   255
      TabIndex        =   6
      Top             =   900
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   405
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   30
      Width           =   1275
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Font Table by Robert Rayment

Option Explicit

Private prevX As Long, prevY As Long
Private ixstep As Long, iystep As Long
Private Color As Long

'Public TFontname As String
'Public TFontsize As Long
'Public TFontBold As Boolean
'Public TFontItalic As Boolean

Dim CurFont As StdFont


Private Sub Form_Load()
   Set CurFont = New StdFont
   ' Can change colors here
   Me.BackColor = &HFFC0C0
   Color = Me.BackColor Xor vbWhite
   Label2.BackColor = PIC.BackColor
   
   DrawGridTable
End Sub

Private Sub cmdFont_Click(Index As Integer)
Dim a$
Dim Flags As Long
Dim cc As CFDialog
   
   Set cc = New CFDialog
   CurFont.Name = TFontname$
   
   If Index = 0 Then
      Flags = &H1
   Else   ' Fixed font
      Flags = &H1 Or &H4000
   End If
   
   If cc.VBChooseFont(CurFont, , Me.hwnd, , , , Flags) Then
      With frmText.picSmallText
         .FontName = CurFont.Name
         .FontSize = CurFont.Size
         .FontItalic = CurFont.Italic
         .FontBold = CurFont.Bold
         .ForeColor = LColor
         '.FontStrikethru = False
         '.FontUnderline = False
      End With

      With frmText.Text1
         .FontName = CurFont.Name
         .FontSize = CurFont.Size
         .FontItalic = CurFont.Italic
         .FontBold = CurFont.Bold
         '.ForeColor = TextColor
      End With
      
      TFontname$ = CurFont.Name
      TFontsize = CurFont.Size
      TFontBold = CurFont.Bold
      TFontItalic = CurFont.Italic
      
      Caption = "Text - " & TFontname$ & " Size " & Str$(TFontsize)
      a$ = "Text - " & TFontname$ & " Size " & Str$(TFontsize)
      If TFontBold Then a$ = a$ & " Bold"
      If TFontItalic Then a$ = a$ & " Italic"
      If Not TFontBold And Not TFontItalic Then a$ = a$ & " Regular"
      
      frmText.Caption = a$
      frmText.Text1_Change

      DrawGridTable
   End If
End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim kx As Long, ky As Long
Dim num As Long
Dim W As Long, H As Long

   ' Clear prev bars
   PIC.Line (0, prevY - 7)-(prevX, prevY + 6), Color, BF     ' Horz bar
   PIC.Line (prevX, prevY - 7)-(prevX - 34, 0), Color, BF    ' Vert bar
   
   kx = (X - 22) \ ixstep - 10
   ky = (Y - 14) \ iystep + 1
   
   num = 10 * ky + kx
   Label1(0).Caption = "Dec" & Str$(num)
   Label1(1).Caption = "Hex " & Hex$(num)
   If num >= 0 And num <= 255 Then
      If num = 9 Or num = 10 Or num = 13 Then num = 0 ' Else affects printing
      PICF.Cls
      PICF.CurrentX = 6
      PICF.CurrentY = 3
      PICF.FontName = TFontname
      W = PICF.TextWidth(Chr$(num))
      If TFontItalic Then W = W + 4
      H = PICF.TextHeight(Chr$(num))
      PICF.Print Chr$(num);
      ' Draw approx outline rectangle
      ' NB doesn't work accurately for all Fonts
      PICF.Line (5, 2)-(5 + W, 3 + H), vbRed, B
   End If
   
   X = ixstep * ((X + 12) \ ixstep) + 22
   Y = iystep * ((Y + 1) \ iystep) + 7
   ' Draw bars
   PIC.Line (0, Y - 7)-(X, Y + 6), Color, BF    ' Horz bar
   PIC.Line (X, Y - 7)-(X - 34, 0), Color, BF   ' Vert bar
   prevX = X: prevY = Y
   
End Sub

Private Sub PIC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim kx As Long, ky As Long
Dim num As Long
Dim a$
   kx = (X - 22) \ ixstep - 10
   ky = (Y - 14) \ iystep + 1
   num = 10 * ky + kx
   If num > 255 Then Exit Sub
   
   a$ = Chr$(num)
   frmText.Text1 = frmText.Text1 & a$
' EG
' FontName  =  "Courier New"
' Selected code  =  Dec  205, Hex CD
End Sub

Private Sub DrawGridTable()
Dim ix As Long
Dim iy As Long
Dim i As Long, j As Long
Dim num As Long
Dim ilo As Long
Dim iup As Long
Dim H$

   PIC.DrawMode = 13
   PIC.Cls
   PICF.Cls
   With PIC
      .ScaleWidth = 367
      .ScaleHeight = 409
      .FontName = TFontname
      .FontSize = 8
   End With
   Caption = PIC.FontName & "  Size " & PIC.FontSize
   PICF.FontName = TFontname
   PICF.FontSize = 14
   PICF.FontBold = TFontBold
   PICF.FontItalic = TFontItalic
   '-----------------------------------------
   PIC.FontName = "Courier New"
   PIC.FontSize = 8
   PIC.FontBold = False
   ix = 22
   ixstep = 34
   For i = 0 To 9
      PIC.CurrentX = ix
      PIC.CurrentY = 0
      PIC.Line (ix, 0)-(ix, PIC.Height), 0
      PIC.CurrentX = ix - 4
      PIC.CurrentY = 0
      PIC.Print " "; Str$(i);
      ix = ix + ixstep
   Next i
   PIC.Line (ix, 0)-(ix, PIC.Height), 0
   
   iy = 12
   iystep = 15
   
   For i = 0 To 25
      PIC.CurrentX = 0
      PIC.CurrentY = iy + 2
      If i < 10 Then PIC.Print " ";
      PIC.Print Str$(i)
      PIC.CurrentX = 0
      PIC.CurrentY = iy
      PIC.Line (0, iy + 2)-(PIC.Width, iy + 2), 0
      iy = iy + iystep
   Next i
   PIC.Line (0, iy + 2)-(PIC.Width, iy + 2), 0
   '-----------------------------------------
   num = 0
   iy = 14
   ilo = 0: iup = 9
   For j = 0 To 25
      PIC.CurrentY = iy
      For i = ilo To iup
         PIC.CurrentX = 25 + (i - ilo) * (ixstep)
         H$ = Hex$(num)
         If Len(H$) < 2 Then H$ = "0" & H$
         H$ = H$ & " "
         If num = 9 Or num = 10 Or num = 13 Then   ' These affect printing, exclude
            PIC.FontName = "Courier New"
            PIC.FontSize = 8
            PIC.FontBold = False
            PIC.FontItalic = False
            PIC.Print H$;
            PIC.FontName = TFontname$
            'PIC.FontSize = TFontSize
            'PIC.FontBold = TFontBold
            PIC.Print Chr$(0);
         Else
            PIC.FontName = "Courier New"
            PIC.FontSize = 8
            PIC.FontBold = False
            PIC.FontItalic = False
            PIC.Print H$;  ' Hex
            ' Actual characters
            PIC.FontName = TFontname$
            'PIC.FontSize = FntSize ' Would need a big Picbox for this
            PIC.FontBold = TFontBold
            PIC.FontItalic = TFontItalic
            PIC.Print Chr$(num);
         End If
         num = num + 1
         If num = 256 Then Exit For
      Next i
      If num = 256 Then Exit For
      iy = iy + iystep
      ilo = ilo + 10
      iup = iup + 10
   Next j
   '-----------------------------------------
   PIC.FontName = "Courier New"
   PIC.FontSize = 8
   PIC.FontBold = False
   PIC.CurrentX = 0
   PIC.CurrentY = 0
   PIC.Print "DHA";
   '-----------------------------------------
   PIC.DrawMode = 7
   PIC.Line (0, 0)-(PIC.Width, 14), Color, BF
   PIC.Line (0, 0)-(21, PIC.Height), Color, BF
   ' Clear prev bars
   PIC.Line (0, prevY - 7)-(prevX, prevY + 6), Color, BF    ' Horz bar
   PIC.Line (prevX, prevY - 7)-(prevX - 34, 0), Color, BF   ' Vert bar
   
   '-----------------------------------------
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

