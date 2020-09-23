VERSION 5.00
Begin VB.Form frmPalette 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " User bmp Palette"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2280
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   152
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picIN 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   2985
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   5
      Top             =   255
      Width           =   1110
   End
   Begin VB.PictureBox picPalette 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   180
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   126
      TabIndex        =   0
      Top             =   135
      Width           =   1920
   End
   Begin VB.Label LabRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1410
      TabIndex        =   4
      Top             =   2610
      Width           =   390
   End
   Begin VB.Label LabRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   915
      TabIndex        =   3
      Top             =   2610
      Width           =   390
   End
   Begin VB.Label LabRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   405
      TabIndex        =   2
      Top             =   2610
      Width           =   390
   End
   Begin VB.Label LabCul 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   2190
      Width           =   795
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmPalette.frm

Option Explicit

Private CommonDialog1 As cOSDialog


Private Sub Form_Load()
Dim Title$, Filt$, InDir$
Dim BMPSpec$

   picPalette.Width = 128
   picPalette.Height = 128

   Me.Move 600, 600
   
   Title$ = "Load palette bmp file <= 64 x 64"
   Filt$ = "Pics bmp|*.bmp"
   BMPSpec$ = ""
   InDir$ = CPath$ 'AppPathSpec$
   Set CommonDialog1 = New cOSDialog

   CommonDialog1.ShowOpen BMPSpec$, Title$, Filt$, InDir$, "", Me.hwnd
   Set CommonDialog1 = Nothing
   
   If Len(BMPSpec$) = 0 Then
      Exit Sub
   End If
   
   picIN.Picture = LoadPicture(BMPSpec$)
   
   TransferSmallToLarge picIN, picPalette
   
   picIN.Picture = LoadPicture
   picIN.Width = 4
   picIN.Height = 4
   
End Sub

Private Sub picPalette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
Dim pr As Byte, pg As Byte, pb As Byte

   Cul = picPalette.Point(X, Y)
   If Cul < 0 Then Cul = 0
   LabCul.BackColor = Cul
   LngToRGB Cul, pr, pg, pb
   LabRGB(0) = pr
   LabRGB(1) = pg
   LabRGB(2) = pb
   If Button = vbLeftButton Then
      Form1.LabColor(0).BackColor = Cul
      LColor = Cul
   ElseIf Button = vbRightButton Then
      Form1.LabColor(1).BackColor = Cul
      RColor = Cul
   End If
End Sub

Private Sub picPalette_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cul As Long
Dim pr As Byte, pg As Byte, pb As Byte

   Cul = picPalette.Point(X, Y)
   If Cul < 0 Then Cul = 0
   LabCul.BackColor = Cul
   LngToRGB Cul, pr, pg, pb
   LabRGB(0) = pr
   LabRGB(1) = pg
   LabRGB(2) = pb
   
End Sub
