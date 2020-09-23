VERSION 5.00
Begin VB.Form frmHelp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5355
   ClientLeft      =   150
   ClientTop       =   0
   ClientWidth     =   8325
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   357
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Top             =   4785
      Width           =   1200
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   2580
      TabIndex        =   1
      Top             =   75
      Width           =   5520
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4545
      Left            =   180
      TabIndex        =   0
      Top             =   75
      Width           =   2340
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessageLong Lib "USER32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Const LB_FINDSTRINGEXACT = &H1A2

Private NumLines As Long

Private Sub Form_Load()
Dim fnum As Integer
Dim k As Long
Dim a$
   Caption = " Tiny GFX Help"
   
'   Public HelpSpec$ = AppPathSpec$ & "GFXHelp.txt"

   fnum = FreeFile
   Open HelpSpec$ For Input As #fnum
   Input #fnum, NumLines
   
   List1.AddItem " "
   For k = 1 To NumLines    ' Number of FVHelp Contents' items
      Line Input #fnum, a$
      List1.AddItem a$
   Next k
   k = 0
   Do Until EOF(1)
      Line Input #fnum, a$
      If k = 0 Then
         List2.AddItem "    "
         List2.AddItem "    TinyGFX by Robert Rayment"
         List2.AddItem "    _________________________"
         List2.AddItem "    "
      Else
         List2.AddItem a$
      End If
      k = k + 1
   Loop
   Close #fnum
End Sub

Private Sub List1_Click()
Dim a$
Dim k As Long
Dim res As Long
   'Select item
   k = List1.ListIndex
   a$ = List1.List(k) & Chr$(0)
   If Len(a$) <> 0 Then
      'Search List2 for Text$ & place at top
      res = SendMessageLong(List2.hwnd, LB_FINDSTRINGEXACT, -1&, ByVal a$)
      List2.ListIndex = res
      If List2.ListIndex > 0 Then
         List2.TopIndex = List2.ListIndex - 1
      End If
   End If
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub


