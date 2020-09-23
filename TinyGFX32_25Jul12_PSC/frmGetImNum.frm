VERSION 5.00
Begin VB.Form frmGetImNum 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Open a recent file"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3915
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImNum 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   2520
      TabIndex        =   5
      Top             =   1365
      Width           =   480
   End
   Begin VB.CommandButton cmdImNum 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   1965
      TabIndex        =   4
      Top             =   1365
      Width           =   480
   End
   Begin VB.CommandButton cmdImNum 
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
      Height          =   345
      Index           =   0
      Left            =   1380
      TabIndex        =   2
      Top             =   1365
      Width           =   480
   End
   Begin VB.Label LabAlphaRestrict 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Alpha Restricted"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   90
      TabIndex        =   6
      Top             =   450
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Image number: "
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   1365
      Width           =   1110
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select an image number to fill                                (NB will fill from image 1 if multi-icon)."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   60
      TabIndex        =   1
      Top             =   705
      Width           =   4395
   End
   Begin VB.Label LabFileName 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "File name: "
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   195
      Width           =   765
   End
End
Attribute VB_Name = "frmGetImNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Accept As Boolean
' Public FileSpec$,TempImageNum

Private Sub cmdImNum_Click(Index As Integer)
   TempImageNum = Index
   Accept = True
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Move 600, 600
   LabFileName = "File name = " & GetFileName(FileSpec$)
   Accept = False
   If aAlphaRestricted Then
      LabAlphaRestrict.Visible = True
   Else
      LabAlphaRestrict.Visible = False
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Accept = False Then TempImageNum = -1
End Sub
