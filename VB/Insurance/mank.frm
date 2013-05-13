VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form load 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "loading..."
   ClientHeight    =   5865
   ClientLeft      =   210
   ClientTop       =   1305
   ClientWidth     =   6300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   2  'Dot
   FillColor       =   &H00800000&
   ForeColor       =   &H80000011&
   Icon            =   "mank.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "mank.frx":000C
   ScaleHeight     =   5865
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   3480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      MaskColor       =   &H00800080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2760
      Top             =   4920
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MANOHAR MISHRA   N KARTHIK                MAYANK SINGH"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "COPYRIGHT AWWA INC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AWWA INSURANCE AGENCY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   1200
      Width           =   5175
   End
   Begin VB.Label loading 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   3000
      Width           =   2655
   End
End
Attribute VB_Name = "load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
loading.Caption = "loading..."
ProgressBar1.Value = 0
Timer1.Enabled = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    
End Sub

Private Sub Form_Load()
    Timer1.Enabled = False
    ProgressBar1.Value = 0
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 10
If ProgressBar1.Value >= 100 Then Timer1.Enabled = False
If Timer1.Enabled = False Then
main1.Show
Unload Me
End If
End Sub
