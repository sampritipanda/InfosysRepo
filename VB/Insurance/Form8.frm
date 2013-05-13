VERSION 5.00
Begin VB.Form admn2 
   Caption         =   "Form8"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13905
   LinkTopic       =   "Form8"
   ScaleHeight     =   7590
   ScaleWidth      =   13905
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2880
      TabIndex        =   12
      Top             =   1200
      Width           =   5775
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   2040
      Width           =   5775
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   2880
      TabIndex        =   10
      Top             =   2760
      Width           =   5775
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   2880
      TabIndex        =   9
      Top             =   3840
      Width           =   5775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "BACK"
      Height          =   735
      Left            =   1800
      TabIndex        =   8
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CLEAR"
      Height          =   735
      Left            =   4800
      TabIndex        =   7
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SUBMIT"
      Height          =   735
      Left            =   7320
      TabIndex        =   6
      Top             =   6480
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   855
      Left            =   3000
      TabIndex        =   5
      Top             =   5040
      Width           =   5535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9960
      TabIndex        =   3
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE AGENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9840
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CREATE AGENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9840
      TabIndex        =   1
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VIEW AGENT INFO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9840
      TabIndex        =   0
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6135
      Left            =   9240
      TabIndex        =   4
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "ADD NEW AGENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   18
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "AGENT NAME"
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "AGENT ID"
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "AGENT ADDRESS"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "AGENT PHONE NO"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "PASSWORD"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   5160
      Width           =   2175
   End
End
Attribute VB_Name = "admn2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
aginfo.Show
End Sub

Private Sub Command2_Click()
addagnt.Show
End Sub

Private Sub Command3_Click()
delet.Show
End Sub

Private Sub Command4_Click()
main1.Show
Unload Me
End Sub

