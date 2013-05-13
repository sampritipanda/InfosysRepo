VERSION 5.00
Begin VB.Form cboDept 
   Caption         =   "Employee Details"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4080
      TabIndex        =   13
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   615
      Left            =   1680
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ListBox lstSkillSet 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      ItemData        =   "frmAsgmt3.frx":0000
      Left            =   2160
      List            =   "frmAsgmt3.frx":001C
      MultiSelect     =   1  'Simple
      TabIndex        =   11
      Top             =   3360
      Width           =   2775
   End
   Begin VB.ComboBox cboDept 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmAsgmt3.frx":005F
      Left            =   2160
      List            =   "frmAsgmt3.frx":0075
      TabIndex        =   9
      Text            =   "Department"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.OptionButton optFemale 
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3480
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.OptionButton optMale 
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   2160
      TabIndex        =   6
      Top             =   2280
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox txtENumber 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtEName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label lblSkillSet 
      AutoSize        =   -1  'True
      Caption         =   "Skill Set :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   10
      Top             =   3480
      Width           =   915
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "Department :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   8
      Top             =   2760
      Width           =   1230
   End
   Begin VB.Label lblGender 
      AutoSize        =   -1  'True
      Caption         =   "Gender :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   5
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label lblENumber 
      AutoSize        =   -1  'True
      Caption         =   "Employee # :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   1230
   End
   Begin VB.Label lblEName 
      AutoSize        =   -1  'True
      Caption         =   "Employee name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1665
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Employee Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   2100
   End
End
Attribute VB_Name = "cboDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub lstSkillSet_Click()

End Sub
