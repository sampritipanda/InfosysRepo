VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form cinfo 
   AutoRedraw      =   -1  'True
   Caption         =   "CUSTOMER PAGE"
   ClientHeight    =   9135
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15375
   LinkTopic       =   "Form7"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Index           =   0
      Left            =   600
      ScaleHeight     =   6315
      ScaleWidth      =   14235
      TabIndex        =   2
      Top             =   1440
      Width           =   14295
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NAME"
         Height          =   375
         Left            =   480
         TabIndex        =   90
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AGE"
         Height          =   375
         Left            =   480
         TabIndex        =   89
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MARITAL STATUS"
         Height          =   495
         Left            =   480
         TabIndex        =   88
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ID NO."
         Height          =   375
         Left            =   480
         TabIndex        =   87
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SEX"
         Height          =   375
         Left            =   480
         TabIndex        =   86
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   8520
         TabIndex        =   28
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   1680
         TabIndex        =   31
         Top             =   2160
         Width           =   1245
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   1680
         TabIndex        =   29
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   8520
         TabIndex        =   27
         Top             =   3600
         Width           =   1680
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   8520
         TabIndex        =   26
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   8520
         TabIndex        =   25
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   1680
         TabIndex        =   24
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   23
         Top             =   3720
         Width           =   1125
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ANNUAL INCOME"
         Height          =   375
         Left            =   6360
         TabIndex        =   15
         Top             =   3600
         Width           =   1875
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CONTACT NUMBER"
         Height          =   375
         Left            =   6360
         TabIndex        =   14
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ADDRESS"
         Height          =   375
         Left            =   6360
         TabIndex        =   13
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DATE OF BIRTH"
         Height          =   375
         Left            =   6360
         TabIndex        =   12
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "1"
         Height          =   495
         Index           =   0
         Left            =   11040
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   1635
         TabIndex        =   30
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   15750
         Left            =   -120
         Picture         =   "Form7.frx":0000
         Top             =   0
         Width           =   25200
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "LOGOUT"
      Height          =   735
      Left            =   17160
      TabIndex        =   74
      Top             =   9360
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CHANGE PASSWORD"
      Height          =   735
      Left            =   13200
      TabIndex        =   73
      Top             =   9360
      Width           =   2535
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6975
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   12303
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "PERSONAL INFO"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "INSURANCES"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "AGENT INFO"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CLAIMANT"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "PREMIUM & MATURITY"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   5055
      Index           =   1
      Left            =   960
      ScaleHeight     =   4995
      ScaleWidth      =   13275
      TabIndex        =   4
      Top             =   1440
      Width           =   13335
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000013&
         Caption         =   "DETAILS"
         Height          =   495
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000013&
         Height          =   315
         ItemData        =   "Form7.frx":34DC0
         Left            =   2160
         List            =   "Form7.frx":34DC2
         TabIndex        =   16
         Text            =   "SELECT POLICY TYPE"
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   4
         Left            =   10440
         TabIndex        =   75
         Top             =   4320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   36
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   3
         Left            =   8280
         TabIndex        =   35
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   32
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INSURED ON"
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NAME"
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "POLICY DURATION"
         Height          =   375
         Left            =   6240
         TabIndex        =   20
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AMOUNT"
         Height          =   375
         Left            =   6240
         TabIndex        =   19
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PREMIUM"
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "POLICY TYPE"
         Height          =   495
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "2"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   33
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   2
         Left            =   8280
         TabIndex        =   34
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Image Image3 
         Height          =   15750
         Left            =   -480
         Picture         =   "Form7.frx":34DC4
         Top             =   240
         Width           =   25200
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Index           =   4
      Left            =   960
      ScaleHeight     =   4875
      ScaleWidth      =   12795
      TabIndex        =   10
      Top             =   1440
      Width           =   12855
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000013&
         Caption         =   "DETAILS"
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H80000013&
         Height          =   315
         Left            =   2880
         TabIndex        =   63
         Text            =   "select your policy"
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ANNUAL"
         Height          =   495
         Index           =   2
         Left            =   3000
         TabIndex        =   68
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   6
         Left            =   9480
         TabIndex        =   72
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   5
         Left            =   9480
         TabIndex        =   71
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   4
         Left            =   9480
         TabIndex        =   70
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   3
         Left            =   9480
         TabIndex        =   69
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   1
         Left            =   3000
         TabIndex        =   67
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "POLICY"
         Height          =   375
         Left            =   960
         TabIndex        =   64
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AMOUNT TO BE  RECEIVED ON MATURITY"
         Height          =   615
         Left            =   7080
         TabIndex        =   62
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MATURITY DATE"
         Height          =   495
         Left            =   7080
         TabIndex        =   61
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NEXT PREMIUM AMOUNT"
         Height          =   495
         Left            =   7080
         TabIndex        =   60
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NEXT PREMIUM DUE"
         Height          =   495
         Left            =   7080
         TabIndex        =   59
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PREMIUM PERIOD"
         Height          =   495
         Left            =   960
         TabIndex        =   58
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INSURANCE AMOUNT"
         Height          =   495
         Left            =   960
         TabIndex        =   57
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INSURANCE DATE"
         Height          =   495
         Left            =   960
         TabIndex        =   56
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "5"
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   0
         Left            =   3000
         TabIndex        =   66
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Image Image5 
         Height          =   15750
         Left            =   120
         Picture         =   "Form7.frx":69B84
         Top             =   0
         Width           =   25200
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Index           =   3
      Left            =   600
      ScaleHeight     =   4395
      ScaleWidth      =   13035
      TabIndex        =   8
      Top             =   1560
      Width           =   13095
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000013&
         Caption         =   "DETAILS"
         Height          =   615
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   240
         Width           =   3135
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2160
         TabIndex        =   54
         Text            =   "SELECT POLICY"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NAME"
         Height          =   615
         Left            =   360
         TabIndex        =   85
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ADDRESS"
         Height          =   615
         Left            =   360
         TabIndex        =   84
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DATE OF BIRTH"
         Height          =   615
         Left            =   360
         TabIndex        =   83
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AGE"
         Height          =   615
         Left            =   360
         TabIndex        =   82
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   0
         Left            =   2520
         TabIndex        =   81
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   1
         Left            =   2520
         TabIndex        =   80
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   5
         Left            =   2520
         TabIndex        =   79
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   7
         Left            =   2520
         TabIndex        =   78
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   11160
         TabIndex        =   77
         Top             =   4080
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Height          =   615
         Index           =   2
         Left            =   0
         TabIndex        =   76
         Top             =   0
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   8
         Left            =   8760
         TabIndex        =   53
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   6
         Left            =   8760
         TabIndex        =   52
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   4
         Left            =   8760
         TabIndex        =   51
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CONTACT NUMBER"
         Height          =   615
         Left            =   6240
         TabIndex        =   50
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SEX"
         Height          =   615
         Left            =   6240
         TabIndex        =   49
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MARITAL STATUS"
         Height          =   615
         Left            =   6240
         TabIndex        =   48
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "4"
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Image Image2 
         Height          =   15750
         Left            =   -360
         Picture         =   "Form7.frx":9E944
         Top             =   0
         Width           =   25200
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Index           =   2
      Left            =   1800
      ScaleHeight     =   4395
      ScaleWidth      =   12075
      TabIndex        =   6
      Top             =   2040
      Width           =   12135
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000013&
         Caption         =   "DETAILS"
         Height          =   615
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2640
         TabIndex        =   42
         Text            =   "SELECT AGENT ID"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   3
         Left            =   7920
         TabIndex        =   47
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Index           =   2
         Left            =   2400
         TabIndex        =   46
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   1
         Left            =   2400
         TabIndex        =   45
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   0
         Left            =   2400
         TabIndex        =   44
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AGENT ID"
         Height          =   495
         Left            =   6120
         TabIndex        =   41
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PHONE NUMBER"
         Height          =   495
         Left            =   480
         TabIndex        =   40
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NAME"
         Height          =   615
         Left            =   480
         TabIndex        =   38
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "3"
         Height          =   495
         Index           =   2
         Left            =   3720
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ADDRESS"
         Height          =   615
         Left            =   480
         TabIndex        =   39
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Image Image4 
         Height          =   15750
         Left            =   0
         Picture         =   "Form7.frx":D3704
         Top             =   0
         Width           =   25200
      End
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER INFO."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   1560
      TabIndex        =   91
      Top             =   360
      Width           =   7095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   735
      Left            =   600
      Top             =   240
      Width           =   9735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1815
      Left            =   12600
      Top             =   8760
      Width           =   7215
   End
   Begin VB.Image Image6 
      Height          =   18000
      Left            =   -480
      Picture         =   "Form7.frx":1084C4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   28800
   End
   Begin VB.Label Label1 
      Caption         =   "CUSTOMER INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "cinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Combo1.Text = "SELECT POLICY TYPE" Then
    MsgBox "First enter Policy Type"
Else
rk.MoveFirst
r2.MoveFirst
Do While Not rk.EOF
    If Combo1.Text = rk.Fields(1) Then
        Do While Not r2.EOF
            If rk.Fields(0) = r2.Fields(4) And rk.Fields(2) = r2.Fields(6) Then
            For i = 0 To Label18.Count - 1
             If i <> 4 Then
                Label18(i).Caption = r2.Fields(i)
                End If
            Next
            Exit Do
            End If
           r2.MoveNext
        Loop
    End If
 rk.MoveNext
 Loop
End If

End Sub

Private Sub Command2_Click()
If Combo2.Text = "SELECT AGENT ID" Then
    MsgBox "ENTER ID FIRST"
Else
rk.MoveFirst
r3.MoveFirst
ag.MoveFirst
Do While Not rk.EOF
    If Combo2.Text = rk.Fields(1) Then
        Do While Not r3.EOF
            If rk.Fields(0) = r3.Fields(2) And rk.Fields(2) = r3.Fields(3) Then
                Do While Not ag.EOF
                    If r3.Fields(0) = ag.Fields(3) Then
                        For i = 0 To Label23.Count - 1
                            Label23(i).Caption = ag.Fields(i)
                        Next
                    
                    Exit Do
                    End If
                ag.MoveNext
                Loop
                Exit Do
            End If
            r3.MoveNext
            Loop
        Exit Do
    End If
    rk.MoveNext
    Loop
End If
End Sub

Private Sub Command3_Click()
If Combo3.Text = "SELECT POLICY" Then
    MsgBox "SELECT POLICY FIRST"
Else
rk.MoveFirst
r4.MoveFirst
Do While Not rk.EOF
    If Combo3.Text = rk.Fields(1) Then
        Do While Not r4.EOF
            If rk.Fields(0) = r4.Fields(2) And rk.Fields(2) = r4.Fields(3) Then
            For i = 0 To Label33.Count - 1
             If i <> 3 Then
                Label33(i).Caption = r4.Fields(i)
                End If
            Next
            Exit Do
            End If
           r4.MoveNext
        Loop
    End If
    rk.MoveNext
 Loop
End If
End Sub

Private Sub Command4_Click()
Dim dat, dat2 As Date
Dim amt As Currency
Dim X1 As Integer
If Combo4.Text = "select your policy" Then
    MsgBox "ENTER POLICY TYPE FIRST"
Else
rk.MoveFirst
r2.MoveFirst
amt = 0
Do While Not rk.EOF
    If Combo4.Text = rk.Fields(1) Then
        Do While Not r2.EOF
            If rk.Fields(0) = r2.Fields(4) And rk.Fields(2) = r2.Fields(6) Then
                Label42(0).Caption = r2.Fields(5)
                Label42(1).Caption = r2.Fields(2)
                dat = r2.Fields(5)
                For X1 = 1 To r2.Fields(3)
                    If dat < Now Then
                        dat = dat + 365
                    End If
                Next
                Label42(3).Caption = dat
                Label42(4).Caption = r2.Fields(1)
                dat2 = r2.Fields(5) + 365 * r2.Fields(3)
                Label42(5).Caption = dat2
                For i = 1 To r2.Fields(3)
                    amt = amt + r2.Fields(1) + 0.0075 * amt * (r2.Fields(3) - i)
                Next
                Label42(6).Caption = amt
            Exit Do
            End If
            r2.MoveNext
        Loop
     Exit Do
     End If
     rk.MoveNext
Loop
End If
End Sub

Private Sub Command5_Click()
pass.Show
End Sub

Private Sub Command6_Click()
main1.Show
Unload Me

End Sub

Private Sub Form_Load()
 Dim i As Integer
On Error Resume Next
cmd1 = "Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False"

Set cn = New adodb.Connection
Set r2 = New adodb.Recordset
Set rk = New adodb.Recordset
Set ag = New adodb.Recordset
Set r3 = New adodb.Recordset
Set r4 = New adodb.Recordset
With cn
 .ConnectionString = cmd1
.CursorLocation = adUseClient
 .Open
 End With
    
r2.Open "SELECT * FROM ins", cn, 2, 3
rk.Open "SELECT * FROM pol", cn, 2, 3
ag.Open "select * from agent", cn, 2, 3
r3.Open "select * from sales", cn, 2, 3
r4.Open "select * from claimant", cn, 2, 3
End Sub

Private Sub TabStrip1_Click()
For intloopindex = 0 To Picture1.Count - 1
 With Picture1(intloopindex)
.Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
Picture1(TabStrip1.SelectedItem.Index - 1).ZOrder 0
End With
Next intloopindex
End Sub

