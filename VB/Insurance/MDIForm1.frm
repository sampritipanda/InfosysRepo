VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.MDIForm main1 
   BackColor       =   &H8000000A&
   Caption         =   "MAIN WINDOW"
   ClientHeight    =   8670
   ClientLeft      =   5655
   ClientTop       =   4980
   ClientWidth     =   15930
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      DrawMode        =   6  'Mask Pen Not
      Height          =   60060
      Left            =   0
      Picture         =   "MDIForm1.frx":18CD06
      ScaleHeight     =   60000
      ScaleMode       =   0  'User
      ScaleWidth      =   25595.46
      TabIndex        =   0
      Top             =   0
      Width           =   20370
      Begin VB.PictureBox Picture3 
         Height          =   3255
         Left            =   8280
         Picture         =   "MDIForm1.frx":30E4F2
         ScaleHeight     =   3195
         ScaleWidth      =   4875
         TabIndex        =   9
         Top             =   6240
         Width           =   4935
      End
      Begin VB.PictureBox Picture2 
         Height          =   3255
         Left            =   8280
         Picture         =   "MDIForm1.frx":341CB4
         ScaleHeight     =   3195
         ScaleWidth      =   4755
         TabIndex        =   8
         Top             =   2520
         Width           =   4815
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0FF&
         Caption         =   "END"
         Height          =   615
         Left            =   15840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7200
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "AGENT LOGIN"
         Height          =   615
         Left            =   15840
         MaskColor       =   &H00FFFF00&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "ADMINISTRATOR LOGIN"
         Height          =   615
         Left            =   15840
         MaskColor       =   &H00C0C000&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3240
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "CUSTOMER LOGIN"
         Height          =   615
         Left            =   15840
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5880
         Width           =   1815
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   4815
         Left            =   2400
         TabIndex        =   7
         Top             =   3360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8493
         _Version        =   393217
         BackColor       =   16761087
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         DisableNoScroll =   -1  'True
         FileName        =   "C:\Users\facti\Desktop\New Text Document (6).txt"
         TextRTF         =   $"MDIForm1.frx":375476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         DrawMode        =   6  'Mask Pen Not
         Height          =   6975
         Index           =   1
         Left            =   2160
         Shape           =   4  'Rounded Rectangle
         Top             =   2280
         Width           =   4770
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   16920
         Picture         =   "MDIForm1.frx":3754F2
         Top             =   480
         Width           =   3060
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ARMY WARDS  WELFARE ASSOCIATION INITIATIVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11520
         TabIndex        =   5
         Top             =   1680
         Width           =   5415
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   495
         Left            =   11040
         Shape           =   4  'Rounded Rectangle
         Top             =   1560
         Width           =   5895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ARMY WELFARE INSURANCE AGENCY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         TabIndex        =   4
         Top             =   600
         Width           =   13815
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   4  'Dash-Dot
         BorderWidth     =   3
         DrawMode        =   8  'Xor Pen
         Height          =   975
         Left            =   2880
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   14055
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   4  'Dash-Dot
         BorderWidth     =   3
         DrawMode        =   6  'Mask Pen Not
         Height          =   8175
         Left            =   7560
         Shape           =   4  'Rounded Rectangle
         Top             =   2280
         Width           =   6255
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         DrawMode        =   6  'Mask Pen Not
         Height          =   6975
         Index           =   0
         Left            =   14280
         Shape           =   4  'Rounded Rectangle
         Top             =   2280
         Width           =   4770
      End
   End
   Begin VB.Menu home 
      Caption         =   "HOME"
   End
   Begin VB.Menu pro 
      Caption         =   "PRODUCTS"
      Begin VB.Menu li 
         Caption         =   "LIFE INSURANCE"
         Begin VB.Menu wlp 
            Caption         =   "WHOLE LIFE POLICY"
         End
         Begin VB.Menu tlp 
            Caption         =   "TERM LIFE POLICY"
         End
         Begin VB.Menu ep 
            Caption         =   "ENDOWMENT POLICY"
         End
         Begin VB.Menu pp 
            Caption         =   "PENSION PLANS"
         End
      End
      Begin VB.Menu gi 
         Caption         =   "GENERAL INSURANCE"
         Begin VB.Menu hi 
            Caption         =   "HOME INSURANCE"
         End
         Begin VB.Menu ai 
            Caption         =   "AUTO INSURANCE"
         End
         Begin VB.Menu fi 
            Caption         =   "FIRE INSURANCE"
         End
      End
   End
   Begin VB.Menu cnt 
      Caption         =   "CONTACT US"
   End
   Begin VB.Menu ab 
      Caption         =   "ABOUT US"
   End
End
Attribute VB_Name = "main1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ab_Click()
RichTextBox1.FileName = "C:\Documents and Settings\Administrator\Desktop\123\proj\abt.txt"
End Sub
Private Sub alog_Click()
frmlogin2.Show
End Sub

Private Sub car_Click()
Form4.Show
End Sub

Private Sub ai_Click()
RichTextBox1.FileName = "C:\Documents and Settings\Administrator\Desktop\123\proj\auto.txt"
End Sub

Private Sub cnt_Click()
RichTextBox1.FileName = "C:\Documents and Settings\Administrator\Desktop\123\proj\New Text Document (2).txt"

End Sub

Private Sub Command1_Click()
frmLogin1.Show
End Sub

Private Sub Command2_Click()
frmLogin.Show
End Sub

Private Sub Command3_Click()
frmlogin2.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub ep_Click()
RichTextBox1.FileName = "C:\Documents and Settings\Administrator\Desktop\123\proj\proj\end.txt"
End Sub

Private Sub fi_Click()
RichTextBox1.FileName = "C:\Documents and Settings\Administrator\Desktop\123\proj\fire.txt"
End Sub

Private Sub hi_Click()
RichTextBox1.FileName = "C:\Documents and Settings\Administrator\Desktop\123\proj\house.txt"
End Sub

Private Sub home_Click()
RichTextBox1.FileName = "C:\Documents and Settings\Administrator\Desktop\123\proj\home.txt"

End Sub

Private Sub pp_Click()
RichTextBox1.FileName = "C:\Documents and Settings\Administrator\Desktop\123\proj\pp.txt"
End Sub

Private Sub tlp_Click()
RichTextBox1.FileName = "C:\Documents and Settings\Administrator\Desktop\123\proj\term.txt"
End Sub

Private Sub wlp_Click()
RichTextBox1.FileName = "C:\Documents and Settings\Administrator\Desktop\123\proj\whole.txt"
End Sub
