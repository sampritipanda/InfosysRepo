VERSION 5.00
Begin VB.Form addpol1 
   BackColor       =   &H80000009&
   Caption         =   "CREATE NEW POLICY"
   ClientHeight    =   11010
   ClientLeft      =   4005
   ClientTop       =   3330
   ClientWidth     =   20370
   DrawMode        =   6  'Mask Pen Not
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check4 
      Caption         =   "SINGLE"
      Height          =   375
      Left            =   11880
      TabIndex        =   70
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "MARRIED"
      Height          =   375
      Left            =   10560
      TabIndex        =   69
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "UNMARRIED"
      Height          =   375
      Left            =   11760
      TabIndex        =   68
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "MARRIED"
      Height          =   375
      Left            =   10560
      TabIndex        =   67
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000005&
      Caption         =   "ADD DATE"
      Height          =   375
      Left            =   11880
      TabIndex        =   59
      Top             =   9600
      Width           =   975
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      ItemData        =   "addpol.frx":0000
      Left            =   12120
      List            =   "addpol.frx":0010
      TabIndex        =   58
      Text            =   "yyyy"
      Top             =   9120
      Width           =   855
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      ItemData        =   "addpol.frx":002C
      Left            =   11280
      List            =   "addpol.frx":0039
      TabIndex        =   57
      Text            =   "mm"
      Top             =   9120
      Width           =   735
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      ItemData        =   "addpol.frx":004C
      Left            =   10560
      List            =   "addpol.frx":006E
      TabIndex        =   56
      Text            =   "dd"
      Top             =   9120
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3600
      TabIndex        =   55
      Text            =   "SELECT THE POLICY"
      Top             =   9120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10560
      TabIndex        =   54
      Top             =   9600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   16920
      TabIndex        =   53
      Top             =   9120
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   16920
      TabIndex        =   52
      Top             =   9600
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   3240
      TabIndex        =   43
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   10560
      MaxLength       =   10
      TabIndex        =   42
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   10560
      TabIndex        =   41
      Top             =   7560
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   3240
      TabIndex        =   40
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   13560
      TabIndex        =   39
      Top             =   6840
      Width           =   255
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   16800
      TabIndex        =   38
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      Height          =   405
      Left            =   16800
      TabIndex        =   37
      Top             =   7560
      Width           =   615
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "addpol.frx":0091
      Left            =   16800
      List            =   "addpol.frx":00B3
      TabIndex        =   36
      Text            =   "dd"
      Top             =   6240
      Width           =   495
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "addpol.frx":00D6
      Left            =   17520
      List            =   "addpol.frx":00FE
      TabIndex        =   35
      Text            =   "mm"
      Top             =   6240
      Width           =   495
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "addpol.frx":0142
      Left            =   18240
      List            =   "addpol.frx":0158
      TabIndex        =   34
      Text            =   "yyyy"
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000009&
      Caption         =   "ADD DOB"
      Height          =   375
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6960
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000009&
      Caption         =   "MALE"
      Height          =   255
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6840
      Width           =   855
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H80000009&
      Caption         =   "FEMALE"
      Height          =   255
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   4200
      TabIndex        =   19
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   10560
      TabIndex        =   18
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   7
      Left            =   13560
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   16560
      TabIndex        =   16
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   5
      Left            =   16560
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   3
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   14
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   2
      Left            =   10440
      TabIndex        =   13
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   1
      Left            =   4200
      TabIndex        =   12
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Index           =   0
      Left            =   4200
      TabIndex        =   11
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHECK ID"
      Height          =   495
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000009&
      Caption         =   "MALE"
      Height          =   255
      Index           =   0
      Left            =   10560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Index           =   4
      Left            =   13320
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000009&
      Caption         =   "FEMALE"
      Height          =   255
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "addpol.frx":0180
      Left            =   16560
      List            =   "addpol.frx":01A2
      TabIndex        =   6
      Text            =   "dd"
      Top             =   2640
      Width           =   615
   End
   Begin VB.ComboBox mm 
      Height          =   315
      ItemData        =   "addpol.frx":01C5
      Left            =   17280
      List            =   "addpol.frx":01ED
      TabIndex        =   5
      Text            =   "mm"
      Top             =   2640
      Width           =   615
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "addpol.frx":022D
      Left            =   18000
      List            =   "addpol.frx":0240
      TabIndex        =   4
      Text            =   "yyyy"
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Caption         =   "ADD DOB"
      Height          =   375
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000013&
      Caption         =   "BACK"
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10440
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "SUBMIT"
      Height          =   495
      Left            =   8880
      MaskColor       =   &H00FFFFFF&
      MousePointer    =   2  'Cross
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10440
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   8640
      TabIndex        =   48
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H80000008&
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   14160
      Top             =   1920
      Width           =   5295
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   495
      Left            =   17040
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   2415
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "POLICY DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   17400
      TabIndex        =   66
      Top             =   8400
      Width           =   1845
   End
   Begin VB.Shape Shape14 
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      Height          =   1335
      Left            =   14160
      Top             =   8880
      Width           =   5295
   End
   Begin VB.Shape Shape13 
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      Height          =   1335
      Left            =   8160
      Top             =   8880
      Width           =   5775
   End
   Begin VB.Shape Shape12 
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      Height          =   1335
      Left            =   1560
      Top             =   8880
      Width           =   6255
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "POLICY "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   65
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "POLICY ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   64
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF INSURANCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   63
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "AMOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14760
      TabIndex        =   62
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "PERIOD OF POLICY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14760
      TabIndex        =   61
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLICK TO GET ID"
      Height          =   195
      Left            =   3600
      TabIndex        =   60
      Top             =   9720
      Width           =   1305
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CLAIMANT DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16920
      TabIndex        =   51
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Shape Shape11 
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      Height          =   495
      Left            =   16680
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   2775
   End
   Begin VB.Shape Shape10 
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      Height          =   2175
      Left            =   1560
      Top             =   6000
      Width           =   6255
   End
   Begin VB.Shape Shape9 
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      Height          =   2175
      Left            =   8160
      Top             =   6000
      Width           =   5775
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   2175
      Left            =   14160
      Top             =   6000
      Width           =   5295
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      ForeColor       =   &H80000013&
      Height          =   375
      Left            =   2040
      TabIndex        =   50
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "PHONE NO"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8640
      TabIndex        =   49
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "SEX"
      ForeColor       =   &H80000013&
      Height          =   255
      Left            =   2040
      TabIndex        =   47
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "MARITAL STATUS"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8640
      TabIndex        =   46
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF BIRTH"
      Height          =   375
      Left            =   14760
      TabIndex        =   45
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "AGE"
      Height          =   375
      Left            =   15720
      TabIndex        =   44
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      ForeColor       =   &H80000013&
      Height          =   375
      Left            =   8520
      TabIndex        =   28
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "MARITAL STATUS"
      ForeColor       =   &H80000013&
      Height          =   495
      Left            =   8520
      TabIndex        =   27
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "SEX"
      ForeColor       =   &H80000013&
      Height          =   255
      Left            =   8520
      TabIndex        =   24
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      ForeColor       =   &H80000013&
      Height          =   375
      Left            =   8520
      TabIndex        =   22
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   375
      Left            =   17040
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENT DETAILS"
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
      Left            =   17280
      TabIndex        =   30
      Top             =   1440
      Width           =   1830
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF BIRTH"
      ForeColor       =   &H80000013&
      Height          =   255
      Left            =   14520
      TabIndex        =   25
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT NO"
      ForeColor       =   &H80000013&
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   8160
      Top             =   1920
      Width           =   5775
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   1560
      Top             =   1920
      Width           =   6255
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "SALARY"
      ForeColor       =   &H80000013&
      Height          =   495
      Left            =   1800
      TabIndex        =   29
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "AGE"
      ForeColor       =   &H80000013&
      Height          =   255
      Left            =   15000
      TabIndex        =   26
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER NAME"
      ForeColor       =   &H80000013&
      Height          =   375
      Left            =   1800
      TabIndex        =   21
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER ID"
      ForeColor       =   &H80000013&
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      DrawMode        =   6  'Mask Pen Not
      Height          =   9015
      Index           =   0
      Left            =   1200
      Top             =   1320
      Width           =   18495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CREATE A NEW POLICY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   795
      Left            =   7080
      TabIndex        =   2
      Top             =   360
      Width           =   6915
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      DrawMode        =   6  'Mask Pen Not
      Height          =   975
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   14655
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   0
      Picture         =   "addpol.frx":0262
      Top             =   0
      Width           =   24000
   End
End
Attribute VB_Name = "addpol1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
Text4(7).Text = "m"
Check2.Value = 0
End Sub

Private Sub Check2_Click()
Text4(7).Text = "um"
Check1.Value = 0
End Sub

Private Sub Check3_Click()
Text11.Text = "m"
Check4.Value = 0
End Sub

Private Sub Check4_Click()
Text11.Text = "um"
Check3.Value = 0
End Sub

Private Sub Command4_Click()
Text4(5).Text = Combo2.Text + "-" + mm.Text + "-" + Combo4.Text
Dim X As Integer
X = Format$(Now, "yyyy") - Combo4.Text
Text4(6) = X
End Sub

Private Sub Command1_Click()
j = 0
If Combo1.Text = "SELECT THE POLICY" Then
    j = 1
ElseIf Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Then
    j = 1

Else: For i = 0 To i < Text4.Count - 1
        If Text4(i).Text = "" Then
            j = 1
            Exit For
        End If
        Next i
End If
If j = 1 Then
MsgBox "FIRST ENTER ALL VALUES"
Else
    r3.Open "select * from sales", cn, 2, 3
    rs.Open "select * from insurance", cn, 2, 3
    r4.Open "select * from claimant", cn, 2, 3
    strsql = "insert into ph values(" & Text4(0).Text & ",'" & Text4(1).Text & "','" & Text4(2).Text & "'," & Text4(3).Text & ",'" & Text5.Text & "'," & Text6.Text & ",'" & Text4(5).Text & "'," & Text4(6).Text & ",'" & Text4(7).Text & "','" & Text4(4).Text & "')"
    strsql1 = "insert into sales values(" & agpg.Label5.Caption & "," & 0.05 * Text2.Text & "," & Text4(0).Text & "," & Label16.Caption & ")"
    strsql2 = "insert into insurance values(" & Text2.Text / Text3.Text & "," & Text2.Text & "," & Text3.Text & "," & Text4(0).Text & "," & Label16.Caption & ",' " & Text1.Text & "')"
    strsql3 = "insert into claimant values('" & Text7.Text & "','" & Text9.Text & "'," & Text4(0).Text & "," & Label16.Caption & ",'" & Text10.Text & "','" & Text12.Text & "','" & Text11.Text & "'," & Text13.Text & "," & Text8.Text & ")"
    If z = 1 Then
    cn.Execute strsql
    End If
    cn.Execute strsql1
    cn.Execute strsql2
    cn.Execute strsql3
    MsgBox "added"
End If
End Sub

Private Sub Command2_Click()
If Text4(0) = "" Then
    MsgBox "ENTER ID FIRST"
ElseIf r2.EOF = True Then
    MsgBox "NO RECORD PRESENT !!!!"
    z = 1
Else
    r2.MoveFirst
    Do While Not r2.EOF
        If r2.Fields(0) = Text4(0).Text Then
            Text4(1).Text = r2.Fields(1)
            Text4(2).Text = r2.Fields(2)
            Text4(3).Text = r2.Fields(3)
            Text5.Text = r2.Fields(4)
            Text6.Text = r2.Fields(5)
            Text4(5).Text = Format$(r2.Fields(6), "dd-mmm-yyyy")
            Combo4.Text = Format$(r2.Fields(6), "yyyy")
            Text4(6).Text = r2.Fields(7)
            Text4(7).Text = r2.Fields(8)
            Text4(4).Text = r2.Fields(9)
            z = 0
            Exit Do
        Else
            z = 1
        End If
    r2.MoveNext
    Loop
    If z = 1 Then
        MsgBox "no record found!!!"
    End If
End If
End Sub

Private Sub Command3_Click()
agpg.Show
agpg.Label13.Visible = True
Unload Me
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()
Text12.Text = Combo3.Text + "-" + Combo5.Text + "-" + Combo6.Text
Dim X As Integer
X = Format$(Now, "yyyy") - Combo6.Text
Text13.Text = X
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command8_Click()
Text1.Text = Combo7.Text + "-" + Combo8.Text + "-" + Combo9.Text
End Sub

Private Sub Form_Load()
On Error Resume Next
cmd1 = "Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False"

Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set r4 = New ADODB.Recordset
Set r1 = New ADODB.Recordset
Set r2 = New ADODB.Recordset
Set r3 = New ADODB.Recordset
With cn
 .ConnectionString = cmd1
.CursorLocation = adUseClient
 .Open
 End With
    
    r1.Open "select * from policy", cn, 2, 3
    r2.Open "select * from ph", cn, 2, 3
    Do While Not r1.EOF
        Combo1.AddItem (r1.Fields(1))
    r1.MoveNext
    
    
 Loop
End Sub



Private Sub Label16_Click()
r1.MoveFirst
Do While Not r1.EOF
    If Combo1.Text = r1.Fields(1) Then
        Label16.Caption = r1.Fields(0)
        Exit Do
        End If
    r1.MoveNext
 Loop
End Sub



Private Sub Option1_Click(Index As Integer)

Text4(4).Text = ""
Text4(4).Text = "m"
End Sub

Private Sub Option2_Click()
Text4(4).Text = ""
Text4(4).Text = "f"
End Sub

Private Sub Option3_Click()
Text10.Text = ""
Text10.Text = "m"
End Sub

Private Sub Option4_Click()
Text10.Text = ""
Text10.Text = "f"
End Sub




Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
If ((KeyCode < 48 Or KeyCode > 57) And KeyCode <> 8) Then
    MsgBox ("invalid character")
    Text2.Text = ""
    End If
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
If ((KeyCode < 48 Or KeyCode > 57) And KeyCode <> 8) Then
    MsgBox ("invalid character")
    Text3.Text = ""
    End If
End Sub

Private Sub Text4_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 0 Or Index = 3 Then
    If ((KeyCode < 48 Or KeyCode > 57) And KeyCode <> 8) Then
    MsgBox ("invalid character")
    Text4(Index).Text = ""
    End If
End If
If Index = 1 Then
    If ((KeyCode < 65 Or KeyCode > 123) And KeyCode <> 8) Then
        MsgBox ("invalid character")
        Text4(1).Text = ""
    End If
End If
End Sub

Private Sub Text4_LostFocus(Index As Integer)
If Index = 3 Then
    If Val(Text4(3).Text) < 1000000000 Then
        MsgBox ("invalid contact no !!!")
        Text4(3).Text = ""
    End If
End If
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
If ((KeyCode < 48 Or KeyCode > 57) And KeyCode <> 8) Then
    MsgBox ("invalid character")
    Text6.Text = ""
    End If
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
If ((KeyCode < 65 Or KeyCode > 123) And KeyCode <> 8) Then
        MsgBox ("invalid character")
        Text7.Text = ""
    End If
End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
If ((KeyCode < 48 Or KeyCode > 57) And KeyCode <> 8) Then
    MsgBox ("invalid character")
    Text8.Text = ""
    End If
End Sub

Private Sub Text8_LostFocus()
If Val(Text8.Text) < 1000000000 Then
    MsgBox ("invalid contact no !!!")
    Text8.Text = ""
End If
End Sub
