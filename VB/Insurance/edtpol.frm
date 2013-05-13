VERSION 5.00
Begin VB.Form edtpol 
   Caption         =   "EDIT WINDOW"
   ClientHeight    =   11010
   ClientLeft      =   -885
   ClientTop       =   -510
   ClientWidth     =   18705
   LinkTopic       =   "Form2"
   Picture         =   "edtpol.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "ACCEPT "
      Height          =   375
      Left            =   9360
      TabIndex        =   56
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   15720
      TabIndex        =   55
      Text            =   "SELECT POLICY"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      ItemData        =   "edtpol.frx":70103
      Left            =   17520
      List            =   "edtpol.frx":70119
      TabIndex        =   54
      Text            =   "yyyy"
      Top             =   7200
      Width           =   735
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "edtpol.frx":70141
      Left            =   16800
      List            =   "edtpol.frx":70154
      TabIndex        =   53
      Text            =   "mm"
      Top             =   7200
      Width           =   615
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "edtpol.frx":70171
      Left            =   16080
      List            =   "edtpol.frx":70193
      TabIndex        =   52
      Text            =   "dd"
      Top             =   7200
      Width           =   615
   End
   Begin VB.CheckBox Check8 
      Caption         =   "UNMARRIED"
      Height          =   375
      Left            =   17280
      TabIndex        =   51
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CheckBox Check7 
      Caption         =   "MARRIED"
      Height          =   375
      Left            =   16080
      TabIndex        =   50
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CheckBox Check6 
      Caption         =   "FEMALE"
      Height          =   435
      Left            =   17160
      TabIndex        =   49
      Top             =   5640
      Width           =   975
   End
   Begin VB.CheckBox Check5 
      Caption         =   "MALE"
      Height          =   375
      Left            =   16080
      TabIndex        =   48
      Top             =   5640
      Width           =   975
   End
   Begin VB.CheckBox Check4 
      Caption         =   "UNMARRIED"
      Height          =   375
      Left            =   5880
      TabIndex        =   47
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      Caption         =   "MARRIED"
      Height          =   375
      Left            =   4560
      TabIndex        =   46
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000013&
      Caption         =   "FEMALE"
      Height          =   375
      Left            =   5880
      TabIndex        =   45
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000013&
      Caption         =   "MALE"
      Height          =   375
      Left            =   4560
      TabIndex        =   44
      Top             =   6240
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "edtpol.frx":701B6
      Left            =   6240
      List            =   "edtpol.frx":701CC
      TabIndex        =   43
      Text            =   "yyyy"
      Top             =   6720
      Width           =   735
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "edtpol.frx":701F4
      Left            =   5280
      List            =   "edtpol.frx":7020A
      TabIndex        =   42
      Text            =   "mm"
      Top             =   6720
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "edtpol.frx":7022C
      Left            =   4560
      List            =   "edtpol.frx":70260
      TabIndex        =   41
      Text            =   "dd"
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox Text13 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   405
      Left            =   16080
      TabIndex        =   32
      Top             =   8040
      Width           =   615
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   18360
      TabIndex        =   31
      Top             =   7200
      Width           =   855
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   18720
      TabIndex        =   30
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   18480
      TabIndex        =   29
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   16080
      TabIndex        =   28
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   16080
      MaxLength       =   10
      TabIndex        =   27
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H80000013&
      Height          =   495
      Left            =   16080
      TabIndex        =   26
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox name1 
      BackColor       =   &H80000013&
      Height          =   495
      Index           =   1
      Left            =   4560
      TabIndex        =   15
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox address 
      BackColor       =   &H80000013&
      Height          =   495
      Index           =   2
      Left            =   4560
      TabIndex        =   14
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox phone 
      BackColor       =   &H80000013&
      Height          =   495
      Index           =   3
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   13
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox sex 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   4
      Left            =   7320
      TabIndex        =   12
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox dob 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   5
      Left            =   7080
      TabIndex        =   11
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox age1 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   4560
      TabIndex        =   10
      Top             =   7320
      Width           =   855
   End
   Begin VB.TextBox ms 
      BackColor       =   &H80000013&
      Height          =   375
      Index           =   7
      Left            =   7320
      TabIndex        =   9
      Top             =   7920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox sal 
      BackColor       =   &H80000013&
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   9240
      Width           =   1695
   End
   Begin VB.TextBox pwd 
      BackColor       =   &H80000013&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4560
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   8520
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000013&
      Height          =   315
      Left            =   8760
      TabIndex        =   6
      Text            =   "SELECT POLICY HOLDER ID"
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000013&
      Caption         =   "EDIT"
      Height          =   735
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10320
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000013&
      Caption         =   "BACK"
      Height          =   615
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10440
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000013&
      Caption         =   "EDIT"
      Height          =   735
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000013&
      Caption         =   "SHOW CLAIMANT INFO"
      Height          =   735
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "SHOW POLICY HOLDER INFO"
      Height          =   735
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CLAIMANT INFO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   14640
      TabIndex        =   40
      Top             =   3000
      Width           =   2865
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   9015
      Left            =   12600
      Top             =   2640
      Width           =   6735
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AGE"
      Height          =   375
      Left            =   14040
      TabIndex        =   39
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATE OF BIRTH"
      Height          =   375
      Left            =   14040
      TabIndex        =   38
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MARITIAL STATUS"
      Height          =   375
      Left            =   14040
      TabIndex        =   37
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SEX"
      Height          =   255
      Left            =   14040
      TabIndex        =   36
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ADDRESS"
      Height          =   375
      Left            =   14040
      TabIndex        =   35
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PHONE NO"
      Height          =   375
      Left            =   14040
      TabIndex        =   34
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAME"
      Height          =   375
      Left            =   14040
      TabIndex        =   33
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "POLICY HOLDER INFO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   25
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   9015
      Left            =   1200
      Top             =   2640
      Width           =   7095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ADDRESS"
      Height          =   495
      Left            =   2160
      TabIndex        =   24
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONTACT NO"
      Height          =   495
      Left            =   2160
      TabIndex        =   23
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SEX"
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATE OF BIRTH"
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AGE"
      Height          =   495
      Left            =   2160
      TabIndex        =   20
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MARITIAL STATUS"
      Height          =   495
      Left            =   2160
      TabIndex        =   19
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PASSWORD"
      Height          =   495
      Left            =   2160
      TabIndex        =   18
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALARY"
      Height          =   495
      Left            =   2160
      TabIndex        =   17
      Top             =   9240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAME"
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EDIT WINDOW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   2
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "edtpol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
sex(4).Text = "m"
Check2.Value = 0
End Sub

Private Sub Check2_Click()
sex(4).Text = "f"
Check1.Value = 0
End Sub

Private Sub Check3_Click()
ms(7).Text = "m"
Check4.Value = 0
End Sub

Private Sub Check4_Click()
ms(7).Text = "um"
Check3.Value = 0
End Sub

Private Sub Check5_Click()
Text10.Text = "m"
Check6.Value = 0
End Sub

Private Sub Check6_Click()
Text10.Text = "f"
Check5.Value = 0
End Sub

Private Sub Check7_Click()
Text11.Text = "m"
Check8.Value = 0
End Sub

Private Sub Check8_Click()
Text11.Text = "um"
Check7.Value = 0
End Sub

Private Sub Combo4_LostFocus()
If Combo2.Text <> "dd" And Combo3.Text <> "mm" And Combo4.Text <> "yyyy" Then
Dim dat As Date
dat = Combo2.Text + "-" + Combo3.Text + "-" + Combo4.Text
dob(5).Text = Format$(dat, "dd-mmm-yyyy")
Dim X As Integer
X = Format$(Now, "yyyy") - Combo4.Text
age1(6).Text = X
End If
End Sub

Private Sub Combo7_LostFocus()
If Combo5.Text <> "dd" And Combo6.Text <> "mm" And Combo7.Text <> "yyyy" Then
Dim dat1 As Date
dat1 = Combo5.Text + "-" + Combo6.Text + "-" + Combo7.Text
Text12.Text = Format$(dat1, "dd-mmm-yyyy")
Dim X1 As Integer
X1 = Format$(Now, "yyyy") - Combo7.Text
Text13.Text = X1
End If
End Sub

Private Sub Command1_Click()
If Combo1.Text = "SELECT POLICY HOLDER ID" Then
MsgBox "SELECT POLICY HOLDER ID FIRST !!!"
ElseIf rs.RecordCount = 0 Then
    MsgBox "NO RECORDS PRESENT !!!!!"
Else
rs.MoveFirst
Do While Not rs.EOF
    If rs.Fields(0) = Combo1.Text Then
        name1(1).Text = rs.Fields(1)
        address(2).Text = rs.Fields(2)
        phone(3).Text = rs.Fields(3)
        sex(4).Text = rs.Fields(9)
        If rs.Fields(9) = "m" Then
        Check1.Value = 1
        Check2.Value = 0
        ElseIf rs.Fields(9) = "f" Then
        Check1.Value = 0
        Check2.Value = 1
        End If
        dob(5).Text = Format$(rs.Fields(6), "dd-mmm-yyyy")
        age1(6).Text = rs.Fields(7)
        ms(7).Text = rs.Fields(8)
        If rs.Fields(8) = "m" Then
        Check3.Value = 1
        Check4.Value = 0
        ElseIf rs.Fields(8) = "um" Then
        Check3.Value = 0
        Check4.Value = 1
        End If
        pwd.Text = rs.Fields(4)
        sal.Text = rs.Fields(5)
     Exit Do
     End If
    rs.MoveNext
 Loop
End If
End Sub

Private Sub Command2_Click()
If Combo1.Text = "SELECT POLICY HOLDER ID" Or Combo8.Text = "SELECT POLICY" Then
MsgBox "SELECT POLICY HOLDER ID AND POLICY FIRST!!!"
ElseIf r2.RecordCount = 0 Then
    MsgBox "NO RECORDS PRESENT !!!!!"
Else
r2.MoveFirst
r1.MoveFirst
Do While Not r1.EOF
    If r1.Fields(1) = Combo8.Text Then
        Exit Do
    End If
    r1.MoveNext
Loop
Do While Not r2.EOF
    If r2.Fields(2) = Combo1.Text And r2.Fields(3) = r1.Fields(0) Then
        Text7.Text = r2.Fields(0)
        Text9.Text = r2.Fields(2)
        Text8.Text = r2.Fields(8)
        Text10.Text = r2.Fields(4)
        If r2.Fields(4) = "m" Then
        Check5.Value = 1
        Check6.Value = 0
        ElseIf r2.Fields(4) = "f" Then
        Check5.Value = 0
        Check6.Value = 1
        End If
        Text12.Text = Format$(r2.Fields(5), "dd-mmm-yyyy")
        Text13.Text = r2.Fields(7)
        Text11.Text = r2.Fields(6)
        If r2.Fields(6) = "m" Then
        Check7.Value = 1
        Check8.Value = 0
        ElseIf r2.Fields(6) = "um" Then
        Check7.Value = 0
        Check8.Value = 1
        End If
     Exit Do
     End If
    r2.MoveNext
 Loop
End If
End Sub

Private Sub Command3_Click()
name1(1).Text = ""
address(2).Text = ""
phone(3).Text = ""
sal.Text = ""
dob(5).Text = ""
ms(7).Text = ""
End Sub

Private Sub Command4_Click()
'Dim flag As Integer
'flag = 0
If Combo1.Text = "SELECT POLICY HOLDER ID" Then
 MsgBox "invalid entry"
 
ElseIf name1(1) = "" Or address(2) = "" Or phone(3) = "" Or (Check1.Value = 0 And Check2.Value = 0) Or (Check3.Value = 0 And Check4.Value = 0) Or sal = "" Then
        MsgBox "COMPLETE THE DETAILS FIRST!!!"
'ElseIf Not (Combo2.Text = "dd" And Combo3.Text = "mm" And Combo4.Text = "yyyy") Then
       ' If (Combo2.Text <> "dd" And Combo3.Text <> "mm" And Combo4.Text <> "yyyy") Then
       Else
            strsql = "update ph set sex='" & sex(4).Text & "',name= '" & name1(1).Text & "',address ='" & address(2).Text & "',phone= " & phone(3).Text & ",sal=" & sal.Text & ",dob='" & dob(5).Text & "',ms='" & ms(7).Text & "' where ph_key=" & Combo1.Text
            MsgBox strsql
            cn.Execute strsql
            MsgBox added
            flag = 1
       ' End If
        'If flag = 0 Then
           ' MsgBox "COMPLETE THE DETAILS FIRST!!!"
       ' End If
End If
End Sub

Private Sub Command5_Click()
agpg.Show
Unload Me
End Sub

Private Sub Command7_Click()
If Text10.Text = "" Or Text7.Text = "" Or Text9.Text = "" Or Text8.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Combo8.Text = "SELECT POLICY" Then
    MsgBox "Enter all values first"
Else
strsql2 = "update claimant set sex='" & Text10.Text & "',name= '" & Text7.Text & "',address ='" & Text9.Text & "',phone= " & Text8.Text & ",dob='" & Text12.Text & "',ms='" & Text11.Text & "' where p_holder=" & Combo1.Text
            MsgBox strsql2
            cn.Execute strsql2
            MsgBox "added"
End If
End Sub

Private Sub Command8_Click()
If r3.RecordCount <> 0 Then
    r3.MoveFirst
    Do While Not r3.EOF
        If r3.Fields(1) = Combo1.Text Then
            Combo8.AddItem (r3.Fields(8))
        End If
        r3.MoveNext
    Loop
End If
End Sub

Private Sub Form_Load()

On Error Resume Next
cmd1 = "Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False"
Set cn = New adodb.Connection
Set rs = New adodb.Recordset
Set r1 = New adodb.Recordset
Set r2 = New adodb.Recordset
Set r3 = New adodb.Recordset
Set r4 = New adodb.Recordset
With cn
 .ConnectionString = cmd1
.CursorLocation = adUseClient
 .Open
 End With
 rs.Open "SELECT * FROM ph", cn, 2, 3
 r2.Open "SELECT * FROM claimant", cn, 2, 3
 r3.Open "SELECT * FROM client", cn, 2, 3
 r1.Open "SELECT * FROM policy", cn, 2, 3
 r4.Open "select * from sales", cn, 2, 3
 If rs.RecordCount <> 0 Then
 rs.MoveFirst
 End If
If r4.RecordCount <> 0 Then
 r4.MoveFirst
 End If
Do While Not r4.EOF
    If agpg.Label5.Caption = r4.Fields(0) Then
        Combo1.AddItem (r4.Fields(2))
    End If
    r4.MoveNext
Loop
r4.MoveFirst
rs.MoveFirst
End Sub

Private Sub name1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 1 Then
    If (KeyCode < 65 Or KeyCode > 123) Then
    MsgBox ("invalid character")
    name1(1).Text = ""
    End If
End If
End Sub

Private Sub phone_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If (KeyCode < 48 Or KeyCode > 57) Then
    MsgBox ("invalid character")
    phone(3).Text = ""
End If
End Sub

Private Sub phone_LostFocus(Index As Integer)
If Val(phone(3).Text) < 1000000000 Then
    MsgBox ("invalid contact no !!!")
    phone(3).Text = ""
End If
End Sub

Private Sub sal_KeyUp(KeyCode As Integer, Shift As Integer)
If (KeyCode < 48 Or KeyCode > 57) Then
    MsgBox ("invalid character")
    sal.Text = ""
End If
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
If (KeyCode < 65 Or KeyCode > 123) Then
    MsgBox ("invalid character")
    Text7.Text = ""
    End If
End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
If (KeyCode < 48 Or KeyCode > 57) Then
    MsgBox ("invalid character")
    Text8.Text = ""
End If
End Sub
