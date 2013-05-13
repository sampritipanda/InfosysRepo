VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form admn 
   BackColor       =   &H8000000D&
   Caption         =   "ADMINISTRATOR"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18345
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   Picture         =   "admn1.frx":0000
   ScaleHeight     =   9465
   ScaleWidth      =   18345
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "REPORT"
      Height          =   735
      Left            =   8880
      TabIndex        =   35
      Top             =   9720
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CHANGE PASSWORD"
      Height          =   735
      Left            =   5880
      TabIndex        =   34
      Top             =   9720
      Width           =   2175
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "admn1.frx":711F8
      Left            =   9000
      List            =   "admn1.frx":711FA
      TabIndex        =   33
      Text            =   "NEW AGENT MAINTAINING POLICES"
      Top             =   7800
      Width           =   3015
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H8000000E&
      Height          =   315
      Left            =   9000
      TabIndex        =   32
      Text            =   "SELECT AGENT TO DELETE"
      Top             =   7320
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Index           =   1
      Left            =   17760
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   17760
      PasswordChar    =   "*"
      TabIndex        =   22
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Index           =   1
      Left            =   17760
      MaxLength       =   10
      TabIndex        =   21
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Index           =   1
      Left            =   17760
      TabIndex        =   20
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   17760
      TabIndex        =   19
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000013&
      Caption         =   "EDIT"
      Height          =   735
      Left            =   16200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7920
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000013&
      Height          =   315
      Left            =   14760
      TabIndex        =   17
      Text            =   "---SELECT AGENT---"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000013&
      Caption         =   "ENTER"
      Height          =   375
      Left            =   17520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H80000013&
      Height          =   315
      Left            =   -6480
      TabIndex        =   15
      Text            =   "SELECT AGENT ID"
      Top             =   10920
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000009&
      Height          =   495
      Index           =   0
      Left            =   4320
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Index           =   0
      Left            =   4320
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Index           =   0
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Index           =   0
      Left            =   4320
      TabIndex        =   5
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Index           =   0
      Left            =   4320
      TabIndex        =   4
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000013&
      Caption         =   "CREATE"
      Height          =   735
      Index           =   0
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "DELETE"
      Height          =   735
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000013&
      Caption         =   "LOGOUT"
      Height          =   735
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9720
      Width           =   2415
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg 
      Height          =   4095
      Left            =   7680
      TabIndex        =   29
      Top             =   3120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7223
      _Version        =   393216
      BackColor       =   -2147483629
      BackColorSel    =   -2147483634
      BackColorBkg    =   -2147483645
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ADMINISTRATOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   855
      Left            =   8400
      TabIndex        =   31
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AGENT INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   495
      Index           =   2
      Left            =   8160
      TabIndex        =   30
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "EDIT AGENT INFO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   735
      Left            =   15600
      TabIndex        =   28
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   14760
      TabIndex        =   27
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE NUMBER"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   14760
      TabIndex        =   26
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   14760
      TabIndex        =   25
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   14760
      TabIndex        =   24
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   9135
      Left            =   14040
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CREATE A NEW AGENT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   405
      Left            =   1560
      TabIndex        =   14
      Top             =   720
      Width           =   3660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   13
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   12
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE NUMBER"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   11
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   10
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ID NUMBER"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   9
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ID NUMBER"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   14760
      TabIndex        =   0
      Top             =   6840
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   9015
      Left            =   480
      Top             =   240
      Width           =   6375
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      Height          =   7815
      Left            =   7320
      Top             =   1440
      Width           =   6255
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   1215
      Left            =   5400
      Top             =   9480
      Width           =   9615
   End
End
Attribute VB_Name = "admn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Combo3.Text = "SELECT AGENT" Or Combo4.Text = "NEW AGENT MAINTAINING POLICES" Then
    MsgBox "SELECT THE AGENT ID FIRST !!!"
ElseIf Combo3.Text = Combo4.Text Then
    MsgBox "NEW AGENT ID MUST DIFFER FROM DELETED AGENT"
Else
    Dim flag As Integer
    If r1.RecordCount <> 0 Then
        r1.MoveFirst
    End If
    If r2.RecordCount <> 0 Then
        r2.MoveFirst
    End If
    If r3.RecordCount <> 0 Then
        r3.MoveFirst
    End If
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
    End If
    
Do While Not r2.EOF
    If Combo3.Text = r2.Fields(0) Then
        r2.Fields(0) = Combo4.Text
    End If
    r2.MoveNext
Loop
Do While Not r1.EOF
    If Combo3.Text = r1.Fields(3) Then
        strsql6 = "delete from agent where id =" & Combo3.Text
        MsgBox strsql6
        cn.Execute strsql6
        MsgBox "deleted"
        Exit Do
    End If
    r1.MoveNext
Loop

If r1.EOF = True Then
    MsgBox "INVALID ENTRY !!!"
End If
End If
End Sub

Private Sub Command2_Click()
pass2.Show
End Sub

Private Sub Command3_Click()
DataReport2.Show
End Sub

Private Sub Command4_Click(Index As Integer)
Dim l As Integer
l = 0
r1.MoveFirst
Do While Not r1.EOF
    If Text4(0) = r1.Fields(3) Then
        MsgBox "ID ALREADY EXISTS"
        Text1(0).Text = ""
        Text5(0).Text = ""
        Text4(0).Text = ""
        Text3(0).Text = ""
        Text2(0).Text = ""
        l = 1
        Exit Do
    End If
r1.MoveNext
Loop
If l = 0 Then
    If Text1(0) = "" Or Text2(0) = "" Or Text3(0) = "" Or Text4(0) = "" Or Text5(0) = "" Then
        MsgBox "COMPLETE THE ENTRIES FIRST !!!!"
    Else
        strsql = "insert into agent values('" & Text1(0).Text & "','" & Text2(0).Text & "'," & Text3(0).Text & "," & Text4(0).Text & ",'" & Text5(0).Text & "')"
        cn.Execute strsql
        MsgBox "Agent successfully added"
        Text1(0).Text = ""
        Text5(0).Text = ""
        Text4(0).Text = ""
        Text3(0).Text = ""
        Text2(0).Text = ""
        
        r1.MoveFirst
    End If
End If
End Sub


Private Sub Command5_Click()
On Error Resume Next
cmd1 = "Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False"
Set cn1 = New adodb.Connection
With cn1
 .ConnectionString = cmd1
.CursorLocation = adUseClient
 .Open
 End With
 
 Set dg.DataSource = r1
  dg.Refresh
End Sub



Private Sub Command7_Click()
If Combo1.Text = "---SELECT AGENT---" Then
    MsgBox "SELECT THE AGENT ID FIRST !!!"
ElseIf Text1(1) = "" Or Text2(1) = "" Or Text3(1) = "" Or Text4(1) = "" Or Text5(1) = "" Then
    MsgBox "COMPLETE THE ENTRIES FIRST !!!!"
Else
    strsql1 = "update agent set name= '" & Text1(1) & "',address='" & Text2(1) & "',phone=" & Text3(1) & ",id= " & Text5(1).Text & " where id=" & Combo1.Text
    MsgBox strsql1
    cn.Execute strsql1
    MsgBox "Entry successfully editted"
End If
End Sub

Private Sub Command8_Click()

If Combo1.Text = "---SELECT AGENT---" Then
    MsgBox "SELECT THE AGENT ID FIRST !!!"
Else
    If r1.RecordCount <> 0 Then
    r1.MoveFirst
    Do While Not r1.EOF
        If r1.Fields(3) = Combo1.Text Then
        Text3(1) = r1.Fields(2)
        Text2(1) = r1.Fields(1)
        Text1(1) = r1.Fields(0)
        Text4(1) = r1.Fields(4)
        Text5(1) = r1.Fields(3)
        Exit Do
        End If
        r1.MoveNext
    Loop
    Else
        MsgBox "NO RECORD FOUND !!!!"
    End If
End If
End Sub

Private Sub Command9_Click()
main1.Show
Unload Me
End Sub

Private Sub Form_Load()

On Error Resume Next
cmd1 = "Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False"

Set cn = New adodb.Connection
Set rs = New adodb.Recordset
Set r1 = New adodb.Recordset
Set r2 = New adodb.Recordset
Set r3 = New adodb.Recordset
Set rk = New adodb.Recordset
Set r4 = New adodb.Recordset
With cn
 .ConnectionString = cmd1
.CursorLocation = adUseClient
 .Open
 End With
 r1.Open "select * from agent", cn, 2, 3
 r2.Open "select * from sales", cn, 2, 3
 r3.Open "select * from insurance", cn, 2, 3
 rs.Open "select * from claimant", cn, 2, 3
 rk.Open "select * from ph", cn, 2, 3
r1.MoveFirst
Do While Not r1.EOF
    Combo1.AddItem (r1.Fields(3))
    Combo3.AddItem (r1.Fields(3))
    Combo4.AddItem (r1.Fields(3))
    
    
    r1.MoveNext
Loop
 Set dg.DataSource = r1
  dg.Refresh

End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 0 Or Index = 1 Then
    If (KeyCode < 65 Or KeyCode > 123) Then
        MsgBox ("invalid character")
        If Index = 0 Then
            Text1(0).Text = ""
        End If
        If Index = 1 Then
            Text1(1).Text = ""
        End If
    End If
End If
End Sub

Private Sub Text3_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If ((KeyCode < 48 Or KeyCode > 57) And KeyCode <> 8) Then
    MsgBox ("invalid character")
    Text3(Index).Text = ""
    End If

End Sub

Private Sub Text3_LostFocus(Index As Integer)
    If Val(Text3(Index).Text) < 1000000000 Then
        MsgBox ("invalid contact no !!!")
        Text3(Index).Text = ""
    End If
End Sub

Private Sub Text4_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 0 Then
    If ((KeyCode < 48 Or KeyCode > 57) And KeyCode <> 8) Then
    MsgBox ("invalid character")
    Text4(0).Text = ""
    End If
End If
End Sub

