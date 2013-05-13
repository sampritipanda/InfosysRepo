VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form agpg 
   BackColor       =   &H00FFFFFF&
   Caption         =   "AGENT"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15495
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000013&
      Caption         =   "LOGOUT"
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000013&
      Caption         =   "CHANGE AGENT PASSWORD"
      Height          =   735
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "CREATE NEW POLICY"
      Height          =   735
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000013&
      Caption         =   "EDIT POLICY"
      Height          =   615
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000013&
      Height          =   315
      Left            =   16680
      TabIndex        =   3
      Text            =   "SELECT POLICY"
      Top             =   5880
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H80000013&
      Height          =   315
      Left            =   16680
      TabIndex        =   2
      Text            =   "SELECT CUSTOMER KEY"
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000013&
      Caption         =   "DELETE SELECTED POLICY"
      Height          =   735
      Left            =   15960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   2655
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg1 
      Height          =   6495
      Left            =   5760
      TabIndex        =   11
      Top             =   2040
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   11456
      _Version        =   393216
      BackColor       =   16761024
      BackColorBkg    =   -2147483629
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "POLICY INFO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   23
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PERSONAL INFO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   22
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Shape Shape3 
      Height          =   7815
      Left            =   240
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label AGE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AGENT ID"
      Height          =   375
      Left            =   600
      TabIndex        =   21
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AGENT NAME"
      Height          =   375
      Left            =   600
      TabIndex        =   19
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONTACT  NO"
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ADDRESS"
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "***NOTE-CHANGES WILL BE REFLECTED IN NEXT LOGIN"
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   8760
      Width           =   5535
   End
   Begin VB.Shape Shape2 
      Height          =   9255
      Left            =   5400
      Top             =   1080
      Width           =   9135
   End
   Begin VB.Shape Shape1 
      Height          =   7815
      Left            =   14760
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "POLICY"
      Height          =   375
      Left            =   15240
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "POLICY HOLDER"
      Height          =   495
      Left            =   15240
      TabIndex        =   6
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "POLICY KEY"
      Height          =   375
      Left            =   15240
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CLICK ME"
      Height          =   375
      Left            =   16680
      TabIndex        =   4
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AGENT INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7095
      TabIndex        =   0
      Top             =   240
      Width           =   4905
   End
End
Attribute VB_Name = "agpg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
addpol1.Show
End Sub


Private Sub Command2_Click()
Dim flag As Integer
    
    
If Combo1.Text = "SELECT POLICY" Or Label12.Caption = "CLICK ME" Or Combo2.Text = "SELECT CUSTOMER KEY" Then
    MsgBox "ENTER COMPLETE DETAILS !!!!"
Else
    r2.MoveFirst
    Do While Not r2.EOF
        If r2.Fields(8) = Combo1.Text And r2.Fields(1) = Combo2.Text Then
            If MsgBox("Are you sure you to delete this customer?", vbYesNo) = vbYes Then
                strsql1 = "delete from insurance where p_holder=" & Combo2.Text & " and policy_key=" & Label12.Caption
                strsql2 = "delete from sales where p_holder=" & Combo2.Text & " and policy_key=" & Label12.Caption
                strsql3 = "delete from claimant where p_holder=" & Combo2.Text & " and policy_key=" & Label12.Caption
            
                If r4.RecordCount <> 0 Then
                    r4.MoveFirst
                Do While Not r4.EOF
                    If r4.Fields(1) = Combo2.Text Then
                        flag = flag + 1
                    End If
                    r4.MoveNext
                Loop
                cn.Execute strsql1
                cn.Execute strsql2
                cn.Execute strsql3
                If flag = 1 Then
                    strsql4 = "delete from ph where ph_key=" & Combo2.Text
                    MsgBox strsql4
                    cn.Execute strsql4
                End If
                MsgBox "deleted"
                Exit Do
            End If
        End If
        r2.MoveNext
   
End If
Loop
End If
End Sub

Private Sub Command3_Click()
pass1.Show
End Sub

Private Sub Command4_Click()
edtpol.Show
End Sub


Private Sub Command5_Click()
  
End Sub

Private Sub Command6_Click()
main1.Show
Unload Me
End Sub

Private Sub Form_Load()

On Error Resume Next
cmd1 = "Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False"
Set cn = New adodb.Connection
Set rs = New adodb.Recordset
Set rk = New adodb.Recordset
Set r1 = New adodb.Recordset
Set r2 = New adodb.Recordset
Set r3 = New adodb.Recordset
Set r4 = New adodb.Recordset
Set r5 = New adodb.Recordset
With cn
 .ConnectionString = cmd1
.CursorLocation = adUseClient
 .Open
 End With

  rs.Open "SELECT * FROM agent", cn, 2, 3
  rk.Open "SELECT * FROM client", cn, 2, 3
  r1.Open "SELECT * FROM policy", cn, 2, 3
  
  Do While Not rs.EOF
    If rs.Fields(3) = j Then
        Label5.Caption = rs.Fields(3)
        Label6.Caption = rs.Fields(0)
        Label7.Caption = rs.Fields(1)
        Label8.Caption = rs.Fields(2)
        Exit Do
    End If
    rs.MoveNext
  Loop
  
  strsql = "select * from client where agent_key= " & j
  r2.Open strsql, cn, 2, 3
  Set dg1.DataSource = r2
  dg1.Refresh
  
  Do While Not r2.EOF
    Combo1.AddItem r2!Type
    Combo2.AddItem r2!ph_key
    r2.MoveNext
  Loop
  
  rs.MoveFirst
  
  Label13.Visible = False
  
  r4.Open "select * from client", cn, 2, 3
End Sub

Private Sub MSHFlexGrid1_Click()

End Sub

Private Sub Label12_Click()
If Label12.Caption = "CLICK ME" Then
    MsgBox "SELECT THE POLICY FIRST"
Else
r3.Open "SELECT * FROM policy", cn, 2, 3
If r3.RecordCount <> 0 Then
r3.MoveFirst
End If
Do While Not r3.EOF
    If r3.Fields(1) = Combo1.Text Then
        Label12.Caption = r3.Fields(0)
        Exit Do
    End If
    r3.MoveNext
Loop
End If
End Sub

