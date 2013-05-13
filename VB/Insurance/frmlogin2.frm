VERSION 5.00
Begin VB.Form frmlogin2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "AGENT LOGIN"
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3780
   LinkTopic       =   "Form6"
   ScaleHeight     =   1530
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000D&
      Caption         =   "cancel"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000D&
      Caption         =   "enter"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AGENT ID"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmlogin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
Dim flag As Integer
flag = 0
If rs.RecordCount <> 0 Then
rs.MoveFirst
Do While Not rs.EOF
    If Text1.Text = rs.Fields(3) And Text2.Text = rs.Fields(4) Then
        flag = 1
        Exit Do
    Else
        rs.MoveNext
    End If
Loop
If flag = 0 Then
    MsgBox "INVALID USER"
    Text1.Text = ""
    Text2.Text = ""
Else
   
    j = rs.Fields(3)
    agpg.Show
    Unload Me
End If
Else
    MsgBox "NO RECORDS PRESENT !!!!"
End If
End Sub

Private Sub Command4_Click()
main1.Show
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
cmd1 = "Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False"

Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
With cn
 .ConnectionString = cmd1
.CursorLocation = adUseClient
 .Open
 End With
    rs.Open "SELECT * FROM agent", cn, 2, 3
End Sub
