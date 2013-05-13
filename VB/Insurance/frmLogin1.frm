VERSION 5.00
Begin VB.Form frmLogin1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ADMINISTRATOR LOGIN"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H8000000D&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H8000000D&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim cmd1 As String
Dim rs As ADODB.Recordset
Dim rk As ADODB.Recordset
Dim i, j As Integer

Dim sqlcmd, sqk, strsql, strname, str1, rate, a, b, c As String


Private Sub cmdCancel_Click()
main1.Show
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim flag As Integer
flag = 0
If rs.RecordCount <> 0 Then
rs.MoveFirst
Do While Not rs.EOF
If txtUserName.Text = rs.Fields(0) And txtPassword.Text = rs.Fields(1) Then
flag = 1
Exit Do
Else
rs.MoveNext
End If
Loop
If flag = 0 Then
MsgBox "INVALID USER"
txtUserName.Text = ""
txtPassword.Text = ""
Else
admn.Show
Unload Me
End If
Else
    MsgBox "NO RECORDS PRESENT !!!!"
End If
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
    rs.Open "SELECT * FROM admlog", cn, 2, 3
End Sub

