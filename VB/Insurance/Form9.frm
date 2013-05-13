VERSION 5.00
Begin VB.Form aginfo 
   BackColor       =   &H8000000D&
   Caption         =   "Form9"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14205
   LinkTopic       =   "Form9"
   ScaleHeight     =   9075
   ScaleWidth      =   14205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "BACK"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1815
   End
End
Attribute VB_Name = "aginfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cn As ADODB.Connection
Dim cmd1 As String
Dim rs As ADODB.Recordset
Dim rk As ADODB.Recordset
Dim i, j As Integer

Dim sqlcmd, sqk, strsql, strname, str1, rate, a, b, c As String


Private Sub Command1_Click()
admn.Show
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
  Set dg.DataSource = rs
  dg.Refresh
 
End Sub
