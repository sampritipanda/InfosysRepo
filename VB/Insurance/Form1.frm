VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form delet 
   Caption         =   "Form1"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5520
      TabIndex        =   2
      Top             =   480
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "delete"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg 
      Height          =   3735
      Left            =   2760
      TabIndex        =   0
      Top             =   1560
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6588
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "delet"
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
If MsgBox("Are you sure you to delete this agent?", vbYesNo) = vbYes Then
    

        rk.Delete
    
    Text1_Change
    If rs.RecordCount > 0 Then
    
             dg.Row = dg.Rows - 1
      dg_Click
   Else

            MsgBox "No customer record exists"

    End If
    
End If
End Sub

Private Sub dg_Click()
If rs.RecordCount > 0 Then
rk.Close
MsgBox dg.Row
dg.Col = 1

strsql = "select * from agent where id = " & dg.Text



rk.Open strsql, cn, 2, 3


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

   rs.Open "SELECT * FROM agent", cn, 2, 3
  Set dg.DataSource = rs
  dg.Refresh
  Set rk = New ADODB.Recordset

rk.Open "select * from agent", cn, 2, 3
End Sub

Private Sub Text1_Change()
strsql = "select * from agent where id like '" & Text1 & "%' "
rs.Close
rs.Open strsql, cn, 2, 3
Set dg.DataSource = rs

End Sub
