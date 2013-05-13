VERSION 5.00
Begin VB.Form addagnt 
   Caption         =   "Form10"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form10"
   ScaleHeight     =   8250
   ScaleWidth      =   15210
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "addagnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
admn.Show
Unload Me
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub



Private Sub Command3_Click()
strsql = "insert into agent values('" & Text1.Text & "','" & Text2.Text & "'," & Text3.Text & "," & Text4.Text & ",'" & Text5.Text & "')"
cn.Execute strsql
MsgBox "added"
rs.Close
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
