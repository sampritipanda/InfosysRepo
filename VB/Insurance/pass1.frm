VERSION 5.00
Begin VB.Form pass1 
   Caption         =   "CHANGE PASSWORD"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   Picture         =   "pass1.frx":0000
   ScaleHeight     =   4380
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHANGE"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE AGENT PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RE-ENTER NEW PASSWORD"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2250
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER NEW PASSWORD"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER OLD PASSWORD"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1920
   End
End
Attribute VB_Name = "pass1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If rs.RecordCount <> 0 Then
    rs.MoveFirst
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
        MsgBox "COMPLETE THE DETAILS !!!"
    Else
        Do While Not rs.EOF
            If rs.Fields(4) = Text1.Text Then
                If Text2.Text = Text3.Text Then
                    strsql = "Update agent set pwd='" & Text2.Text & "'where id=" & rs.Fields(3)
                    cn.Execute strsql
                    MsgBox "password changed successfully"
                    Unload Me
                Else
                    MsgBox "re-enter new passwords"
                    Text2.Text = ""
                    Text3.Text = ""
                End If
            Exit Do
            End If
            
            rs.MoveNext
            If rs.EOF = True Then
                MsgBox "WRONG PASSWORD !! RE ENTER"
                Text2.Text = ""
                Text1.Text = ""
                Text3.Text = ""
            End If
       Loop
    End If
Else
    MsgBox "NO RECORD PRESENT !!!"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
cmd1 = "Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=False"
Set cn = New adodb.Connection

Set r2 = New adodb.Recordset
With cn
 .ConnectionString = cmd1
.CursorLocation = adUseClient
 .Open
 End With


End Sub

