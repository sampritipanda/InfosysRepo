VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   1710
   ClientTop       =   1395
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   6780
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      Height          =   735
      Left            =   1320
      TabIndex        =   4
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtEname 
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Text            =   "Sajjad"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtEno 
      Height          =   735
      Left            =   3000
      TabIndex        =   2
      Text            =   "E001"
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblEname 
      Caption         =   "Enter Ename:"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblEno 
      Caption         =   "Enter ENO:"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDisplay_Click()
    confirm = MsgBox("Are you " & txtEname.Text, 4, "Confirm")
    If confirm = 6 Then
        X = MsgBox("Hi, " & txtEname.Text & Chr(10) & "Your employee number is " & txtEno.Text)
    Else
        X = MsgBox("Then why did you enter " & txtEname.Text & "'s name.", vbOKOnly, "Wrong info")
    End If
End Sub
