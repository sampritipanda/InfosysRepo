VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form2"
   ScaleHeight     =   5790
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdResult 
      Caption         =   "Result"
      Height          =   615
      Left            =   1920
      TabIndex        =   12
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox txtText6 
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtText5 
      Height          =   615
      Left            =   2880
      TabIndex        =   9
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtText4 
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtText3 
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtText2 
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtText1 
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblResult 
      Caption         =   "Result"
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblMark3 
      Caption         =   "Enter Mark 3"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblMark2 
      Caption         =   "Enter Mark 2"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblMark1 
      Caption         =   "Enter Mark 1"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblName 
      Caption         =   "Enter Student Name"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblNo 
      Caption         =   "Enter Student No."
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sum As Double
Dim avg As Double

Private Sub cmdResult_Click()
    sum = Val(txtText3.Text) + Val(txtText4.Text) + Val(txtText5.Text)
    avg = sum / 3
    txtText6.Text = IIf(avg >= 50, "Passed", "Failed")
    txtText6.Visible = True
End Sub

Private Sub txtText3_LostFocus()
    If Val(txtText3.Text) > 100 Then
        X = MsgBox("Subject Mark cannot exceed 100", vbExclamation)
        txtText3.SetFocus
        txtText3.Text = ""
    End If
End Sub

Private Sub txtText4_LostFocus()
    If Val(txtText4.Text) > 100 Then
        X = MsgBox("Subject Mark cannot exceed 100", vbExclamation)
        txtText4.SetFocus
        txtText4.Text = ""
    End If
End Sub

Private Sub txtText5_LostFocus()
    If Val(txtText5.Text) > 100 Then
        X = MsgBox("Subject Mark cannot exceed 100", vbExclamation)
        txtText5.SetFocus
        txtText5.Text = ""
    End If
End Sub
