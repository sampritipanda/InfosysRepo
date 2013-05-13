VERSION 5.00
Begin VB.Form frmFirstApp 
   Caption         =   "My First VB Appllication"
   ClientHeight    =   4485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameFormBkColor 
      Caption         =   "Background Color of the Form"
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   5175
      Begin VB.OptionButton optRed 
         Caption         =   "Red"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optGreen 
         Caption         =   "Green"
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optBlue 
         Caption         =   "Blue"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   525
      Left            =   2520
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame frameTextProperties 
      Caption         =   "Property setting of the Text Box"
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   5175
      Begin VB.CheckBox chkLocked 
         Caption         =   "Locked"
         Height          =   495
         Left            =   3480
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "Visible"
         Height          =   495
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   1440
      MaxLength       =   25
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   660
   End
End
Attribute VB_Name = "frmFirstApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************
    ' Purpose : To Exit the application by
    ' Parameters:
    '        1) Input: None
    '        2) Output : None
    ' Return Value(s) : None
    ' Limitations, if any : None
'*******************************************

Private Sub cmdExit_Click()
    If MsgBox("Do you want to exit the application?", vbYesNo) = vbYes Then
        Unload Me
    End If
End Sub

'******************************************************************
    ' Purpose : To set background color to Red using option button
    ' Parameters:
    '        1) Input: None
    '        2) Output : None
    ' Return Value(s) : None
    ' Limitations, if any : None
'******************************************************************
Private Sub optRed_Click()
    frmFirstApp.BackColor = vbRed
End Sub

'******************************************************************
    ' Purpose : To set background color to Blue using option button
    ' Parameters:
    '        1) Input: None
    '        2) Output : None
    ' Return Value(s) : None
    ' Limitations, if any : None
'******************************************************************

Private Sub optBlue_Click()
    frmFirstApp.BackColor = vbBlue
End Sub

'*******************************************************************
    ' Purpose : To set background color to Green using option button
    ' Parameters:
    '        1) Input: None
    '        2) Output : None
    ' Return Value(s) : None
    ' Limitations, if any : None
'*******************************************************************

Private Sub optGreen_Click()
    frmFirstApp.BackColor = vbGreen
End Sub

'****************************************************************************
    ' Purpose : To explore the use of Enabled property
    ' Parameters:
    '        1) Input: None
    '        2) Output : None
    ' Return Value(s) : None
    ' Limitations, if any : None
'****************************************************************************

Private Sub chkEnabled_Click()
    If chkEnabled.Value = vbChecked Then
        txtName.Enabled = True
    Else
        txtName.Enabled = False
    End If
End Sub

'****************************************************************************
    ' Purpose : To explore the use of Locked property
    ' Parameters:
    '        1) Input: None
    '        2) Output : None
    ' Return Value(s) : None
    ' Limitations, if any : None
'****************************************************************************
Private Sub chkLocked_Click()
    If chkLocked.Value = vbChecked Then
        txtName.Locked = True
    Else
        txtName.Locked = False
    End If
End Sub

'****************************************************************************
    ' Purpose : To explore the use of Visible property
    ' Parameters:
    '        1) Input: None
    '        2) Output : None
    ' Return Value(s) : None
    ' Limitations, if any : None
'****************************************************************************
Private Sub chkVisible_Click()
    If chkVisible.Value = vbChecked Then
        txtName.Visible = True
    Else
        txtName.Visible = False
    End If
End Sub

'****************************************************************************
    ' Purpose : To accept only alphadets for the txtName text box
    ' Parameters:
    '        1) Input: KeyAscii - Integer ASCII code of current key press
    '        2) Value Output : None
    ' Return (s) : None
    ' Limitations, if any : None
'****************************************************************************

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 65) Or (KeyAscii > 90)) And _
    ((KeyAscii < 97) Or (KeyAscii > 122)) Then
        KeyAscii = 0
    End If
End Sub

