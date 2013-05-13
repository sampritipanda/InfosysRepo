VERSION 5.00
Begin VB.Form frmLockerAlloc 
   Caption         =   "Locker Allocation and Deallocation"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4200
      TabIndex        =   6
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdDeAllocate 
      Caption         =   "Deallocate"
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Frame frameAllocateLockers 
      Caption         =   "Allocate Lockers"
      Height          =   4215
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
      Begin VB.ListBox lstAllocatedLockers 
         Height          =   3660
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdAllocate 
      Caption         =   "Allocate"
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.ComboBox cboAvailableLockers 
      Height          =   420
      ItemData        =   "frmLockerAlloc.frx":0000
      Left            =   2760
      List            =   "frmLockerAlloc.frx":0002
      TabIndex        =   1
      Text            =   "Available Lockers"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblLockerNo 
      AutoSize        =   -1  'True
      Caption         =   "Locker No. to allocate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2265
   End
End
Attribute VB_Name = "frmLockerAlloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAllocate_Click()
    locker = cboAvailableLockers.Text
    lstAllocatedLockers.AddItem locker
    cboAvailableLockers.RemoveItem (cboAvailableLockers.ListIndex)
    lockerIndex = Val(Mid(locker, 14))
    cboAvailableLockers.SelText = "locker number " & lockerIndex + 1
End Sub

Private Sub Form_Load()
    cboAvailableLockers.AddItem "locker number 1"
    cboAvailableLockers.AddItem "locker number 2"
    cboAvailableLockers.AddItem "locker number 3"
    cboAvailableLockers.AddItem "locker number 4"
    cboAvailableLockers.AddItem "locker number 5"
    cboAvailableLockers.AddItem "locker number 6"
    cboAvailableLockers.AddItem "locker number 7"
    cboAvailableLockers.AddItem "locker number 8"
    cboAvailableLockers.AddItem "locker number 9"
    cboAvailableLockers.AddItem "locker number 10"
End Sub

'****************************************************************************
    ' Purpose : To de-allocate the selected lockers
    ' Parameters:
    '        1) Input: None
    '        2) Output : None
    ' Return Value(s) : None
    ' Limitations, if any : None
'****************************************************************************

Private Sub cmdDeallocate_Click()

    Dim iCount As Integer
    
    iCount = lstAllocatedLockers.ListCount - 1
    
    While (iCount >= 0)
        If lstAllocatedLockers.Selected(iCount) Then
            cboAvailableLockers.AddItem lstAllocatedLockers.List(iCount)
            lstAllocatedLockers.RemoveItem (iCount)
        End If
        iCount = iCount - 1
    Wend
End Sub

