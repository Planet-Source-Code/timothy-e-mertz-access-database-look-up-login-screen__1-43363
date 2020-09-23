VERSION 5.00
Begin VB.Form frmAddDeleteUser 
   Caption         =   "User Changes"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Change Password"
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1335
      TabIndex        =   2
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton cmdAddUser 
      Caption         =   "&Add User"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteUser 
      Caption         =   "&Delete User"
      Height          =   495
      Left            =   1950
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ListBox lstUser 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmAddDeleteUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddUser_Click()
On Error Resume Next
frmAddUser.Show
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
mrstDB.Update
frmLogin.Show
Unload Me
End Sub

Private Sub cmdChangePassword_Click()
frmChangePassword.Show
End Sub

Private Sub cmdDeleteUser_Click()
lstUserInput = lstUser.ListIndex + 1
mrstDB.AbsolutePosition = lstUserInput

If lstUserInput = 0 Then
MsgBox "Select A User To Delete.", vbExclamation, "Error!"
Exit Sub
End If

YesNo = MsgBox("Are You Sure You Want To Delete This User?", vbYesNo, "Are You Sure?")
If YesNo = vbNo Then
Exit Sub
Else
mrstDB.Delete
lstUser.RemoveItem (lstUser.ListIndex)
mrstDB.Update
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
For UserNameInput = 1 To mrstDB.RecordCount
    mrstDB.AbsolutePosition = UserNameInput
    lstUser.AddItem mrstDB("fldUser")
Next UserNameInput

End Sub
