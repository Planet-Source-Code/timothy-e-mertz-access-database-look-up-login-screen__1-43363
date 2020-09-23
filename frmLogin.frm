VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   495
      Left            =   2430
      TabIndex        =   5
      Top             =   5280
      Width           =   735
   End
   Begin VB.ListBox lstUser 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdAddDeleteUser 
      Caption         =   "&Change User Information"
      Height          =   495
      Left            =   1095
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   1305
      TabIndex        =   3
      Top             =   4440
      Width           =   690
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNextUser_Click()
mrstDB.MoveNext
If mrstDB.EOF Then
    mrstDB.MoveLast
End If

End Sub

Private Sub cmdAddDeleteUser_Click()
frmAddDeleteUser.Show
Unload Me
End Sub


Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdLogin_Click()
Dim lstUserIndex As String
lstUserIndex = lstUser.ListIndex + 1

If lstUserIndex = 0 Then
MsgBox "Sorry You Are Required To Select A User Name.", vbExclamation, "Select A User Name!"
End If

If lstUserIndex <> 0 Then
mrstDB.AbsolutePosition = lstUserIndex
If txtPassword.Text = mrstDB("fldPassword") Then
    MsgBox "PASS!!!", vbExclamation, "PASS!!!"
Else
    MsgBox "FAIL!!!", vbCritical, "FAIL!!!"
End If
End If
End Sub

Private Sub Form_Initialize()
On Error Resume Next
DBConnection.Mode = adModeReadWrite
DBConnection.CursorLocation = adUseClient
DBConnection.Provider = "Microsoft.Jet.OLEDB.3.51"
DBConnection.ConnectionString = _
    "Persist Security Info=False;" & _
    "Data Source=" & App.Path & "\Login.mdb"

    
DBConnection.Open                                        'for this? This is my only thing left to fix'
Set DBCommand.ActiveConnection = DBConnection
DBCommand.CommandType = adCmdTable
DBCommand.CommandText = "tblUserPassword"
mrstDB.LockType = adLockOptimistic
mrstDB.CursorLocation = adUseClient
mrstDB.CursorType = adOpenKeyset
mrstDB.Open DBCommand

End Sub

Private Sub Form_Load()

txtPassword.Enabled = False
txtPassword.BackColor = &H8000000F

Call LoadCurrentRecord
End Sub

Private Sub LoadCurrentRecord()
For UserNameInput = 1 To mrstDB.RecordCount
    mrstDB.AbsolutePosition = UserNameInput
    lstUser.AddItem mrstDB("fldUser")
Next UserNameInput
End Sub

Private Sub Form_Terminate()
DBConnection.Close
End Sub

Private Sub lstUser_Click()
txtPassword.Enabled = True
txtPassword.BackColor = vbWhite
End Sub
