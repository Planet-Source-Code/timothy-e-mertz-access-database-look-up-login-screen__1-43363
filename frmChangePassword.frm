VERSION 5.00
Begin VB.Form frmChangePassword 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCurrentPassword 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   4800
      Width           =   2775
   End
   Begin VB.ListBox lstUser 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox txtVerifyPassword 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   6180
      Width           =   2775
   End
   Begin VB.CommandButton cmdAddUser 
      Caption         =   "C&hange Password"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   6900
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   6900
      Width           =   1215
   End
   Begin VB.Label lblCurrentPassword 
      Caption         =   "Current Password:"
      Height          =   435
      Left            =   120
      TabIndex        =   8
      Top             =   4650
      Width           =   765
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   5565
      Width           =   735
   End
   Begin VB.Label lblVerifyPassword 
      Caption         =   "Verify Password:"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   6075
      Width           =   735
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddUser_Click()
lstUserInput = lstUser.ListIndex + 1
mrstDB.AbsolutePosition = lstUserInput

If txtCurrentPassword.Text <> mrstDB("fldPassword") Then
    MsgBox "The Inputed Current Password Is Incorrect.", vbExclamation, "Error!"
    Exit Sub
End If

If txtPassword.Text = "" Then
    MsgBox "Please Insert A Password.", vbExclamation, "Error!"
Else
    If txtVerifyPassword.Text = "" Then
        MsgBox "Please Verify Password.", vbExclamation, "Error!"
    Else
        If txtPassword.Text <> txtVerifyPassword.Text Then
            MsgBox "Passwords Do Not Match.", vbExclamation, "Error!"
        Else
            YesNo = MsgBox("Are You Sure You Want To Change The Password?", vbYesNo, "Are You Sure?")
                        
            If YesNo = vbYes Then
            lstUserInput = lstUser.ListIndex + 1
            mrstDB.AbsolutePosition = lstUserInput
            
            mrstDB("fldPassword") = txtPassword.Text
            Unload Me
            Else
            Unload Me
            Exit Sub
            End If
        End If
    End If
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
For UserNameInput = 1 To mrstDB.RecordCount
    mrstDB.AbsolutePosition = UserNameInput
    lstUser.AddItem mrstDB("fldUser")
Next UserNameInput
End Sub
