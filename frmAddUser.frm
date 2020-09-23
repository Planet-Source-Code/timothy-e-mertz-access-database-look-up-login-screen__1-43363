VERSION 5.00
Begin VB.Form frmAddUser 
   Caption         =   "Add User"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddUser 
      Caption         =   "&Add User!"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtVerifyPassword 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   900
      Width           =   2775
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   195
      Width           =   3135
   End
   Begin VB.Label lblVerifyPassword 
      Caption         =   "Verify Password:"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1455
      Width           =   735
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   945
      Width           =   735
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "User:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddUser_Click()
On Error Resume Next
If txtUser.Text = "" Or txtPassword.Text = "" Then
MsgBox "Please Insert A User Name And/Or Password.", vbExclamation, "Error!"
Exit Sub
Else
    If txtVerifyPassword.Text = "" Then
    MsgBox "Please Verify Password.", vbExclamation, "Error!"
    Else
        If txtPassword.Text <> txtVerifyPassword Then
        MsgBox "Passwords Do Not Match.", vbExclamation, "Error!"
        Else
mrstDB.AddNew
mrstDB("fldUser") = txtUser.Text
mrstDB("fldPassword") = txtPassword.Text
mrstDB.Update
Unload Me
        End If
    End If
End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
Unload Me
End Sub

