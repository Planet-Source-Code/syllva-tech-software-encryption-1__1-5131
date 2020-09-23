VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Verification"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPassword.frx":0442
   ScaleHeight     =   2160
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox txtPassCheck 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Please enter the same password again..."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the password..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If txtPassword.Text <> txtPassCheck.Text Then
MsgBox "The two passwords are not the same!", vbExclamation, "Encryption 1"
Else
Password$ = txtPassword.Text
Unload Me
End If
End Sub

Private Sub Form_Load()
Password$ = ""
End Sub

Private Sub txtPassCheck_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Command1_Click
KeyAscii = 0
End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
txtPassCheck.SetFocus
KeyAscii = 0
End If
End Sub
