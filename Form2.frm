VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encryption1 Notepad"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlEdit 
      Left            =   2280
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text File (*.txt)|*.txt"
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Encrypt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Paste"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Copy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cu&t"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save &Only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   4815
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Text1.Text = ""
End Sub

Private Sub Command2_Click()
On Error GoTo oops
Dim AFile As String
cdlEdit.ShowSave
AFile = cdlEdit.filename
If AFile <> "" Then CopyToFile (AFile)
MsgBox "The file has been saved under " & cdlEdit.filename, vbInformation, "Encryption1"
oops:
'Unload Form2
Exit Sub
End Sub

Private Sub Command3_Click()
Clipboard.Clear
Clipboard.SetText Text1.SelText
Text1.SelText = ""
End Sub

Private Sub Command4_Click()
Clipboard.Clear
Clipboard.SetText Text1.SelText
End Sub

Private Sub Command5_Click()
Text1.SelText = Clipboard.GetText
End Sub

Private Sub Command6_Click()
Dim AFile As String
cdlEdit.ShowSave
AFile = cdlEdit.filename
If AFile <> "" Then CopyToFile (AFile)
MsgBox "The file has been saved under " & cdlEdit.filename, vbInformation, "Encryption1"
frmSecret.txtFile.Text = cdlEdit.filename
Form2.Hide
AFile = cdlEdit.filename
If AFile <> "" Then CopyToBox (AFile)
frmSecret.Text1.SetFocus
Unload Form2
End Sub

Sub CopyToFile(TheFile As String)
Dim FileNum As Integer
FileNum = FreeFile
Open TheFile For Output As #FileNum
Print #FileNum, Text1.Text
Close #FileNum
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Sub CopyToBox(TheFile As String)
Dim FileNum As Integer
FileNum = FreeFile
Open TheFile For Input As #FileNum
frmSecret.Text1.Text = Input(LOF(FileNum), #FileNum)
Close #FileNum
End Sub

