VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSecret 
   Caption         =   "Encryption 1"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7725
   Icon            =   "frmSecret.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Make &New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frmSecret.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Make my own file"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frmSecret.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Goodbye!"
      Top             =   6840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlOne 
      Left            =   480
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   10335
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   10335
   End
   Begin VB.CommandButton cmdDecipher 
      Caption         =   "&Decipher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frmSecret.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Make this file readable"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncipher 
      Caption         =   "&Encipher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frmSecret.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Make this file unreadable"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "frmSecret.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Look for a file"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "File Path:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   180
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "Make &New"
      End
      Begin VB.Menu mnuBarfour 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBrowse 
         Caption         =   "&Browse"
      End
      Begin VB.Menu mnuBartwo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileEncipher 
         Caption         =   "&Encipher"
      End
      Begin VB.Menu mnuFileDecipher 
         Caption         =   "&Decipher"
      End
      Begin VB.Menu mnuFileBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmSecret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

DefLng A-Z

Private Sub cmdBrowse_Click()
cdlOne.DialogTitle = "Encryption 1"
cdlOne.Flags = cdlOFNHideReadOnly
cdlOne.Filter = "All files (*.*)|*.*|Text files (*.txt)|*.txt"
cdlOne.CancelError = True
On Error Resume Next
cdlOne.ShowOpen
If Err = 0 Then
txtFile.Text = cdlOne.filename
End If
On Error GoTo 0
Dim A$
Dim i, ndx
MousePointer = vbHourglass
Open txtFile.Text For Binary As #1
A$ = Space$(LOF(1))
Get #1, , A$
Close #1
Do
ndx = InStr(A$, Chr$(0))
If ndx = 0 Or ndx > 5000 Then Exit Do
Mid$(A$, ndx, 1) = Chr$(1)
Loop
MousePointer = vbDefault
Me.Text1.Text = A$
Me.Caption = "Encryption 1 - " & txtFile.Text
End Sub

Private Sub cmdDecipher_Click()
frmPassword.Show vbModal
If Password$ = "" Then Exit Sub
If InStr(Head$, Hash$(Password$)) = 9 Then
MousePointer = vbHourglass
cmdEncipher.Enabled = False
cmdDecipher.Enabled = False
cmdBrowse.Enabled = False
mnuFileEncipher.Enabled = False
mnuFileDecipher.Enabled = False
Refresh
Decipher
txtFile_Change
cmdBrowse.Enabled = True
mnuFileEncipher.Enabled = True
MousePointer = vbDefault
Dim A$
Dim i, ndx
MousePointer = vbHourglass
Open txtFile.Text For Binary As #1
A$ = Space$(LOF(1))
Get #1, , A$
Close #1
Do
ndx = InStr(A$, Chr$(0))
If ndx = 0 Or ndx > 5000 Then Exit Do
Mid$(A$, ndx, 1) = Chr$(1)
Loop
MousePointer = vbDefault
Me.Text1.Text = A$
Me.Caption = "Encryption 1 - " & txtFile.Text
Else
MsgBox "Sorry, password incorrect for this file.", 48, "Encryption 1"
End If
End Sub

Private Sub cmdEncipher_Click()
frmPassword.Show vbModal
If Password$ = "" Then Exit Sub
MousePointer = vbHourglass
cmdEncipher.Enabled = False
cmdDecipher.Enabled = False
cmdBrowse.Enabled = False
mnuFileEncipher.Enabled = False
mnuFileDecipher.Enabled = False
Refresh
Encipher
txtFile_Change
cmdBrowse.Enabled = True
mnuFileBrowse.Enabled = True
MousePointer = vbDefault
Dim A$
Dim i, ndx
MousePointer = vbHourglass
Open txtFile.Text For Binary As #1
A$ = Space$(LOF(1))
Get #1, , A$
Close #1
Do
ndx = InStr(A$, Chr$(0))
If ndx = 0 Or ndx > 5000 Then Exit Do
Mid$(A$, ndx, 1) = Chr$(1)
Loop
MousePointer = vbDefault
Me.Text1.Text = A$
Me.Caption = "Encryption 1 - " & txtFile.Text
End Sub

Private Sub Command1_Click()
mnuFileExit_Click
End Sub


Private Sub Command2_Click()
Text1.Text = ""
txtFile.Text = ""
frmSecret.Caption = "Encryption 1"
Load Form2
Form2.Show
End Sub


Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
cmdEncipher.Enabled = False
cmdDecipher.Enabled = False
mnuFileEncipher.Enabled = False
mnuFileDecipher.Enabled = False
txtFile.Text = ""
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Load frmEgg
frmEgg.Show
End Sub

Private Sub mnuAbout_Click()
Load Form1
Form1.Show
End Sub

Private Sub mnuFileBrowse_Click()
cmdBrowse_Click
End Sub

Private Sub mnuFileDecipher_Click()
cmdDecipher_Click
End Sub

Private Sub mnuFileEncipher_Click()
cmdEncipher_Click
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub


Private Sub mnuFileNew_Click()
Command2_Click
End Sub

Private Sub mnuHelp_Click()
Load frmBrowser
frmBrowser.WebBrowser1.Navigate App.Path & "\Help.htm"
frmBrowser.Show
End Sub


Private Sub txtFile_Change()
Dim FileLen
On Error Resume Next
FileLen = Len(Dir$(txtFile.Text))
If Err <> 0 Or FileLen = 0 Or Len(txtFile.Text) = 0 Then
cmdEncipher.Enabled = False
cmdDecipher.Enabled = False
Exit Sub
End If
Open txtFile.Text For Binary As #1
Head$ = Space$(16)
Get #1, , Head$
Close #1
If InStr(Head$, "[Secret]") = 1 Then
cmdEncipher.Enabled = False
cmdDecipher.Enabled = True
mnuFileEncipher.Enabled = False
mnuFileDecipher.Enabled = True
Else
cmdEncipher.Enabled = True
cmdDecipher.Enabled = False
mnuFileEncipher.Enabled = True
mnuFileDecipher.Enabled = False
End If
End Sub

Sub Encipher()
Open txtFile.Text For Binary As #1
Dim A$, H$
A$ = Space$(LOF(1))
Get #1, , A$
H$ = "[Secret]" & Hash$(Password$) & vbCrLf
Process A$
Put #1, 1, H$
Put #1, , A$
Close #1
End Sub

Sub Decipher()
Open txtFile.Text For Binary As #1
Dim A$, H$
H$ = Space$(18)
A$ = Space$(LOF(1) - 18)
Get #1, , H$
Get #1, , A$
Close #1
Process A$
Kill txtFile.Text
Open txtFile.Text For Binary As #1
Put #1, , A$
Close #1
End Sub

Sub Process(A$)
Dim n1, n2, n3
Dim i
For i = 1 To Len(Password$)
n1 = n1 + Asc(Mid$(Password$, i, 1))
n1 = (n1 * 367 + 331) Mod &HFFF
n2 = ((n2 + n1) * 743 + 599) Mod &HFFF
n3 = ((n3 + n2) * 563 + 787) Mod &HFFF
Next i
Cipher A$, n1, n2, n3
End Sub
