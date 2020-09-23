VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encryption1 Help"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand SSCommand3 
      Height          =   615
      Left            =   7440
      TabIndex        =   3
      Top             =   120
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "frmBrowser.frx":0742
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "frmBrowser.frx":0A5C
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "frmBrowser.frx":0EAE
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8055
      ExtentX         =   14208
      ExtentY         =   5530
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SSCommand1_Click()
On Error GoTo oops
WebBrowser1.GoBack
oops:
Exit Sub
End Sub

Private Sub SSCommand2_Click()
On Error GoTo whoops
WebBrowser1.GoForward
whoops:
Exit Sub
End Sub

Private Sub SSCommand3_Click()
Unload frmBrowser
End Sub
