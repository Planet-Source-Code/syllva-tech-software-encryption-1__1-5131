VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Encryption1"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Encryption1 For Windows 95/98"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   2
      Font3D          =   3
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0742
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Copyright 1998 by Millenium Software"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "Form1.frx":082D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub
