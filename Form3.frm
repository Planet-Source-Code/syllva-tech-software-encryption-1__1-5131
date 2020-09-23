VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form3 
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   2415
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame3 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   "Languages Available"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Begin VB.Label Label12 
         Caption         =   "English"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Basque"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Spanish"
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Italian"
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "German"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "French"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _Version        =   65536
      _ExtentX        =   9551
      _ExtentY        =   1508
      _StockProps     =   15
      Caption         =   "Encryption1 for Windows 95/98"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   2
      Font3D          =   3
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 2.50"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 1998 by Millenium Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   5295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Load frmSecret
frmSecret.Show
Unload Form3
End Sub

Private Sub Image1_Click()
Load frmEgg
frmEgg.Show
End Sub
