VERSION 5.00
Begin VB.Form frmEgg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "You found me!"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ClipControls    =   0   'False
   Icon            =   "frmEgg.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "This Program was written by Sean Young. 1997."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Always keep your head in the clouds, but your feet on the ground."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   120
      Picture         =   "frmEgg.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4440
   End
End
Attribute VB_Name = "frmEgg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

