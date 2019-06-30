VERSION 5.00
Begin VB.Form frmMsg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2175
   ClientLeft      =   4515
   ClientTop       =   4125
   ClientWidth     =   5085
   ControlBox      =   0   'False
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image ImgPrinter 
      Height          =   495
      Left            =   2160
      Picture         =   "frmMsg.frx":0442
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image imgFileOpen 
      Height          =   495
      Left            =   2160
      Picture         =   "frmMsg.frx":074C
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label labWait 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait..."
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
      Left            =   1320
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label labMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
