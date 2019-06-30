VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.0#0"; "SHDOCVW.DLL"
Begin VB.Form frmWebBrowser 
   Caption         =   "RRAES Web Browser"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7335
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   11415
      Object.Height          =   489
      Object.Width           =   761
      AutoSize        =   0
      ViewMode        =   1
      AutoSizePercentage=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   615
      Left            =   10800
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox URLAddress 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   8535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   570
   End
End
Attribute VB_Name = "frmWebBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If URLAddress = StartURLAddress Then
        frmWebBrowser.Hide
    Else
        WebBrowser1.GoBack
    End If
End Sub



Private Sub WebBrowser1_NavigateComplete(ByVal URL As String)
    URLAddress = WebBrowser1.LocationURL
End Sub
