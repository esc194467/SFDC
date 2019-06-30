VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form FrmDocuments 
   Caption         =   "Operation Documents"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   LinkTopic       =   "Form2"
   ScaleHeight     =   7200
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   360
      TabIndex        =   10
      Top             =   4680
      Width           =   3495
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2415
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4260
      _Version        =   327680
      Enabled         =   -1  'True
      RightMargin     =   3
      TextRTF         =   $"FrmDocuments.frx":0000
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3840
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   240
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Op Documents"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Op Description"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Work Centre"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Operation No"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Order No"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Dim myfile As String
   myfile = "c:\temp\my_sap.rtf"
   RichTextBox1.LoadFile myfile, rtfRTF
End Sub


