VERSION 5.00
Begin VB.Form frmStopMessage 
   Caption         =   "SFDC"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmStopMessage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "frmStopMessage.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblDisableMsg 
      Height          =   975
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmStopMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DisplayPeriod = 5 'seconds
Dim Dummy As Boolean

Private Sub Form_Activate()

    'Set the Start Time and Enable the Timer
    Dummy = OK2DisplayForm(True, DisplayPeriod)
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()

    If Not OK2DisplayForm(False, DisplayPeriod) Then
        Timer1.Enabled = False
        Me.Hide
    End If

End Sub
