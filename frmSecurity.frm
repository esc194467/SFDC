VERSION 5.00
Begin VB.Form frmSecurity 
   Caption         =   "Encrypt Password"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3255
   Icon            =   "frmSecurity.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      Caption         =   "Enter password"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdEncrypt_Click()
  
  On Error GoTo Err_cmdEncrypt_Click
  
  'Encrypt only once (until password is changed)
  'take only first 8 chars as this is SAP max password
  txtPassword = Left(EncryptPassword(txtPassword), 8)
  cmdEncrypt.Enabled = False
  txtPassword.SetFocus
  
Exit_cmdEncrypt_Click:
  Exit Sub
  
Err_cmdEncrypt_Click:
  MsgBox Error$
  Resume Exit_cmdEncrypt_Click
  
End Sub


Private Sub Form_Activate()
  
On Error GoTo Err_Form_Activate
  
  txtPassword.SetFocus
  
Exit_Form_Activate:
  Exit Sub
  
Err_Form_Activate:
  MsgBox Error$
  Resume Exit_Form_Activate
  
End Sub

Private Sub txtPassword_Change()

    cmdEncrypt.Enabled = True
    
End Sub
