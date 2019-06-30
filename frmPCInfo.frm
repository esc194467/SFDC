VERSION 5.00
Begin VB.Form frmPCInfo 
   Caption         =   "PC Information"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboPlant 
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      ToolTipText     =   "Click this button to Cancel"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdRegisterPC 
      Caption         =   "&Register PC"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "After Clicking - Swipe your company ID card through the magnetic stripe reader"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Enter the Location"
      Top             =   2160
      Width           =   3135
   End
   Begin VB.ComboBox cboBuilding 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Select the Building"
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox txtComputerName 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblEnterDetails 
      Caption         =   "Please complete the following information and click Register PC."
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lblLocation 
      Caption         =   "Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblBuilding 
      Caption         =   "Building:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblPlant 
      Caption         =   "Plant:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblComputerName 
      Caption         =   "Computer Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmPCInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

  End
  
End Sub

Private Sub cmdRegisterPC_Click()

  gblClockStatus = "Update"
  frmClocking.MediaPlayer1.Play
  
  ' Check if fields are populated
  If cboPlant.Text = "" Or Null Then
    MsgBox "Invalid entry for Plant" & _
            Chr(13) & Chr(13) _
            & " - Please select a plant from the list" _
            , vbExclamation, gblDialogTitleClocking
  End If
  If cboBuilding.Text = "" Or Null Then
    MsgBox "Invalid entry for Building" & _
            Chr(13) & Chr(13) _
            & " - Please select a Building from the list" _
            , vbExclamation, gblDialogTitleClocking
  End If
  If txtLocation.Text = "" Or Null Then
    MsgBox "Location cannot be blank" & _
            Chr(13) & Chr(13) _
            & " - Please enter a value" _
            , vbExclamation, gblDialogTitleClocking
  End If
  
End Sub

Private Sub Form_Load()

  Dim i As Integer

  txtComputerName = gblInstallation.Name
  txtPlant = gblInstallation.Plant
  cboBuilding = gblInstallation.Building
  txtLocation = gblInstallation.Location
  
  
  ' Add items to Building combo box
  'For i = 0 To UBound(gblPlant)
    'cboPlant.AddItem gblBuilding(i)
  'Next
  
  
End Sub
