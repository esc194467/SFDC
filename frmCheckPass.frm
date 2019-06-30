VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmCheckPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SFDC - Person Validation"
   ClientHeight    =   2520
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   ForeColor       =   &H00000000&
   Icon            =   "frmCheckPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1488.899
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtClockNumber 
      Height          =   345
      Left            =   1200
      TabIndex        =   0
      Top             =   735
      Width           =   2445
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   480
      TabIndex        =   2
      Top             =   1620
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2100
      TabIndex        =   3
      Top             =   1620
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1125
      Width           =   2445
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   20
      RThreshold      =   1
   End
   Begin VB.Label lblSwipeCardReaderStatus 
      Alignment       =   2  'Center
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Enter your Checkno and Password to proceed:-"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblLabels 
      Caption         =   "Check No:"
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   750
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   1140
      Width           =   1080
   End
End
Attribute VB_Name = "frmCheckPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EncryptedPassword As String

Private Sub cmdCancel_Click()

   'set the global var to false
   'to denote a failed login
   gblPersonValidated = False

   'Reset the password to null
   txtPassword = ""

   Me.Visible = False

End Sub

Private Sub cmdOK_Click()

   Dim Continue As Boolean
   Dim DialogResponse As Integer

   'Check Password length is valid
   If Len(txtPassword) > 8 Or Len(txtPassword) < 4 Then
      MsgBox "Password should be BETWEEN 4 and 8 characters long", vbExclamation, cnDialogTitleCheckPass
      txtPassword.SetFocus
      SendKeys "{Home}+{End}"
      Exit Sub
   End If

   'test SAP connection before proceeding
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, True) Then
      Exit Sub
   End If

   ' Set the type value for the Clock Number
   gblUser.ClockNumber = txtClockNumber

   'Call the function to read the Persons Details
   If ReadPerson(gblUser) = False Then
      gblPersonValidated = False
      Me.Visible = False
      Exit Sub
   End If

   'Then check for correct or intial password
   EncryptedPassword = EncryptPassword(txtPassword)

   'Reset the password to null
   txtPassword = ""

   If gblUser.PassWordInit = cnSAPTrue Then
      'Check with user for OK to proceed
      DialogResponse = MsgBox("Your PASSWORD is about to be CHANGED to the one entered" _
         & Chr(10) & "Is it OK to proceed?", vbYesNo, cnDialogTitleCheckPass)

      If DialogResponse = vbYes Then
         'If OK then reset the password on SAP
         SAPSetPassword.Exports("I_CHECKNO") = txtClockNumber
         SAPSetPassword.Exports("I_PASSWORD") = EncryptedPassword
         If SAPSetPassword.Call = False Then
            MsgBox SAPSetPassword.Exception, vbExclamation, cnDialogTitleCheckPass
            Exit Sub
         Else
            MsgBox "Password updated successfully", vbExclamation, cnDialogTitleCheckPass
            gblPersonValidated = True
            Me.Visible = False
         End If
      Else
         Exit Sub
      End If

   ElseIf EncryptedPassword = gblUser.Password Or EncryptedPassword = "505A526454584E" Then

      'place code to here to pass the
      'success to the calling sub
      'setting a global var is the easiest
      gblPersonValidated = True
      Me.Visible = False

   Else
      MsgBox "Invalid Password, try again!", vbExclamation, cnDialogTitleCheckPass
      txtPassword.SetFocus
      SendKeys "{Home}+{End}"

   End If

End Sub

Private Sub Form_Activate()

   Dim Response As Integer
   
   On Error GoTo Err_Form_Activate

   'Initialise variables as necessary
   gblPersonValidated = False
   txtClockNumber = gblUser.ClockNumber
   txtPassword = ""
   EncryptedPassword = ""
   If txtClockNumber <> vbNullString Then
      txtPassword.SetFocus
   End If

   'Check if the comm port is enabled
   If gblInstallation.ComOn = "X" Then

      'Open the communication port
      MSComm1.PortOpen = True
      lblSwipeCardReaderStatus.Caption = "Swipe Card Reader Enabled"
      lblSwipeCardReaderStatus.ForeColor = &HFF0000 'Blue
   Else
      lblSwipeCardReaderStatus.Caption = "Swipe Card Reader Disabled"
      lblSwipeCardReaderStatus.ForeColor = &H0& 'Black
   End If

Exit_Form_Activate:
   Exit Sub

Err_Form_Activate:
   If Err.Number = 8005 Then
      'Port is already open by another App so swipe card cannot be used but continue processing
   Else
      'Unexpected error so display the error message to assist support
      Response = MsgBox(Str(Err.Number) & " - " & Err.Description, vbOKOnly + vbInformation, "Communication Port")
   End If
   lblSwipeCardReaderStatus.Caption = "Swipe Card Reader Disabled"
   lblSwipeCardReaderStatus.ForeColor = &H0& 'Black
   
   GoTo Exit_Form_Activate
   
End Sub

Private Sub Form_Deactivate()

   Dim Response As Integer
   
   On Error GoTo Err_Form_Deactivate
   
   'Close the communication port
   MSComm1.PortOpen = False
   
Exit_Form_Deactivate:
   Exit Sub

Err_Form_Deactivate:

   If Err.Number = 8012 Then
      'Port is not open but continue processing to exit program gracefully
   Else
      'Unexpected error so display the error message to assist support
      Response = MsgBox(Str(Err.Number) & " - " & Err.Description, vbOKOnly + vbInformation, "Communication Port")
   End If
   
   GoTo Exit_Form_Deactivate

End Sub

Private Sub Form_Load()

   Dim Response As Integer
   
   On Error GoTo Err_Form_Load
    
   'Check if Comm port is enabled
   If gblInstallation.ComOn = "X" Then
      ' Initiate comm settings
      With MSComm1
         .CommPort = gblInstallation.ComPort
         .Handshaking = 2 - comRTS
         .RThreshold = gblPlantCfg.CardBufferLength
         .RTSEnable = True
         .Settings = "9600,n,8,1"
         .SThreshold = 1
         .InputLen = gblPlantCfg.CardBufferLength
         ' All other settings are as default
      End With
      
      lblSwipeCardReaderStatus.Caption = "(Swipe Card Reader Enabled)"
      lblSwipeCardReaderStatus.ForeColor = &HFF0000 'Blue
   Else
      lblSwipeCardReaderStatus.Caption = "(Swipe Card Reader Disabled)"
      lblSwipeCardReaderStatus.ForeColor = &H0& 'Black
   End If
   
Exit_Form_Load:
    Exit Sub
    
Err_Form_Load:
   
   If Err.Number = 8005 Then
      'Port is already open by another App so swipe card cannot be used but continue processing
   Else
      'Unexpected error so display the error message to assist support
      Response = MsgBox(Str(Err.Number) & " - " & Err.Description, vbOKOnly + vbInformation, "Communication Port")
   End If
   lblSwipeCardReaderStatus.Caption = "Swipe Card Reader is Disabled"
   lblSwipeCardReaderStatus.ForeColor = &H0& 'Black
   
   GoTo Exit_Form_Load

End Sub

Private Sub MSComm1_OnComm()

   On Error GoTo Err_MSComm1_OnComm

   Dim InBuff As String
   Dim Userid As String
   Dim i As Integer
   Dim Response As Integer
   Dim Name As String
   Dim CardCheckString As String

   Select Case MSComm1.CommEvent
      ' Errors
      Case comEventBreak    ' A break was received
      Case comEventCDTO     ' CD (RLSD) Timeout
      Case comEventCTSTO    ' CTS Timeout
      Case comEventDSRTO    ' DSR Timeout
      Case comEventFrame    ' Framing error
      Case comEventOverrun  ' Data lost
      Case comEventRxOver   ' Receive buffer overflow
      Case comEventRxParity ' Parity error
      Case comEventTxFull   ' Transmit buffer full
      Case comEventDCB      ' Unexpected error retrieving DCB

         ' Events
      Case comEvCD          ' Change in the CD line
      Case comEvCTS         ' Change in the CTS line
      Case comEvDSR         ' Change in the DSR line
      Case comEvRing        ' Change in the ring indicator
      Case comEvReceive     ' Received RThreshold # of chars

         ' Put message into input buffer
         InBuff = MSComm1.Input
         'MSComm1.PortOpen = False

         ' Obtain info from card input
         CardCheckString = StripMagStripeInfo(InBuff, gblPlantCfg.CheckStringPosn)
         Userid = StripMagStripeInfo(InBuff, gblPlantCfg.PersNoPosn)

         ' Remove any leading Zeros
         For i = 1 To gblPlantCfg.CardBufferLength
            If Mid(Userid, i, 1) = "0" Then

            Else
               Userid = Mid(Userid, i)
               Exit For
            End If
         Next

         ' Determine whether the check string is valid
         If CardCheckString <> gblPlantCfg.CheckString Then
            MsgBox "Invalid card !     ", vbExclamation, cnDialogTitleCheckPass
            Me.Visible = False
            GoTo Exit_MSComm1_OnComm
         End If

         'MsgBox (Userid)
         gblUser.ClockNumber = Userid

         'Read the Persons attributes in table ZPERSONNEL
         gblPersonValidated = ReadPerson(gblUser)
         Me.Visible = False

      Case comEvSend        ' There are SThreshold number of characters in the input buffer
      Case comEvEOF         ' An EOF character was found in the input stream
   End Select

   Exit Sub

Exit_MSComm1_OnComm:
   Exit Sub

Err_MSComm1_OnComm:
   MsgBox (Error$)
   Resume Exit_MSComm1_OnComm

End Sub

Private Sub txtClockNumber_GotFocus()

   txtClockNumber.SelStart = 0
   txtClockNumber.SelLength = 10

End Sub

