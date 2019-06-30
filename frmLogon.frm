VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogon 
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7455
   DrawStyle       =   5  'Transparent
   Icon            =   "frmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmLogon.frx":0442
   ScaleHeight     =   7200
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   3360
      Top             =   6000
   End
   Begin VB.ListBox lstSysMessages 
      BackColor       =   &H0000FFFF&
      Height          =   645
      Left            =   360
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.CommandButton cmdZeroBooking 
      Caption         =   "&Zero Time Booking"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   16
      Top             =   6000
      Width           =   2895
   End
   Begin VB.TextBox OrderNumber 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox OrderDesc 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3120
      Width           =   4215
   End
   Begin VB.CommandButton cmdMultiBookings 
      Caption         =   "&Multi Bookings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2880
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdViewPrevBookings 
      Caption         =   "&View Previous Bookings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   11
      Top             =   6000
      Width           =   3015
   End
   Begin VB.CommandButton cmdWorkBook 
      Caption         =   "&Work Booking"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox OpDescription 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3840
      Width           =   4575
   End
   Begin VB.TextBox WorkCentre 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox OpNumber 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3840
      Width           =   615
   End
   Begin MSMask.MaskEdBox ConfirmationNumber 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000000000"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdShowOpinfo 
      Caption         =   "&Op Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4800
      TabIndex        =   2
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   4920
      Picture         =   "frmLogon.frx":56B6C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Description"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   2880
      Width           =   1230
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Number"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Op Description"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2400
      TabIndex        =   9
      Top             =   3600
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Work Centre"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Op No"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   3600
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmation Number"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1470
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu HelpFile 
         Caption         =   "Open Help File"
         Shortcut        =   {F1}
      End
      Begin VB.Menu HelpNull1 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About Shop Floor Data"
      End
   End
   Begin VB.Menu mnuPCInfo 
      Caption         =   "&PC Info"
      Begin VB.Menu mnuComputerName 
         Caption         =   "Computer Name"
      End
      Begin VB.Menu mnuPlant 
         Caption         =   "Plant"
      End
      Begin VB.Menu mnuBuilding 
         Caption         =   "Building"
      End
      Begin VB.Menu mnuLocation 
         Caption         =   "Location"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMessage1 
         Caption         =   "message1"
      End
      Begin VB.Menu mnuMessage2 
         Caption         =   "message2"
      End
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Dummy As Boolean

Private Sub cmdJumpTo_Click()

   'test SAP connection
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
      Exit Sub
   End If

   Select Case gblPlantCfg.BookingMode

      Case cnBookOnBookOff
         'This function removed for ver 3.1
         'DoJumpToOP

      Case cnPostEventBooking
         'Zero Time Booking
         DoZeroTimeBooking

   End Select

End Sub

Private Sub cmdMultiBookings_Click()
PCNF = cnSAPTrue ' -- TCR7117 --
   'test SAP connection
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
      Exit Sub
   End If

   Select Case gblPlantCfg.BookingMode

      Case cnBookOnBookOff
         'Zero Time Booking
         DoZeroTimeBooking

      Case cnPostEventBooking
         'MultiBooking
         DoMultiBooking

   End Select

End Sub

Private Sub cmdShowOpinfo_Click()

   'test SAP connection
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
      Exit Sub
   End If

   DoShowOpInfo

End Sub

Private Sub cmdViewPrevBookings_Click()

   'test SAP connection
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
      Exit Sub
   End If

   'then do the function
   DoViewPrevBookings

End Sub

Private Sub cmdWorkBook_Click()
PCNF = cnSAPTrue ' -- TCR7117 --
   'first test SAP connection
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
      Exit Sub
   End If

   Select Case gblPlantCfg.BookingMode
      'test for the Booking Mode and open the appropriate form

      Case cnBookOnBookOff  'BOBO Mode
         DoBOBO

      Case cnPostEventBooking 'POST EVENT BOOKING Mode
         DoWorkBook

   End Select

End Sub

Private Sub cmdZeroBooking_Click()

   DoZeroTimeBooking

End Sub

Private Sub ConfirmationNumber_Change()

   ' Unload the form as the information may change
   Unload frmOpInfo

End Sub

Private Sub ConfirmationNumber_GotFocus()

   'select existing number for easy change
   ConfirmationNumber.SelStart = 0
   ConfirmationNumber.SelLength = 10

End Sub

Private Sub ConfirmationNumber_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      ConfirmationNumber_LostFocus
   End If

End Sub

Private Sub ConfirmationNumber_LostFocus()

   'if a conf number has been entered call SAP function to validate it
   If ConfirmationNumber <> "" Then
      ' Read the operation via the confirmation number
      gblOperation.ConfNo = ConfirmationNumber
      If ReadOrderOp(gblOrder, gblOperation, gblOrdTypeCfg) = True Then

         'Populate the Form Fields
         With frmLogon
            .OrderNumber = Format(gblOrder.Number, "##########")
            .OrderDesc = gblOrder.Desc
            .OpNumber = gblOperation.Number
            .WorkCentre = gblOperation.WorkCentre
            .OpDescription = gblOperation.Desc
         End With

         'Set the global to indicate successful read of Operation
         gblOperationWasFound = True

      Else
         gblOperationWasFound = False
      End If
   Else
      ResetLogonForm
   End If

End Sub

Private Sub Form_Activate()

   ' Set Locale date format
   Call SetLocalDateFormat

   ReadSysMessages

End Sub

Private Sub Form_Load()

   Dim Host As String

   'Set the SFDC Version Number global variable
   gblSFDCVersion = App.Major & "." & App.Minor & "." & App.Revision
   
   'Determine the HostName AKA Computer Name or Tag Number
   gblInstallation.Name = RTrim(GetHost())

   'Get SAP Logon details etc.. from server-side .ini file
   GetSAPparameters
   
   'Check for invalid CodePage if requested by server-side ini file
   If gblSAP.CodePageCheck = "Y" Then
    If IsCodePageOK() = False Then
     Unload Me
     Exit Sub
    End If
   End If
   
  
   'Determine the LAN username
   gblLANUser = GetThreadUserName()
   
   'Determine the IP address by the host name
   If SocketsInitialize() Then
  
    'obtain and pass the host address to the function
    Host = GetMachineName()
    gblIPAddress = GetIPFromHostName(Host)
     
    SocketsCleanup
     
   Else
        MsgBox "Windows Sockets for 32 bit Windows " & _
               "environments is not successfully responding."
        Unload Me
        Exit Sub
   End If

   'Create Logon and Connection Objects
   Set oLogonControl = CreateObject("SAP.LogonControl.1")
   Set oConnection = oLogonControl.NewConnection

   'Set the Connection parameters based on Load Balancing
   If gblSAP.LoadBalancing = cnYes Then
      ' Load balancing system info
      oConnection.MessageServer = gblSAP.Server
      oConnection.GroupName = gblSAP.GroupName
   Else
      ' Direct access system info
      oConnection.ApplicationServer = gblSAP.Server
   End If

   'Set remaining Connection parameters
   With oConnection
      .System = gblSAP.System
      .SystemNumber = gblSAP.SystemNumber
      .RFCwithDialog = 0
      .TraceLevel = gblSAP.TraceLevel
      .Client = gblSAP.Client
      .User = gblSAP.Userid
      .password = gblSAP.password
      .Language = gblSAP.Language
   End With

   'Attempt to make the connection to SAP
   If Connect2SAP() = False Then
      'Do not proceed as there is a problem connecting
      Unload Me
      Exit Sub
   End If

   'Create the SAP Function Objects
   Call CreateSAPFunctionObjects

   'Test SAP Connection and get/set Installation Details
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckStartup, False) Then
      Unload Me
      Exit Sub
   End If

   'Read the SFDC Config Tables
   If SAPReadConfig.Call = False Then
      MsgBox SAPReadConfig.Exception, vbExclamation, cnDialogTitleLogon
      Unload Me
      Exit Sub
   End If

   'Read the appropriate rows in the Config Tables
   gblPlantCfg.Plant = gblInstallation.Plant
   'call the function
   If SetPlantConfig(gblPlantCfg) = False Then
      MsgBox "PLANT NOT FOUND IN CONFIG TABLES", vbExclamation, cnDialogTitleLogon
      Unload Me
      Exit Sub
   End If

   'Check if multi instances are allowed
   If gblInstallation.SFDCMulti <> "X" Then
      'Check if SFDC is already running on this PC
      If App.PrevInstance Then
         With frmStopMessage
            .lblDisableMsg = "SFDC IS ALREADY ACTIVE ON THIS PC" & Chr(10) & _
               "PLEASE SWITCH TO THE ACTIVE SESSION"
            .Show vbModal
         End With
         End
      End If
   End If
   
   'Read the Method tables and load the Method array
   Call SetMethods

   'Read the Variance Reasons for the Plant
   SAPReadReasons.Exports("PLANT") = gblInstallation.Plant
   If SAPReadReasons.Call = False Then
      MsgBox SAPReadConfig.Exception, vbExclamation, cnDialogTitleLogon
      Unload Me
      Exit Sub
   End If

   'Configure Form and Button Captions based on Booking Mode
   Caption = "Shop Floor Data Collection Ver " & App.Major & "." & _
      App.Minor & "." & App.Revision & " (Connected to " & gblSAP.System & _
      " Client " & gblSAP.Client & ")"

   Select Case gblPlantCfg.BookingMode
      Case cnBookOnBookOff
         With frmLogon
            .cmdWorkBook.Caption = "&Book On/Off"
            .cmdMultiBookings.Visible = False
         End With
      Case cnPostEventBooking
         'Do nothing these are the standard design settings
      Case Else
         MsgBox "INVALID BOOKING MODE FOR PLANT", vbExclamation, cnDialogTitleLogon
         Unload Me
         Exit Sub
   End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Dim RetVal As Long
   Dim FrmCnt As Integer

   On Error GoTo ErrorHandler

   'test SAP connection and, if connected, update ZVBI and logoff SAP
   If oConnection.IsConnected = tloRfcConnected Then
   
      Call SAPisConnected(gblSAP, gblInstallation, cnCheckCloseDown, False)

       oConnection.Logoff
    
       'Loop thru the forms collection and unload them
       For FrmCnt = 0 To Forms.Count - 1
          Unload Forms(FrmCnt)
       Next
   End If

   Exit Sub

ErrorHandler:
   End

End Sub

Private Sub About_Click()

   frmAbout.Show vbModal

End Sub

Private Sub HelpFile_Click()

   On Error Resume Next

   Dim Y As Integer
   Dim ErrMsg As String

   'Execute application and file
   Y = ShellExecute(0, "open", gblPlantCfg.HelpFilePath, 0, "", 5)

   If Y > 32 Then
      ' open operation successful
   Else
      ' open operation unsuccesful
      ErrMsg = ShellExErrMsg(Y)

      MsgBox "Unable to open the file - " & "'" _
         & gblPlantCfg.HelpFilePath & "'" & Chr(13) & Chr(13) _
         & ErrMsg & Chr(13) & Chr(13) _
         & "Please contact local support" _
         , vbExclamation, cnDialogTitleOpInfo

   End If
End Sub

Public Sub DoWorkBook()

   Dim TimeDiff As Double

   'set the function
   gblFunction = cnWorkBook

   'Test that the Confirmation Number Operation was found
   If Not ConfNoWasOK Then
      Exit Sub
   End If

   'Check Op can be confirmed via Status checks on Order and Op
   If Not OpisOK2Confirm(gblOrdTypeCfg, gblOperation) Then
      Exit Sub
   End If

   'Determine difference between time now and last booking
   If Val(gblLastBookingTime) = 0 Then gblLastBookingTime = Now - 0.1
   TimeDiff = Now - gblLastBookingTime

   'Check for continuous bookings and lapsed time since last booking
   If gblContinueBooking Then
      If TimeDiff > (gblPlantCfg.BookingTimeOut / 60 / 60 / 24) Then
         gblContinueBooking = False
      End If
   End If

   'Then if necessary display the checkno/password dialog
   If gblContinueBooking Then
      gblPersonValidated = True
      gblContinueBooking = False
   Else
      frmCheckPass.Show vbModal
   End If

   'If successful password entry open Work Book form
   If gblPersonValidated Then
      'Reset the ContinueBooking global and open the Work Booking form
      gblContinueBooking = False
     'Check the Person is authorised to Work Book via SFDC
      If CanWorkBook() = False Then
          Exit Sub
      End If
      frmWorkBook.Show vbModal
   End If

   'Disable the frmWorkBook Timer
   frmWorkBook.Timer1.Enabled = False

End Sub

Public Sub DoBOBO()

   'set the function
   gblFunction = cnBookOnBookOff

   'Get user to identify him/herself
   frmCheckPass.Show vbModal

   'Test Person validated
   If Not gblPersonValidated Then
      Exit Sub
   End If
   
   'Check the Person is authorised to Work Book via SFDC
   If CanWorkBook() = False Then
    Exit Sub
   End If
   
   'Test CICO Status
   TestCICOStatus

   'Show the Book On/Off form
   frmBoBo.Show vbModal

   'Disable the frmBoBo timer
   frmBoBo.Timer1.Enabled = False

   'Unlock the appropriate rows in ZBOBO
   SAPUnlockBoBo.Exports("PERSON") = gblUser.ClockNumber

   'Call the RFC
   If SAPUnlockBoBo.Call = False Then
      MsgBox SAPReadBoBo.Exception, vbExclamation, cnDialogTitleLogon
   End If

End Sub

Public Sub DoZeroTimeBooking()

   Dim Conf As ConfInfo

   'set the function
   gblFunction = cnZeroTimeConfirm

   'Test that the Confirmation Number Operation was found
   If Not ConfNoWasOK Then
      Exit Sub
   End If

   'Check Op can be confirmed via Status checks on Order and Op
   If OpisOK2Confirm(gblOrdTypeCfg, gblOperation) = False Then
      Exit Sub
   End If

   'Then display the checkno/password dialog
   frmCheckPass.Show vbModal

   'Check the Person validated OK
   If Not gblPersonValidated Then
      Exit Sub
   End If
   
   'Check the Person is authorised to Work Book via SFDC
   If CanWorkBook() = False Then
    Exit Sub
   End If

   'Test for Op Plant not the same as Person Plant
   'and warn user
   ' 1/11/99 change to error - requested by Andy Dickinson
   If gblUser.Plant <> gblOperation.Plant Then
      MsgBox "PLANNED WORKCENTRE IS NOT IN YOUR PLANT", _
         vbExclamation, cnDialogTitleWorkBook
      Exit Sub
   End If
    
   'Initialise internal serial number table in case yield is to be recorded on a PP activity
   IntPPSerNoLocnTable.FreeTable

   'Display batch tracking form
   frmBatchTrack.Show vbModal
   'Ensure timer in Batch Track form is disabled on return
   frmBatchTrack.Timer1.Enabled = False

   Me.Refresh

   'Test BatchTrack form was exited correctly
   If Not gblBatchTrackOK Then
      Exit Sub
   End If

   'Set Backflush indicator based on Pre-Op CheckBox
   If frmBatchTrack.ChkPreOp.Value = cnChecked Then
      Conf.BackFlush = cnSAPFalse
   Else
      Conf.BackFlush = cnSAPTrue
   End If

   'Check for Completed Activity/Operation
   If frmBatchTrack.ChkComplete.Value = cnChecked Then
      Conf.Complete = cnSAPTrue
      Conf.FinalConf = cnSAPTrue
   Else
      Conf.Complete = cnSAPFalse
      Conf.FinalConf = cnSAPFalse
   End If

   'Set the remaining values to make a zero time Final Confirmation
   With Conf
      .ActType = gblOperation.ActivityType
      .ActWork = "0"
      .ConfText = frmBatchTrack.ConfirmationText
      .EndDate = gblSAP.SAPFormatDate
      .EndTime = gblSAP.SAPFormatTime
      .Plant = gblOperation.Plant
      .PostDate = gblSAP.SAPFormatDate
      .WorkCentre = gblOperation.WorkCentre
      .WorkUnits = gblOperation.WorkUnits
      .DevReason = frmBatchTrack.cmboReasons
      .Yield = frmBatchTrack.txtYieldQty
      .Scrap = frmBatchTrack.txtScrapQty
      .Conf_no = frmLogon.ConfirmationNumber.Text
   End With
   'call the function to make the booking
   If MakeConfirmation(Conf, gblOrder, gblOperation) = False Then
      Exit Sub
   Else
      MsgBox "ZERO TIME CONFIRMATION SUCCESSFUL", vbExclamation
      Exit Sub
   End If

End Sub

Public Sub DoMultiBooking()

   'set the function
   gblFunction = cnMultiBook

   'Display the checkno/password dialog
   frmCheckPass.Show vbModal

   'Check the Person Validated
   If Not gblPersonValidated Then
      Exit Sub
   End If
      
   'Check the Person is authorised to Work Book via SFDC
   If CanWorkBook() = False Then
    Exit Sub
   End If

   'Check the Person is authorised for Multi Booking
   If gblUser.CanMultiBook <> cnSAPTrue Then
      MsgBox "YOU ARE NOT AUTHORISED TO USE MULTI BOOKINGS", vbExclamation, cnDialogTitleWorkBook
      Exit Sub
   End If
   
   'Initialise the Flexgrid
   frmMultiBooking.flxMultiBookings.Rows = 1
   frmMultiBooking.flxMultiBookings.Row = 0
   frmMultiBooking.flxMultiBookings.RowSel = 0
   
   'Open the work booking form
   frmMultiBooking.Show vbModal

   'Disable frmMultiBooking timer
   frmMultiBooking.Timer1.Enabled = False

End Sub

Public Sub DoJumpToOP()

   Dim DialogResponse As Integer

   'Test that the Confirmation Number Operation was found
   If Not ConfNoWasOK Then
      Exit Sub
   End If

   'Test the Order Type - as only ZRPR orders can be "jumped"
   If gblOrder.Type <> cnRepairOrderType Then
      MsgBox "ONLY REPAIR ORDER OPERATIONS MAY BE 'JUMPED TO'", vbExclamation, cnDialogTitleLogon
      Exit Sub
   End If

   'Determine if "jump" is for Pre-Opping and set Backflush indicator accordingly
   DialogResponse = MsgBox("IS THIS A 'JUMP TO' A PRE-OP?", _
      (vbYesNoCancel + vbDefaultButton2), cnDialogTitleLogon)

   Select Case DialogResponse
      Case vbYes
         SAPJumpToOP.Exports("BACKFLUSH") = cnSAPFalse
      Case vbNo
         SAPJumpToOP.Exports("BACKFLUSH") = cnSAPTrue
      Case Else
         Exit Sub
   End Select

   'Set other function export parameters
   SAPJumpToOP.Exports("ORDER_NO") = gblOrder.Number
   SAPJumpToOP.Exports("WIP_WORKCTR") = gblOperation.Arbid
   SAPJumpToOP.Exports("WIP_OPNO") = gblOperation.Number

   ' Call SAP function to "jump" order to new op
   If SAPJumpToOP.Call = False Then
      MsgBox SAPJumpToOP.Exception, vbExclamation, cnDialogTitleLogon
   Else
      MsgBox "ORDER MOVED SUCCESSFULLY", vbExclamation, cnDialogTitleLogon
   End If

   'Set logon screen fields to null ready for new Operation
   ResetLogonForm
   frmLogon.ConfirmationNumber.SetFocus

End Sub

Public Sub DoViewPrevBookings()

   'Then display the checkno/password dialog
   frmCheckPass.Show vbModal

   If gblPersonValidated Then
   
      'Check the Person is authorised to Work Book via SFDC
      If CanWorkBook() = False Then
        Exit Sub
      End If

      'open the work booking form
      frmPrevBookings.Show vbModal
   End If

   'Disable frmPrevBookings timer
   frmPrevBookings.Timer1.Enabled = False

End Sub

Public Sub DoShowOpInfo()

   Dim RowCount As Integer
   Dim i As Integer
   Dim LongText As String
   Dim DocListHeader As String
   Dim ToolListHeader As String
   Dim CompListHeader As String
   Dim DocName, DocVersion As String
   Dim tmpDoc As DocumentInfo
   Dim PRTType As String * 1
   Dim PRTNumber As String * 18
   Dim PRTDesc As String * 50
   Dim PRTQty As String * 8
   Dim PRTUnits As String * 3
   Dim PRTString As String * 79

   'Unload form in case it was already open
   Unload frmOpInfo

   'Test that the Confirmation Number Operation was found
   If Not ConfNoWasOK Then
      Exit Sub
   End If

   'Set Form field values
   With frmOpInfo
      .OrderNumber = Format(gblOrder.Number, "############")
      .OpNumber = gblOperation.Number
      .WorkCentre = gblOperation.WorkCentre
      .txtSalesOrderNumber = Format(gblOrder.SalesOrder, "##########")
      .txtConfNo = gblOperation.ConfNo
   End With

   'Clear longtext control
   frmOpInfo.OpLongText1.Text = ""
   frmOpInfo.OpLongText1.Refresh

   'clear the Command Line Doc Refs table
   CmdLineDocRefs.FreeTable

   'Set pointer to hourglass during SAP processing
   Screen.MousePointer = vbHourglass

   'Test for existence of operation longtext
   If gblOperation.LongTextExists <> cnSAPFalse Then

      'If longtext exists proceed to read it
      'Set export parameters for SAPReadOpText

      SAPReadOpText.Exports("I_AUFPL") = gblOrder.Aufpl
      SAPReadOpText.Exports("I_APLZL") = gblOperation.Aplzl
      SAPReadOpText.Exports("I_AUFNR") = gblOrder.Number
      SAPReadOpText.Exports("ORDER_CATEGORY") = gblOrder.Category
      OpTextTable.FreeTable ' Clear the Op Text Table

      If SAPReadOpText.Call = False Then
         'set pointer to standard
         Screen.MousePointer = vbDefault
         MsgBox SAPReadOpText.Exception, vbExclamation, cnDialogTitleLogon
         'Exit Sub
      Else
         RowCount = OpTextTable.RowCount
         For i = 1 To RowCount
            LongText = LongText + OpTextTable.Value(i, "TDFORMAT") + OpTextTable.Value(i, "TDLINE")
         Next i

         frmOpInfo.OpLongText1.TextRTF = LongText

      End If

   End If

   'Set pointer to hourglass during SAP processing
   Screen.MousePointer = vbHourglass

   'Set export parameters fo SAPReadOpPRTs
   SAPReadOpPRTs.Exports("I_AUFPL") = gblOrder.Aufpl
   SAPReadOpPRTs.Exports("I_APLZL") = gblOperation.Aplzl

   DocumentsTable.FreeTable 'remove existing data in table
   PRTsTable.FreeTable 'remove existing data in table
   Call ClearFlexGrid(frmOpInfo.GrdDocumentList) 'clear the documents flex grid

   'Set flexi grid column headers
   DocListHeader = "<Document Description" + Space(100) + "|^Issue" + Space(12)
   ToolListHeader = "<Tool Number" + Space(50) + "|<Tool Description" + Space(80) + "|^ Unit " + "|^  Quantity  "
   CompListHeader = "<Component Number" + Space(39) + "|<Component Description" + Space(69) + "|^ Unit " + "|^  Quantity  "

   frmOpInfo.GrdDocumentList.FormatString = DocListHeader
   frmOpInfo.GrdToolList.FormatString = ToolListHeader
   frmOpInfo.GrdComponentList.FormatString = CompListHeader

   'Call the RFC
   If SAPReadOpPRTs.Call = False Then
      'set pointer to standard
      Screen.MousePointer = vbDefault
      MsgBox SAPReadOpPRTs.Exception, vbExclamation, cnDialogTitleLogon
      'Exit Sub
   End If

   RowCount = DocumentsTable.RowCount

   For i = 1 To RowCount
      'Reading documents
      With tmpDoc
         .Name = DocumentsTable.Value(i, "DKTXT")
         .DocType = DocumentsTable.Value(i, "DOKAR")
         .FileName = DocumentsTable.Value(i, "FILEP")
         .ApplnType = DocumentsTable.Value(i, "DAPPL")
         .AppPath = DocumentsTable.Value(i, "APPFD")
         .PrefixPath = DocumentsTable.Value(i, "PRFXP")
         .RegPath = DocumentsTable.Value(i, "ZREGFD")
         .FullFilePath = tmpDoc.PrefixPath + tmpDoc.FileName
      End With
      ' If appln type is RDM then indicate version number corresponds to the master
      If DocumentsTable.Value(i, "DAPPL") = "RDM" Then
         tmpDoc.RevDate = DocumentsTable.Value(i, "DOKVR") + "(M)"
      Else
         tmpDoc.RevDate = DocumentsTable.Value(i, "DOKVR")
      End If

      'Add the item to the doc list (filepath already created above)
      Call AddItemtoDocList(tmpDoc)

   Next i

   'Read the Packages and add them as additional items in the Document List
   'but only if the Corena Manual Id is present and Corena links are requested
   'nb. the Packages have been read via the Op Confirmation Number query to SAP
   If gblOrder.CorenaManId > vbNullString And gblPlantCfg.CorenaLinks = cnSAPTrue Then
      'Set the default values for the documents
      With tmpDoc
         .DocType = "EM"
         .ApplnType = "FRS"
         .AppPath = "%auto%"
         .PrefixPath = gblPlantCfg.URLRoot
         .RevDate = "Latest"
      End With

      RowCount = PackagesTable.RowCount
      For i = 1 To RowCount
         If PackagesTable.Value(i, "KTEX2") > vbNullString Then
            With tmpDoc
               .Name = "Scheme " + PackagesTable.Value(i, "KTEX2")
               .FileName = PackagesTable.Value(i, "KTEX2")
               .FullFilePath = ""
            End With

            'Create the URL and add the item to the doc list
            Call AddItemtoDocList(tmpDoc)

         End If
      Next i

      'Read any Command Line Doc References from the Op long Text
      'but again only if the Corena Manual Id has been established
      'nb. the op long text was analysed in SAP to determine the command line doc refs
      RowCount = CmdLineDocRefs.RowCount
      For i = 1 To RowCount
         If CmdLineDocRefs.Value(i, "SUBTASKID") > "" Then
            tmpDoc.FileName = CmdLineDocRefs.Value(i, "TASKID") + _
               "&SubTaskId=" + _
               CmdLineDocRefs.Value(i, "SUBTASKID")
         Else
            tmpDoc.FileName = CmdLineDocRefs.Value(i, "TASKID")
         End If
         With tmpDoc
            .DocType = CmdLineDocRefs.Value(i, "DOCTYPE")
            .Name = tmpDoc.DocType + " TaskId=" + tmpDoc.FileName
            .FullFilePath = ""
         End With

         'Create the URL and add the item to the doc list
         Call AddItemtoDocList(tmpDoc)

      Next i

      'Create the reference to the Corena IPC for this Location
      'but only for ZRPR orders
      If gblOrder.Type = "ZRPR" Then

         With tmpDoc
            .DocType = "EIPC"
            .RevDate = "Latest"
            .FileName = gblOrder.ATALocn
            .Name = "IPC Fig " + tmpDoc.FileName
            .FullFilePath = ""
         End With

         'Create the URL and add the item to the doc list
         Call AddItemtoDocList(tmpDoc)

      End If

   End If

   'Create the reference to the WorkScope Document if present
   If gblOrder.WorkScopeFileName > vbNullString Then

      With tmpDoc
         .DocType = "WS"
         .RevDate = "Latest"
         .FileName = gblOrder.WorkScopeFileName
         .Name = "Workscope Document " + tmpDoc.FileName
         .FullFilePath = gblPlantCfg.SOPath & _
            Format(gblOrder.SalesOrder, "##########") & _
            "\" & tmpDoc.FileName
      End With

      'Create the URL and add the item to the doc list
      Call AddItemtoDocList(tmpDoc)

   End If

   '***********************PRTs****************************
   '*******************************************************

   'Read the PRTs and present them in the flexgrid
   RowCount = PRTsTable.RowCount

   For i = 1 To RowCount

      PRTType = PRTsTable.Value(i, "PRT_CAT")
      PRTNumber = PRTsTable.Value(i, "TOOLNO")
      PRTDesc = PRTsTable.Value(i, "TOOLDESC")
      PRTQty = PRTsTable.Value(i, "QUANTITY")
      PRTUnits = PRTsTable.Value(i, "QTY_UNITS")

      Select Case PRTType
         Case "C"
            frmOpInfo.GrdComponentList.AddItem (PRTNumber + Chr(9) + PRTDesc + Chr(9) + _
               PRTUnits + Chr(9) + PRTQty)
         Case Else
            frmOpInfo.GrdToolList.AddItem (PRTNumber + Chr(9) + PRTDesc + Chr(9) + _
               PRTUnits + Chr(9) + PRTQty)

      End Select
   Next i

   'set pointer to standard
   Screen.MousePointer = vbDefault

   frmOpInfo.Show

End Sub



Private Sub mnuPCInfo_Click()

   mnuComputerName.Caption = "Computer Name: " & gblInstallation.Name
   mnuPlant.Caption = "Plant: " & gblInstallation.Plant
   mnuBuilding.Caption = "Building: " & gblInstallation.Building
   mnuLocation.Caption = "Location: " & gblInstallation.Location
   mnuMessage1.Caption = "If these details are incorrect contact"
   mnuMessage2.Caption = "your Local IT Support Team on " & gblPlantCfg.ContactPhone

End Sub

Private Sub Timer1_Timer()

   On Error GoTo ErrorHandler

   'MsgBox "frmLogon Timer Check"

   If Screen.ActiveForm.Name = Me.Name Then
      'Test SAP Connection and get/set Installation Details
      'and stop the application if necessary
      'MsgBox "Check SAP Connection"
      If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
         Unload Me
         Exit Sub
      End If
   End If

   Exit Sub

ErrorHandler:
   'MsgBox "Error Trap"
   Exit Sub

End Sub

Private Sub AddItemtoDocList(Doc As DocumentInfo)
   
   'Create the path (URL)
   If Doc.FullFilePath = "" Then
      Doc.FullFilePath = Doc.PrefixPath _
         + "EngineType=" + gblOrder.CorenaManId _
         + "&RevDate=" + Doc.RevDate _
         + "&Variant=" + gblOrder.EngineType + gblOrder.EngMark + gblOrder.EngVar _
         + "&PubType=" + Doc.DocType _
         + "&DocId=" + Doc.FileName
   End If

   'Add the row to the document list
   frmOpInfo.GrdDocumentList.AddItem _
      (Doc.Name + Chr(9) + _
      Doc.RevDate + Chr(9) + _
      Doc.ApplnType + Chr(9) + _
      Doc.DocType + Chr(9) + _
      Doc.AppPath + Chr(9) + _
      Doc.FullFilePath)
      

End Sub

