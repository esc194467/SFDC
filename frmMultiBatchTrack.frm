VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMultiBatchTrack 
   Caption         =   "SFDC - Multi Booking Batch Tracking Questionnaire"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmMultiBatchTrack.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox ConfirmationText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      MaxLength       =   40
      TabIndex        =   19
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   960
      Top             =   5280
   End
   Begin VB.CommandButton cmdNextBookOff 
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      TabIndex        =   5
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CheckBox ChkComplete 
      Caption         =   "Is the Operation/Activity Completed?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3840
      Width           =   7095
   End
   Begin VB.CheckBox ChkPreOp 
      Caption         =   "Was Operation/Activity performed Out-of-Sequence? (Pre-Op)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4440
      Width           =   6975
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox WorkCentre 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox OpDescription 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   4575
   End
   Begin VB.TextBox OrderDesc 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   840
      Width           =   4215
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DtpPostTime 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Format          =   104071170
      CurrentDate     =   36917
   End
   Begin MSComCtl2.DTPicker DtpPostDate 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   104071168
      CurrentDate     =   36917
   End
   Begin VB.Label lblStatusMessage 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   600
      TabIndex        =   23
      Top             =   6120
      Width           =   7215
      WordWrap        =   -1  'True
   End
   Begin MSForms.ComboBox cmboReasons 
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Top             =   3120
      Width           =   855
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "1508;661"
      ListWidth       =   7055
      TextColumn      =   1
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;7055"
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Reason"
      Height          =   195
      Left            =   480
      TabIndex        =   21
      Top             =   2880
      Width           =   555
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
      Height          =   195
      Left            =   480
      TabIndex        =   20
      Top             =   2880
      Width           =   630
   End
   Begin VB.Label Batch 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Tracking: Please consider all questions CAREFULLY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   120
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   5760
      Picture         =   "frmMultiBatchTrack.frx":0442
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2340
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Post Dated - DATE"
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Top             =   2040
      Width           =   1365
   End
   Begin VB.Label Label12 
      Caption         =   "Post Dated - TIME"
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -240
      TabIndex        =   15
      Top             =   3840
      Width           =   6135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Op No"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   480
      TabIndex        =   14
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Work Centre"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   1320
      TabIndex        =   13
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Op Description"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   2520
      TabIndex        =   12
      Top             =   1320
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Number"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   600
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Description"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   2040
      TabIndex        =   10
      Top             =   600
      Width           =   1230
   End
End
Attribute VB_Name = "frmMultiBatchTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Dummy As Boolean
Dim LastEventTime As Date

Private Sub SetFormValues(Row As Integer)

   Dim tmpLastEvent As String
   Dim mbOrdTypeCfg As OrderTypeCfgInfo
   Dim Dummy As Boolean
   Dim PrevOpsCNF As String * 1

   'Initialise
   lblStatusMessage = ""

   'Read the appropriate Order Type Config parameters for this row
   mbOrdTypeCfg.Plant = frmMultiBooking.flxMultiBookings.TextMatrix(Row, 11)
   mbOrdTypeCfg.OrderType = frmMultiBooking.flxMultiBookings.TextMatrix(Row, 10)
   'nb. return the function to dummy as the Order Type Config will have
   'already been read when the Op was selected by its Conf Number
   Dummy = SetOrdTypeConfig(mbOrdTypeCfg)

   'Initialise the Popup form attributes
   If mbOrdTypeCfg.DefFinalConf = cnSAPTrue Then
      frmMultiBatchTrack.ChkComplete.Enabled = True
      frmMultiBatchTrack.ChkComplete.Value = cnChecked
   Else
      frmMultiBatchTrack.ChkComplete.Enabled = True
      frmMultiBatchTrack.ChkComplete.Value = cnUnChecked
   End If

   If gblUser.CanFinalConf <> cnSAPTrue Then
      frmMultiBatchTrack.ChkComplete.Enabled = False
      frmMultiBatchTrack.ChkComplete.Value = cnUnChecked
      lblStatusMessage = _
         "PLEASE NOTE: User not authorised for Final Confirmations"
   End If

   'Check whether this order type requires prev ops should be confirmed
   'before allowing the current op to be finally confirmed
   If mbOrdTypeCfg.Chk4PrevCNF = cnSAPTrue Then
      'Disable the Op out of Sequence Checkbox as backlush is NOT running
      frmMultiBatchTrack.ChkPreOp.Enabled = False
      'Read the whether the previous ops have been confirmed by calling the
      'SAP function module
      'First set the export parameter
      SAPChkPrevOps.Exports("I_RUECK") = frmMultiBooking.flxMultiBookings.TextMatrix(Row, 0)
      'then perform the RFC
      If SAPChkPrevOps.Call = False Then
         MsgBox SAPChkPrevOps.Exception, vbExclamation, cnDialogTitleCheckPass
         PrevOpsCNF = cnSAPFalse
      Else
         PrevOpsCNF = SAPChkPrevOps.Imports("OK2CNF")
      End If
      'If the previous milestone ops are not finally confirmed then prevent this op
      'from being finally confirmed
      If PrevOpsCNF = cnSAPFalse Then
         frmMultiBatchTrack.ChkComplete.Enabled = False
         frmMultiBatchTrack.ChkComplete.Value = cnUnChecked
         lblStatusMessage = _
            "PLEASE NOTE: Previous Milestone Ops Not Finally Confirmed - Further Final Confirmations NOT Permitted"
      End If
   Else
      frmMultiBatchTrack.ChkPreOp.Enabled = True
   End If

   With frmMultiBatchTrack
      .OrderNumber = frmMultiBooking.flxMultiBookings.TextMatrix(Row, 1)
      .OrderDesc = frmMultiBooking.flxMultiBookings.TextMatrix(Row, 9)
      .OpNumber = frmMultiBooking.flxMultiBookings.TextMatrix(Row, 2)
      .WorkCentre = frmMultiBooking.flxMultiBookings.TextMatrix(Row, 3)
      .OpDescription = frmMultiBooking.flxMultiBookings.TextMatrix(Row, 4)
      .ChkPreOp.Value = cnUnChecked
      .ConfirmationText = ""
      .DtpPostDate = gblSAP.LocalDateTime
      .DtpPostDate.MaxDate = gblSAP.LocalDateTime
      .DtpPostDate.MinDate = gblSAP.LocalDateTime - gblPlantCfg.CancelPeriod
      .DtpPostTime = gblSAP.LocalDateTime
   End With

End Sub

Private Sub cmdNextBookOff_Click()

   Dim tmpTime As String
   Dim PostDateTime As Date
   Dim ErrMsg As String
   Dim Row As Integer
   Dim BookOffFound As Boolean

   On Error GoTo ErrorHandler
   
   'Disable the Timer
   Timer1.Enabled = False
   
    'Check for Post dated Booking
    PostDateTime = DateValue(DtpPostDate) + TimeValue(DtpPostTime)

   'and set checkbox accordingly
   If PostDateTime < gblSAP.TimeOnEntry Then
      frmBatchTrack.ChkPostDated = cnChecked
   End If
   
    'Check for posting date/time in the future
    If PostDateTime > gblSAP.LocalDateTime Then
        MsgBox "CANNOT SET POSTING DATE/TIME IN FUTURE" & Chr(10) & _
                "CURRENT SAP TIME IS " & _
                FormatDateTime(gblSAP.LocalDateTime, vbGeneralDate), vbExclamation
        GoTo ExitOnError
    End If

   'Exit form and continue to update SAP
   gblBatchTrackOK = True
   Me.Hide

   Exit Sub

ExitOnError:
   'Reset StartTime for Display Period
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   'Switch on the timer
   Timer1.Enabled = True
   Exit Sub

ErrorHandler:
   MsgBox "Error:" & Err.Description & " in " & Err.Source
   Exit Sub

End Sub

Private Sub ConfirmationText_KeyPress(KeyAscii As Integer)

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub DtpPostDate_DropDown()

    'Reset StartTime
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    'Disable the timer
    Timer1.Enabled = False
    
End Sub


Private Sub DtpPostDate_LostFocus()

    'Enable the timer
    Timer1.Enabled = True
    
End Sub


Private Sub DtpPostTime_Change()

   'Reset the Start Time
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub DtpPostTime_Click()

   'Reset the Start Time
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub DtpPostTime_KeyPress(KeyAscii As Integer)

   'Reset the Start Time
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub Form_Activate()

 
    'Set the Start Time and Enable the Timer
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    Timer1.Enabled = True
        
    SetFormValues (1)


End Sub

Private Sub Form_Deactivate()

    Timer1.Enabled = False
    
End Sub



Private Sub Timer1_Timer()

    If Not OK2DisplayForm(False, gblPlantCfg.FormTimeOut) Then
        Timer1.Enabled = False
        Me.Hide
    End If

End Sub
