VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBookOff 
   Caption         =   "SFDC - Book Off Questionnaire"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   Icon            =   "frmBookOff.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOrderType 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   840
      Width           =   855
   End
   Begin VB.Frame fraYield 
      Caption         =   "Record Yield/Scrap"
      Height          =   1335
      Left            =   480
      TabIndex        =   22
      Top             =   3720
      Width           =   3255
      Begin VB.OptionButton optTimeOnly 
         Caption         =   "Time Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optPartialYield 
         Caption         =   "Partial Yield"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   2895
      End
      Begin VB.OptionButton optFullYield 
         Caption         =   "Full / Remaining Yield"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.TextBox txtOrderCategory 
      Height          =   375
      Left            =   6840
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
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
      MaxLength       =   39
      TabIndex        =   18
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   6120
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
      Left            =   2760
      TabIndex        =   5
      Top             =   5280
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
      Top             =   4320
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
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DtpPostTime 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   96075778
      CurrentDate     =   36917
   End
   Begin MSComCtl2.DTPicker DtpPostDate 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   96075776
      CurrentDate     =   36917
   End
   Begin VB.Label lblOrderType 
      BackStyle       =   0  'Transparent
      Caption         =   "Order Type"
      Height          =   255
      Left            =   6480
      TabIndex        =   28
      Top             =   600
      Width           =   1215
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
      Left            =   120
      TabIndex        =   26
      Top             =   4320
      Width           =   6135
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
      TabIndex        =   20
      Top             =   6480
      Width           =   7215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
      Height          =   195
      Left            =   480
      TabIndex        =   19
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
      TabIndex        =   17
      Top             =   120
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   5880
      Picture         =   "frmBookOff.frx":0442
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2340
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "BOOK OFF Date"
      Height          =   195
      Left            =   480
      TabIndex        =   16
      Top             =   2160
      Width           =   1185
   End
   Begin VB.Label Label12 
      Caption         =   "BOOK OFF Time"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   2160
      Width           =   1815
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
Attribute VB_Name = "frmBookOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Dummy As Boolean
Dim LastEventTime As Date

Dim YieldRow As Integer
Dim ScrapRow As Integer
Dim YieldRowSelected As Boolean
Dim ScrapRowSelected As Boolean
Dim PPSerNoRowCount As Integer

Private Sub SetFormValues(Row As Integer)


  If gblOrder.Category = "10" And gblPlantCfg.BookingMode = cnPostEventBooking Then ' ' TCR7117 Start ...
    Dim bbOrder As OrderInfo
    Dim bbOp As OpInfo
    Dim ListPosition As Integer
    Dim DialogResponse As Integer
    Dim RowSelected As Boolean
    Dim boborowcount As Integer
    Dim WarningFlagSet As Boolean

'Add a row to the BoBo table with appropriate values
        'nb. rows indexed from 1
        BoBoTable.AppendRow
        Row = "1" 'BoBoTable.RowCount
        BoBoTable.Value(Row, "MANDT") = gblSAP.Client
        BoBoTable.Value(Row, "PERS_NO") = gblUser.ClockNumber
        BoBoTable.Value(Row, "CONF_NO") = gblOperation.ConfNo
        BoBoTable.Value(Row, "UN_WORK") = gblOperation.WorkUnits
        BoBoTable.Value(Row, "ACTION") = cnBookOnAction
        BoBoTable.Value(Row, "ORDER_NO") = gblOrder.Number
        BoBoTable.Value(Row, "ACTIVITY") = gblOperation.Number
        BoBoTable.Value(Row, "APPLICATION") = gblOrder.Category
        BoBoTable.Value(Row, "AUFPL") = gblOrder.Aufpl
        BoBoTable.Value(Row, "APLZL") = gblOperation.Aplzl
        BoBoTable.Value(Row, "PLANT") = gblOperation.Plant
        BoBoTable.Value(Row, "WORK_CNTR") = gblOperation.WorkCentre
        BoBoTable.Value(Row, "ACT_TYPE") = gblOperation.ActivityType
        BoBoTable.Value(Row, "ARBID") = bbOp.Arbid
        BoBoTable.Value(Row, "ORD_DESC") = gblOrder.Desc
        BoBoTable.Value(Row, "ORD_TYPE") = gblOrder.Type
        BoBoTable.Value(Row, "OP_DESC") = gblOperation.Desc
  End If ' TCR7117 end ...
  
   Dim tmpLastEvent As String
   Dim bbOrdTypeCfg As OrderTypeCfgInfo
   Dim Dummy As Boolean
   Dim PrevOpsCNF As String * 1
   Dim LastEventDate As Date
   Dim LastEventDate_TS As String
   

   'Initialise
   lblStatusMessage = ""
  If gblPlantCfg.BookingMode <> cnPostEventBooking Then ' TCR7117 Start ...
    'Set the Last Event Time
   tmpLastEvent = BoBoTable.Value(Row, "LE_DATE") & " " _
      & BoBoTable.Value(Row, "LE_TIME")
   LastEventTime = CDate(tmpLastEvent)
  End If

   'Read the appropriate Order Type Config parameters for this BOBO row
   bbOrdTypeCfg.Plant = BoBoTable.Value(Row, "PLANT")
   bbOrdTypeCfg.OrderType = BoBoTable.Value(Row, "ORD_TYPE")
   'nb. return the function to dummy as the Order Type Config will have
   'already been read when the Op was selected by its Conf Number
   Dummy = SetOrdTypeConfig(bbOrdTypeCfg)

   'Initialise the Popup form attributes
   If bbOrdTypeCfg.DefFinalConf = cnSAPTrue Then
      frmBookOff.ChkComplete.Enabled = True
      frmBookOff.ChkComplete.Value = cnChecked
   Else
      frmBookOff.ChkComplete.Enabled = True
      frmBookOff.ChkComplete.Value = cnUnChecked
   End If

   If gblUser.CanFinalConf <> cnSAPTrue Then
      frmBookOff.ChkComplete.Enabled = False
      frmBookOff.ChkComplete.Value = cnUnChecked
      lblStatusMessage = _
         "PLEASE NOTE: User not authorised for Final Confirmations"
   End If

   'Check whether this order type requires prev ops should be confirmed
   'before allowing the current op to be finally confirmed
   If bbOrdTypeCfg.Chk4PrevCNF = cnSAPTrue Then
      'Disable the Op out of Sequence Checkbox as backflush
      'is NOT runnning
      frmBookOff.ChkPreOp.Enabled = False
      'Read the whether the previous ops have been confirmed by calling the
      'SAP function module
      'First set the export parameter
      SAPChkPrevOps.Exports("I_RUECK") = BoBoTable.Value(Row, "CONF_NO")
      'then perform the RFC
      If SAPChkPrevOps.Call = False Then
         MsgBox SAPChkPrevOps.Exception, vbExclamation, cnDialogTitleCheckPass
         PrevOpsCNF = cnSAPFalse
         PCNF = PrevOpsCNF ' -- TCR7117 --
      Else
         PrevOpsCNF = SAPChkPrevOps.Imports("OK2CNF")
         PCNF = PrevOpsCNF ' -- TCR7117 --
      End If
      'If the previous milestone ops are not finally confirmed then prevent this op
      'from being finally confirmed
      If PrevOpsCNF = cnSAPFalse Then
         frmBookOff.ChkComplete.Enabled = False
         frmBookOff.ChkComplete.Value = cnUnChecked
         lblStatusMessage = _
            "PLEASE NOTE: Previous Milestone Ops Not Finally Confirmed - Further Final Confirmations NOT Permitted"
      End If
   Else
      frmBookOff.ChkPreOp.Enabled = True

   End If
   
    If gblPlantCfg.BookingMode <> cnPostEventBooking Then  '  TCR7117 Start ...
        'LAST_EVENT is in SAP Server time (non local) and must be used when comparing to gblSAP.DateTime
        LastEventDate_TS = BoBoTable.Value(Row, "LAST_EVENT")
        LastEventDate = Mid(LastEventDate_TS, 1, 4) & "/" & Mid(LastEventDate_TS, 5, 2) & "/" & Mid(LastEventDate_TS, 7, 2)
   End If '  TCR7117 Start ...
   
   With frmBookOff
      .OrderNumber = Format(BoBoTable.Value(Row, "ORDER_NO"), "############")
      .OrderDesc = BoBoTable.Value(Row, "ORD_DESC")
      .OpNumber = BoBoTable.Value(Row, "ACTIVITY")
      .WorkCentre = BoBoTable.Value(Row, "WORK_CNTR")
      .OpDescription = BoBoTable.Value(Row, "OP_DESC")
      .ChkPreOp.Value = cnUnChecked
      .DtpPostDate.MaxDate = gblSAP.LocalDateTime
      .DtpPostDate.MinDate = DateValue(LastEventDate)
      '.DtpPostDate.MinDate = BoBoTable.Value(Row, "LE_DATE")
      .txtOrderCategory = BoBoTable.Value(Row, "APPLICATION")
      .txtOrderType = BoBoTable.Value(Row, "ORD_TYPE")
   End With

   If BoBoTable.Value(Row, "WARNING") = cnSAPTrue Then
      With frmBookOff
         .DtpPostDate.Enabled = True
         .DtpPostTime.Enabled = True
      End With
   Else
      With frmBookOff
         .DtpPostDate.Enabled = False
         .DtpPostTime.Enabled = False
      End With
   End If

   'Set the colour of the row being processed in frmBoBo
  If gblPlantCfg.BookingMode <> cnPostEventBooking Then ' TCR7117 Start ...
    frmBoBo.flxBoBo.Row = Row
    frmBoBo.flxBoBo.CellBackColor = vbRed
  End If ' TCR7117 Start ...
End Sub

Private Sub SetOrderCatFormValues(Row As Integer)

'Enable/disable form fields dependant on the Order category
    If frmBookOff.txtOrderCategory.Text = "10" Then
        If gblPlantCfg.BookingMode = cnPostEventBooking Then '  TCR7117 Start ...
            gblBoBoCurrentRow = 1
        End If
        If BoBoTable.Value(gblBoBoCurrentRow, "ORD_TYPE") = "MISC" Then
            With frmBookOff
                .ChkComplete.Visible = True
                .ChkPreOp.Visible = False
                .fraYield.Visible = True
                .optFullYield.Visible = True
                .optFullYield.Enabled = False
                .optFullYield.Value = False
                .optPartialYield.Visible = True
                .optPartialYield.Enabled = False
                .optPartialYield.Value = False
                .optTimeOnly.Visible = True
                .optTimeOnly.Enabled = True
                .optTimeOnly.Value = True
            End With
        Else
            With frmBookOff
                .ChkComplete.Visible = False
                .ChkPreOp.Visible = False
                .fraYield.Visible = True
                .optFullYield.Visible = True
                .optFullYield.Enabled = True
                .optFullYield.Value = False
                .optPartialYield.Visible = True
                .optPartialYield.Enabled = True
                .optPartialYield.Value = False
                .optTimeOnly.Visible = True
                .optTimeOnly.Visible = True
                .optTimeOnly.Value = True
                
            End With
        End If
    Else 'Order categrory is not 10
    
          With frmBookOff
              .ChkComplete.Visible = True
              .ChkPreOp.Visible = True
              .fraYield.Visible = False
              .optFullYield.Visible = False
              .optPartialYield.Visible = False
              .optTimeOnly.Visible = False
          End With
    End If
End Sub
Private Sub cmdNextBookOff_Click()

    Dim tmpTime As String
    Dim ThisPostTime As Date
    Dim ErrMsg As String
    Dim Row As Integer
    Dim BookOffFound As Boolean
    Dim QtyConsumed As Boolean
    Dim PPSerNoRowCount As Integer
   
    Dim RowIndex As Integer
    Dim Column As Integer
    Dim YieldScrapRow As PPSerNoList
    Dim NextWorkCentre As String
    Dim IntSerNos_Exist As Boolean
    Dim NonScrapped As Integer
    Dim AvailableYield As Double 'Integer --TCR7117 --
    Dim ExtRow As Integer
    Dim IntRow As Integer
    Dim IntPPSerNoRowCount As Integer
    Dim Scrap As Double 'Integer --TCR7117 --
    Dim InternalRecordedScrap As Integer


   ' On Error GoTo ErrorHandler
    
    'Disable the Timer
    frmBoBo.Timer1.Enabled = False
    gblYieldScrapFail = False
    
   'If gblOrder.Category = "10" Then ' --TCR7117--
   'If frmBookOff.optTimeOnly.Value = False Then '-- TCR7117 --
    'If PCNF = cnSAPFalse Then
     'If AvailableYield = (SAPReadYieldScrap.Imports("TOTAL_ORDER_QTY") - (SAPReadYieldScrap.Imports("OP_CURRENT_YIELD") + Scrap)) Then
          'cmdNextBookOff = False
          'MsgBox "Previous Milestone Ops Not Finally Confirmed - Further Final Confirmations NOT Permitted", _
           ' vbExclamation
          'Exit Sub
     'End If
    'End If
   'End If ' --TCR7117--
   'End If ' --TCR7117--
    'If necessary (ie. Warning flag set) validate the Book Off DateTime entered
    If BoBoTable.Value(gblBoBoCurrentRow, "WARNING") = cnSAPTrue Then

        'Set the appropriate DateTime Values for comparison
        'tmpTime = Format(DtpPostDate, "DD/MM/YYYY") & " " & Format(DtpPostTime, "HH:MM:SS")
        'tmpTime = Format(DtpPostDate, gblLocalDateFormat) & " " & Format(DtpPostTime, "HH:MM:SS")
          
        'ThisPostTime = CDate(tmpTime)
        ThisPostTime = DateValue(DtpPostDate) + TimeValue(DtpPostTime)

        If ThisPostTime < LastEventTime Then
            ErrMsg = "CANNOT POST DATE BOOK-OFF BEFORE LAST EVENT" & _
                Chr(10) & "on " & FormatDateTime(LastEventTime, vbGeneralDate)
            MsgBox ErrMsg, vbExclamation
            GoTo ExitOnError
        ElseIf ThisPostTime > gblSAP.LocalDateTime Then
            MsgBox "CANNOT SET BOOK-OFF DATE/TIME IN FUTURE" & Chr(10) & _
                "CURRENT SAP TIME IS " & _
                FormatDateTime(gblSAP.LocalDateTime, vbGeneralDate), vbExclamation
            GoTo ExitOnError
        End If

    End If

    'Set the colour of the row processed in frmBoBo
    If gblPlantCfg.BookingMode <> cnPostEventBooking Then ' TCR7117 Start ...
        frmBoBo.flxBoBo.Row = gblBoBoCurrentRow
        frmBoBo.flxBoBo.CellBackColor = vbGreen
    End If
    '*****************************************************
    'For PP orders, check which yield option was selected and process yield/serial numbers
    If frmBookOff.txtOrderCategory.Text = "10" Then
   
        'Calculate yield values
        'Read Order and Operation qtys
        SAPReadYieldScrap.Exports("CONFNO") = BoBoTable.Value(gblBoBoCurrentRow, "CONF_NO")
        SAPReadYieldScrap.Exports("ORDERNO") = BoBoTable.Value(gblBoBoCurrentRow, "ORDER_NO")
        SAPReadYieldScrap.Exports("SAP_DEBUG") = gblInstallation.SAPDebug
    
        'Call the RFC
        If SAPReadYieldScrap.Call = False Then
            'set pointer to standard
            Screen.MousePointer = vbDefault
            MsgBox SAPReadYieldScrap.Exception, vbExclamation, cnDialogTitleLogon
            Me.Hide
            Exit Sub
        End If
        
        InternalRecordedScrap = 0
        IntSerNos_Exist = False
        'Check if Serial Nos for the OrderNo already exist in the internal serial number table
        If IntPPSerNoLocnTable.RowCount > 0 Then
            For Row = 1 To IntPPSerNoLocnTable.RowCount
                If IntPPSerNoLocnTable.Value(Row, "AUFNR") = BoBoTable.Value(gblBoBoCurrentRow, "ORDER_NO") Then
                    IntSerNos_Exist = True
                    'Determine scrap quantity that is recorded in the internal table.
                    If IntPPSerNoLocnTable.Value(Row, "SCRAPIND") = "X" Then
                        InternalRecordedScrap = InternalRecordedScrap + 1
                    End If
                End If
            Next
        End If
        
        QtyConsumed = False
        'Check whether yield and scrap qtys are already consumed for the full order qty
        'First check whether to use the internal serial table scrap qty or the scrap
        'quantity already recorded in SAP
   
        If SAPReadYieldScrap.Imports("ORDER_CURRENT_SCRAP") < InternalRecordedScrap Then
            Scrap = InternalRecordedScrap
        Else
            Scrap = SAPReadYieldScrap.Imports("ORDER_CURRENT_SCRAP")
        End If

        AvailableYield = (SAPReadYieldScrap.Imports("TOTAL_ORDER_QTY") - (SAPReadYieldScrap.Imports("OP_CURRENT_YIELD") + Scrap))
        If sernp <> "" Then ' -- TCR7117 --
         If AvailableYield < 1 Then
          QtyConsumed = True
         End If
        Else
         If AvailableYield <= 0 Then
          QtyConsumed = True
         End If
        End If ' -- TCR7117 --
        If IntSerNos_Exist = False Then
            'No serial nos have been retrieved yet for the order number so determine Serial Numbers from Sap table or Order header
            'Set export parameters for SAPReadPPSerNos
            SAPReadPPSerNos.Exports("CONFNO") = BoBoTable.Value(gblBoBoCurrentRow, "CONF_NO")
            SAPReadPPSerNos.Exports("ORDERNO") = BoBoTable.Value(gblBoBoCurrentRow, "ORDER_NO")
            SAPReadPPSerNos.Exports("ORDERQTY") = SAPReadYieldScrap.Imports("TOTAL_ORDER_QTY")
            SAPReadPPSerNos.Exports("SAP_DEBUG") = gblInstallation.SAPDebug
   
   
      
            'Clear the tables prior to calling func module
            PPSerNoTable.FreeTable

            'Call the RFC to retrieve the serial numbers. These will either be the full list from the
            'order header or the ones from the Serial Number Location table
            If SAPReadPPSerNos.Call = False Then
                'Ignore MISC orders as these will not have any Serial Nos
                If frmBookOff.txtOrderType <> "MISC" Then
                    'set pointer to standard
                    Screen.MousePointer = vbDefault
                    MsgBox SAPReadPPSerNos.Exception, vbExclamation, cnDialogTitleLogon
                    Me.Hide
                    Exit Sub
                End If
            End If
            
            'Transfer records found into Internal PPSerNoLocn Table
            For ExtRow = 1 To PPSerNoTable.RowCount
                IntPPSerNoLocnTable.AppendRow
                IntRow = IntPPSerNoLocnTable.RowCount
                IntPPSerNoLocnTable(IntRow, "AUFNR") = PPSerNoTable.Value(ExtRow, "AUFNR")
                IntPPSerNoLocnTable(IntRow, "MATNR") = PPSerNoTable.Value(ExtRow, "MATNR")
                IntPPSerNoLocnTable(IntRow, "SERNR") = PPSerNoTable.Value(ExtRow, "SERNR")
                IntPPSerNoLocnTable(IntRow, "ARBPL") = PPSerNoTable.Value(ExtRow, "ARBPL")
                IntPPSerNoLocnTable(IntRow, "SCRAPIND") = PPSerNoTable.Value(ExtRow, "SCRAPIND")
                IntPPSerNoLocnTable(IntRow, "RUECK") = PPSerNoTable.Value(ExtRow, "RUECK")
                IntPPSerNoLocnTable(IntRow, "ZDATE") = PPSerNoTable.Value(ExtRow, "ZDATE")
                IntPPSerNoLocnTable(IntRow, "ZTIME") = PPSerNoTable.Value(ExtRow, "ZTIME")
            Next
            
        End If
       
       matnr = CStr(SAPReadPPSerNos.Imports("MATNR"))
       sernp = CStr(SAPReadPPSerNos.Imports("SERNP"))
       
        'Find work centre of the next operation
        'Set export parameters for SAPReadNextWrkCtr
        SAPReadNextWrkCtr.Exports("CURRENT_OP") = BoBoTable.Value(gblBoBoCurrentRow, "ACTIVITY")
        SAPReadNextWrkCtr.Exports("CURRENT_WRKCTR") = BoBoTable.Value(gblBoBoCurrentRow, "WORK_CNTR")
        SAPReadNextWrkCtr.Exports("I_AUFPL") = BoBoTable.Value(gblBoBoCurrentRow, "AUFPL")
        SAPReadNextWrkCtr.Exports("SAP_DEBUG") = gblInstallation.SAPDebug
      
        'Call the RFC to retrieve the next Work Centre.
        If SAPReadNextWrkCtr.Call = False Then
            'set pointer to standard
            Screen.MousePointer = vbDefault
            MsgBox SAPReadNextWrkCtr.Exception, vbExclamation, cnDialogTitleLogon
            Me.Hide
            Exit Sub
        End If
      
        NextWorkCentre = SAPReadNextWrkCtr.Imports("NEXT_WRKCTR")
      
        'Count the number of serial numbers retrieved
        PPSerNoRowCount = PPSerNoTable.RowCount
      
        'Check PP booking option (Time, Full or Partial Yield)
        If frmBookOff.optTimeOnly = True Then
            frmBookOff.ChkComplete = cnUnChecked
            BoBoTable.Value(gblBoBoCurrentRow, "YIELD") = 0
            BoBoTable.Value(gblBoBoCurrentRow, "SCRAP") = 0
        Else
            If frmBookOff.optFullYield = True Then
                'Populate yield qty on frmRecordYield
                frmRecordYield.txtConfirmedYieldQty = AvailableYield
                frmRecordYield.txtConfirmedScrapQty = 0
            Else
                frmRecordYield.txtConfirmedYieldQty = 0
                frmRecordYield.txtConfirmedScrapQty = 0
            End If
                
            'Check whether quantity is already fully consumed
            If QtyConsumed = True Then
                MsgBox "QUANTITY ALREADY FULLY CONSUMED" & Chr(13) & Chr(13) & "ONLY TIME BOOKING OPTION ALLOWED", vbExclamation, cnDialogTitleLogon
                frmBookOff.optTimeOnly.Value = True
                Exit Sub
            
            End If
            
            'Count records for the Order Number in the internal SerNoLocn Table that are not scrapped
            NonScrapped = 0
            For Row = 1 To IntPPSerNoLocnTable.RowCount
                If IntPPSerNoLocnTable.Value(Row, "AUFNR") = BoBoTable.Value(gblBoBoCurrentRow, "ORDER_NO") Then
                    If IntPPSerNoLocnTable.Value(Row, "SCRAPIND") <> "X" Then
                        NonScrapped = NonScrapped + 1
                    End If
                End If
            
            Next
            
            'Check for inconsistencies -> available yield should be less or equal the amount of records in the serial no list (non scrapped)
           
            If sernp <> "" Then ' -- TCR7117 --
             If AvailableYield > NonScrapped Then
                MsgBox "THERE ARE INCONSISTENCIES BETWEEN AVAILABLE YIELD AND THE SERIAL NUMBER LISTINGS" & Chr(13) & Chr(13) & "ONLY TIME BOOKING ALLOWED UNTIL THE SERIAL NUMBER LISTINGS HAVE BEEN CORRECTED", vbExclamation, cnDialogTitleLogon
                frmBookOff.optTimeOnly.Value = True
                Exit Sub
                End If
            End If
            
                    
            'Populate Header information on Record Yield/Scrap form
            With frmRecordYield
                .txtOrderNumber = Format(BoBoTable.Value(gblBoBoCurrentRow, "ORDER_NO"), "############")
                .txtOpNumber = BoBoTable.Value(gblBoBoCurrentRow, "ACTIVITY")
                .txtWorkCentre = BoBoTable.Value(gblBoBoCurrentRow, "WORK_CNTR")
                .txtConfNo = BoBoTable.Value(gblBoBoCurrentRow, "CONF_NO")
                .txtOpDescription = BoBoTable.Value(gblBoBoCurrentRow, "OP_DESC")
                 If sernp <> "" Then ' -- TCR7117 --
                .txtPartNumber = PPSerNoTable.Value(1, "MATNR")
                Else
                .txtPartNumber = matnr
                 End If
            End With
         
            'Populate Order and Operation qtys on the form
            With frmRecordYield
                  If sernp <> "" Then ' -- TCR7117 --
                  .txtOriginalOrderQty = Int(SAPReadYieldScrap.Imports("TOTAL_ORDER_QTY"))
                  Else
                  .txtOriginalOrderQty = FormatNumber(SAPReadYieldScrap.Imports("TOTAL_ORDER_QTY"), 3, , , vbFalse)
                  End If
                  If sernp <> "" Then ' -- TCR7117 --
                  .txtOrderScrapToDate = Scrap
                  .txtMaxYieldForOp = AvailableYield
                  Else
                  .txtOrderScrapToDate = FormatNumber(Scrap, 3, , , vbFalse)
                  .txtMaxYieldForOp = FormatNumber(AvailableYield, 3, , , vbFalse)
                  End If
                  If sernp <> "" Then ' -- TCR7117 --
                  .txtConfirmedYieldForOp = Int(SAPReadYieldScrap.Imports("OP_CURRENT_YIELD"))
                  Else
                  .txtConfirmedYieldForOp = FormatNumber(SAPReadYieldScrap.Imports("OP_CURRENT_YIELD"), 3, , , vbFalse)
                  End If
                  .txtNextWorkCtr = NextWorkCentre
             End With
         
            'Reset the Yield Flexgrid
            frmRecordYield.flxYield.Rows = 1
            frmRecordYield.flxYield.Row = 0
            frmRecordYield.flxYield.RowSel = 0
            YieldRowSelected = False
    
            'Reset the Scrap Flexgrid
            frmRecordYield.flxScrap.Rows = 1
            frmRecordYield.flxScrap.Row = 0
            frmRecordYield.flxScrap.RowSel = 0
            ScrapRowSelected = False
         
            IntPPSerNoRowCount = IntPPSerNoLocnTable.RowCount
      
            'Populate the Yield and Scrap flexigrids with serial numbers
            For RowIndex = 1 To IntPPSerNoRowCount
   
                If IntPPSerNoLocnTable.Value(RowIndex, "AUFNR") = BoBoTable.Value(gblBoBoCurrentRow, "ORDER_NO") Then
                    If IntPPSerNoLocnTable.Value(RowIndex, "SCRAPIND") <> "X" Then
                        YieldScrapRow.SerialNo = IntPPSerNoLocnTable.Value(RowIndex, "SERNR")
            
                        YieldRow = frmRecordYield.flxYield.Rows
                        frmRecordYield.flxYield.AddItem ("")
                        frmRecordYield.flxYield.TextMatrix(YieldRow, 0) = YieldScrapRow.SerialNo
      
                        ScrapRow = frmRecordYield.flxScrap.Rows
                        frmRecordYield.flxScrap.AddItem ("")
                        frmRecordYield.flxScrap.TextMatrix(ScrapRow, 0) = YieldScrapRow.SerialNo
                    End If
                 End If
            Next RowIndex
  
            'Display record yield form
            frmRecordYield.txtYieldRecordsSelected.Text = 0
            frmRecordYield.txtScrapRecordsSelected.Text = 0
            frmRecordYield.Show vbModal
            'Ensure timer in Record Yield form is disabled on return
            frmRecordYield.Timer1.Enabled = False
         
        End If
        
        
    End If
    GavailYield = AvailableYield  ' -- TCR7117 --
    gsval = AvailableYield  ' -- TCR7117 --
    If gblYieldScrapFail = True Then
        Exit Sub
    End If
    '**********************************************************

    'Set the Final Conf values in BOBO according to the popup responses
    If frmBookOff.ChkComplete = cnChecked Then
        BoBoTable.Value(gblBoBoCurrentRow, "FIN_CONF") = cnSAPTrue
    Else
        BoBoTable.Value(gblBoBoCurrentRow, "FIN_CONF") = cnSAPFalse
    End If

    'Set the Backflush values in BOBO according to the popup responses
    If frmBookOff.ChkPreOp = cnChecked Then
        BoBoTable.Value(gblBoBoCurrentRow, "BACKFLUSH") = cnSAPFalse
    Else
        BoBoTable.Value(gblBoBoCurrentRow, "BACKFLUSH") = cnSAPTrue
    End If

    'Set the Confirmation text
    BoBoTable.Value(gblBoBoCurrentRow, "CONF_TEXT") = frmBookOff.ConfirmationText

    'Set the date/time for the post dated booking if warning flag is set
    'otherwise book off date/time will be left blank to be set in SAP
    If BoBoTable.Value(gblBoBoCurrentRow, "WARNING") = cnSAPTrue Then
        BoBoTable.Value(gblBoBoCurrentRow, "OFF_DATE") = Format(frmBookOff.DtpPostDate, "YYYYMMDD")
        BoBoTable.Value(gblBoBoCurrentRow, "OFF_TIME") = Format(frmBookOff.DtpPostTime, "HHMMSS")
    End If
   


    'Determine the next Book Off Row and set the global variable
    'to reflect the row number
    BookOffFound = False
    For Row = gblBoBoCurrentRow + 1 To BoBoTable.RowCount
        If BoBoTable.Value(Row, "ACTION") = cnBookOffAction Then
            gblBoBoCurrentRow = Row
            BookOffFound = True
            Exit For
        End If
    Next Row

    If BookOffFound = True Then
        'Set the form fields for next iteration
        SetFormValues (gblBoBoCurrentRow)
        'Set the order category form fields for next iteration
        SetOrderCatFormValues (gblBoBoCurrentRow)
        
        
        GoTo ExitOnError
    Else
        'Exit form and continue to update SAP
        gblBatchTrackOK = True
        Me.Hide
    End If

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

    'Set initial values for Book Off Date and Time for the first iteration
    'the values set by the user will then remain for each iteration
    gblBatchTrackOK = False
    
    'Set the Start Time and Enable the Timer
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    Timer1.Enabled = True
    If gblPlantCfg.BookingMode = cnPostEventBooking Then
    With frmBookOff
        .DtpPostDate = gdate '*-- TCR7117 ---*
        .DtpPostTime = gdate '*-- TCR7117 ---*
        .ConfirmationText = ""
    End With
    Else
    With frmBookOff
        .DtpPostDate = gblSAP.LocalDateTime
        .DtpPostTime = gblSAP.LocalDateTime
        .ConfirmationText = ""
    End With
    End If
    
    SetFormValues (gblBoBoCurrentRow)
   
    SetOrderCatFormValues (gblBoBoCurrentRow)
    
    

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
