VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmBatchTrack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SFDC - Batch Tracking Questionnaire"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7890
   Icon            =   "frmBatchTrack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtScrapQty 
      Height          =   285
      Left            =   5880
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtYieldQty 
      Height          =   285
      Left            =   5880
      TabIndex        =   16
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame fraYield 
      Caption         =   "Record Yield/Scrap"
      Height          =   1455
      Left            =   360
      TabIndex        =   11
      Top             =   1560
      Width           =   3255
      Begin VB.OptionButton optConfirmationOnly 
         Caption         =   "Confirmation Only"
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.TextBox ConfirmationText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MaxLength       =   39
      TabIndex        =   7
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   3840
   End
   Begin VB.CheckBox ChkPostDated 
      Caption         =   "Is this a Post Dated or ""Catch Up"" Booking?"
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
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   6855
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
      Height          =   570
      Left            =   360
      TabIndex        =   3
      Top             =   1695
      Width           =   6975
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
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Tool tip text"
      Top             =   1200
      Width           =   5415
   End
   Begin VB.CommandButton CmdContinue 
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
      Height          =   855
      Left            =   2760
      TabIndex        =   1
      Top             =   3960
      Width           =   2055
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
      Left            =   0
      TabIndex        =   15
      Top             =   1680
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
      Left            =   480
      TabIndex        =   10
      Top             =   5040
      Width           =   7215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblReason 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   5760
      TabIndex        =   9
      Top             =   3360
      Width           =   1965
      WordWrap        =   -1  'True
   End
   Begin MSForms.ComboBox cmboReasons 
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   3360
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
   Begin VB.Label lblReasons 
      AutoSize        =   -1  'True
      Caption         =   "Zero Booking Reason"
      Height          =   195
      Left            =   4800
      TabIndex        =   6
      Top             =   3120
      Width           =   1560
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   3120
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
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   5520
      Picture         =   "frmBatchTrack.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2340
   End
End
Attribute VB_Name = "frmBatchTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Dummy As Boolean


Private Sub cmboReasons_Change()

   Dim Index As Integer

   Index = frmBatchTrack.cmboReasons.ListIndex
   If Index > -1 Then
      frmBatchTrack.lblReason.Caption = frmBatchTrack.cmboReasons.Column(1, Index)
   End If

End Sub

Private Sub cmboReasons_Click()

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   
End Sub

Private Sub cmboReasons_DropButtonClick()

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   
End Sub

Private Sub cmdContinue_Click()

Dim QtyConsumed As Boolean 'Integer ' --TCR7117 --
Dim AvailableYield As Double 'Integer ' --TCR7117 --
Dim IntSerNos_Exist As Boolean
Dim Row As Integer
Dim ExtRow As Integer
Dim IntRow As Integer
Dim NextWorkCentre As String
Dim PPSerNoRowCount As Integer
Dim NonScrapped As Integer
Dim YieldRowSelected As Boolean
Dim ScrapRowSelected As Boolean
Dim IntPPSerNoRowCount As Integer
Dim RowIndex As Integer
Dim YieldScrapRow As PPSerNoList
Dim YieldRow As Integer
Dim ScrapRow As Integer
Dim PrevOpsCNF As String * 1 ' -- TCR7117 --
    
   'Check that a zero time confirmation reason has been entered
   If gblFunction = cnZeroTimeConfirm Then
      If cmboReasons = "" Then
         MsgBox "REASON MUST BE ENTERED FOR ZERO BOOKING", _
            vbExclamation, cnDialogTitleWorkBook
         cmboReasons.SetFocus
         GoTo ExitOnError
      End If
   End If
    '*****************************************************
    'For PP orders, check which yield option was selected and process yield/serial numbers
    If gblOrder.Category = "10" Then
   
        'Calculate yield values
        'Read Order and Operation qtys
        SAPReadYieldScrap.Exports("CONFNO") = gblOperation.ConfNo
        SAPReadYieldScrap.Exports("ORDERNO") = gblOrder.Number
        SAPReadYieldScrap.Exports("SAP_DEBUG") = gblInstallation.SAPDebug
    
        'Call the RFC
        If SAPReadYieldScrap.Call = False Then
            'set pointer to standard
            Screen.MousePointer = vbDefault
            MsgBox SAPReadYieldScrap.Exception, vbExclamation, cnDialogTitleLogon
            Me.Hide
            GoTo ExitOnError
        End If
        
        QtyConsumed = False
        'Check whether yield and scrap qtys are already consumed for the full order qty
        AvailableYield = SAPReadYieldScrap.Imports("TOTAL_ORDER_QTY") - SAPReadYieldScrap.Imports("OP_CURRENT_YIELD") - SAPReadYieldScrap.Imports("ORDER_CURRENT_SCRAP")
        If sernp <> "" Then ' -- TCR7117 --
         If AvailableYield < 1 Then
          QtyConsumed = True
         End If
        Else
         If AvailableYield <= 0 Then
          QtyConsumed = True
         End If
        End If ' -- TCR7117 --
        
        IntSerNos_Exist = False
        'Check if Serial Nos for the OrderNo already exist in the internal serial number table
        If IntPPSerNoLocnTable.RowCount > 0 Then
            For Row = 1 To IntPPSerNoLocnTable.RowCount
                If IntPPSerNoLocnTable.Value(Row, "AUFNR") = gblOrder.Number Then
                    IntSerNos_Exist = True
                    Exit For
                End If
            Next
        End If
    
        If IntSerNos_Exist = False Then
            'No serial nos have been retrieved yet for the order number so determine Serial Numbers from Sap table or Order header
            'Set export parameters for SAPReadPPSerNos
            SAPReadPPSerNos.Exports("CONFNO") = gblOperation.ConfNo
            SAPReadPPSerNos.Exports("ORDERNO") = gblOrder.Number
            SAPReadPPSerNos.Exports("ORDERQTY") = SAPReadYieldScrap.Imports("TOTAL_ORDER_QTY")
            SAPReadPPSerNos.Exports("SAP_DEBUG") = gblInstallation.SAPDebug
   
   
      
            'Clear the tables prior to calling func module
            PPSerNoTable.FreeTable

            'Call the RFC to retrieve the serial numbers. These will either be the full list from the
            'order header or the ones from the Serial Number Location table
            If SAPReadPPSerNos.Call = False Then
                'Ignore MISC orders as these will not have any Serial Nos
                If gblOrder.Type <> "MISC" Then
                    'set pointer to standard
                    Screen.MousePointer = vbDefault
                    MsgBox SAPReadPPSerNos.Exception, vbExclamation, cnDialogTitleLogon
                    Me.Hide
                    GoTo ExitOnError
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
         matnr = CStr(SAPReadPPSerNos.Imports("MATNR")) '-- TCR7117 --
        'Find work centre of the next operation
        'Set export parameters for SAPReadNextWrkCtr
        SAPReadNextWrkCtr.Exports("CURRENT_OP") = gblOperation.Number
        SAPReadNextWrkCtr.Exports("CURRENT_WRKCTR") = gblOperation.WorkCentre
        SAPReadNextWrkCtr.Exports("I_AUFPL") = gblOrder.Aufpl
        SAPReadNextWrkCtr.Exports("SAP_DEBUG") = gblInstallation.SAPDebug
      
        'Call the RFC to retrieve the next Work Centre.
        If SAPReadNextWrkCtr.Call = False Then
            'set pointer to standard
            Screen.MousePointer = vbDefault
            MsgBox SAPReadNextWrkCtr.Exception, vbExclamation, cnDialogTitleLogon
            Me.Hide
            GoTo ExitOnError
        End If
       'Call function module for previous Op check
        SAPChkPrevOps.Exports("I_RUECK") = GCONF ' -- TCR7117 --
        'then perform the RFC
        If SAPChkPrevOps.Call = True Then
         PrevOpsCNF = SAPChkPrevOps.Imports("OK2CNF") ' -- TCR7117 --
         PCNF = PrevOpsCNF
         'If PCNF = cnSAPFalse Then ' -- TCR7117 --
          'If AvailableYield = SAPReadYieldScrap.Imports("TOTAL_ORDER_QTY") - SAPReadYieldScrap.Imports("OP_CURRENT_YIELD") - SAPReadYieldScrap.Imports("ORDER_CURRENT_SCRAP") Then
             'cmdContinue = False
            ' MsgBox "Previous Milestone Ops Not Finally Confirmed - Further Final Confirmations NOT Permitted", _
            ' vbExclamation
           ' Exit Sub ' -- TCR7117 --
         ' End If
         'End If
        End If     '-- TCR7117 --
        GCONF = "" '-- TCR7117 --
        GavailYield = AvailableYield  ' -- TCR7117 --
        gsval = AvailableYield  ' -- TCR7117 --
        NextWorkCentre = SAPReadNextWrkCtr.Imports("NEXT_WRKCTR")
      
        'Count the number of serial numbers retrieved
        PPSerNoRowCount = PPSerNoTable.RowCount
      
        'Check PP booking option (Conf Only, Full or Partial Yield)
        If frmBatchTrack.optConfirmationOnly = True Then
            'frmBatchTrack.ChkComplete.Value = 0
            frmBatchTrack.txtYieldQty = 0
            frmBatchTrack.txtScrapQty = 0
            gblYieldScrapFail = False
        Else
            If frmBatchTrack.optFullYield = True Then
                'Populate yield qty on frmRecordYield
                frmRecordYield.txtConfirmedYieldQty = AvailableYield
                frmRecordYield.txtConfirmedScrapQty = 0
            Else
                frmRecordYield.txtConfirmedYieldQty = 0
                frmRecordYield.txtConfirmedScrapQty = 0
            End If
                
            'Check whether quantity is already fully consumed
            If QtyConsumed = True Then
                MsgBox "QUANTITY ALREADY FULLY CONSUMED - ONLY TIME CONFIRMATION OPTION ALLOWED", vbExclamation, cnDialogTitleLogon
                frmBatchTrack.optConfirmationOnly.Value = True
                gblBatchTrackOK = False
                GoTo ExitOnError
            
            End If
            
            'Count records for the Order Number in the internal SerNoLocn Table that are not scrapped
            NonScrapped = 0
            For Row = 1 To IntPPSerNoLocnTable.RowCount
                If IntPPSerNoLocnTable.Value(Row, "AUFNR") = gblOrder.Number Then
                    If IntPPSerNoLocnTable.Value(Row, "SCRAPIND") <> "X" Then
                        NonScrapped = NonScrapped + 1
                    End If
                End If
            
            Next
            
            'Check for inconsistencies -> available yield should be less or equal the amount of records in the serial no list (non scrapped)
            If sernp <> "" Then ' -- TCR7117 --
            If AvailableYield > NonScrapped Then
                MsgBox "THERE ARE INCONSISTENCIES BETWEEN AVAILABLE YIELD AND THE SERIAL NUMBER LISTINGS" & Chr(13) & Chr(13) & "ONLY CONFIRMATION OPTION ALLOWED UNTIL THE SERIAL NUMBER LISTINGS HAVE BEEN CORRECTED", vbExclamation, cnDialogTitleLogon
                frmBatchTrack.optConfirmationOnly.Value = True
                GoTo ExitOnError
            End If
            End If
                    
            'Populate Header information on Record Yield/Scrap form
            With frmRecordYield
                .txtOrderNumber = Format((gblOrder.Number), "############")
                .txtOpNumber = gblOperation.Number
                .txtConfNo = gblOperation.ConfNo
                .txtWorkCentre = gblOperation.WorkCentre
                .txtOpDescription = gblOperation.Desc
                 If gblOrder.Category = "10" Then
                 .txtPartNumber = matnr
                 Else
                .txtPartNumber = PPSerNoTable.Value(1, "MATNR")
                End If
            End With
         
            'Populate Order and Operation qtys on the form
            With frmRecordYield
                If sernp <> "" Then ' -- TCR7117 --
                .txtOriginalOrderQty = Int(SAPReadYieldScrap.Imports("TOTAL_ORDER_QTY"))
                Else
                '.txtOriginalOrderQty = SAPReadYieldScrap.Imports("TOTAL_ORDER_QTY")
                .txtOriginalOrderQty = FormatNumber(SAPReadYieldScrap.Imports("TOTAL_ORDER_QTY"), 3, , , vbFalse)
                End If
                
                If sernp <> "" Then ' -- TCR7117 --
                .txtOrderScrapToDate = Int(SAPReadYieldScrap.Imports("ORDER_CURRENT_SCRAP"))
                Else
                .txtOrderScrapToDate = FormatNumber(SAPReadYieldScrap.Imports("ORDER_CURRENT_SCRAP"), 3, , , vbFalse)
                End If
                 If sernp <> "" Then ' -- TCR7117 --
                 .txtMaxYieldForOp = AvailableYield
                 Else
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
   
                If IntPPSerNoLocnTable.Value(RowIndex, "AUFNR") = gblOrder.Number Then
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
   
    If gblYieldScrapFail = True Then
        GoTo ExitOnError
    End If
    '**********************************************************
   

   gblBatchTrackOK = True
   Timer1.Enabled = False
   Me.Hide

   Exit Sub

ExitOnError:
   'Reset StartTime for Display Period
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   'Switch on the timer
   Timer1.Enabled = True
   Exit Sub

End Sub



Private Sub ConfirmationText_KeyPress(KeyAscii As Integer)

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub Form_Activate()

   'Initialise global to ensure user has not exited abnormally
   gblBatchTrackOK = False

   'Load the Reasons combo box
   Call LoadReasons

   'Set the Start Time and Enable the Timer
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

   'Initialise
   Timer1.Enabled = True
   lblStatusMessage = ""
   lblReason = ""

   'Check function and set default values for BatchTracking accordingly
   Select Case gblFunction

      Case cnWorkBook

         With frmBatchTrack
            'Set default settings for the form
            Call SetFormDefaults
            'Now selectively overide form settings if required
            '.ChkPostDated.Value = cnUnChecked
            .ChkPreOp.Caption = "Was the Operation performed Out-of-Sequence? (Pre-Op)"
            .ChkPreOp.Value = cnUnChecked
            .ChkComplete.Caption = "Is the Operation/Activity Completed?"
            .ChkPostDated.Caption = "Is this a Post Dated or ""Catch Up"" Booking?"
            .ConfirmationText = ""
            .cmboReasons.Visible = False
            .cmboReasons.Enabled = False
            .lblReasons.Visible = False
         End With

      Case cnZeroTimeConfirm

         If gblOrder.Category = "10" Then
            If gblOrder.Type = "MISC" Then
                With frmBatchTrack
                    'Set default settings for the form
                    Call SetFormDefaults
                    'Now selectively overide form settings if required
                    'For MISC PP orders disable Post Dated option
                    .ChkPostDated.Value = cnUnChecked
                    .ChkPostDated.Enabled = False
                    .ChkPostDated.Visible = False
                    '.ChkPostDated.Caption = "Is this a Post Dated or ""Catch Up"" Booking?"
                    
                    'For MISC PP orders disable PreOp option
                    .ChkPreOp.Value = cnUnChecked
                    .ChkPreOp.Visible = False
                    .ChkPreOp.Caption = "Was the Operation performed Out-of-Sequence? (Pre-Op)"
                    
                    'For MISC PP orders use default settings completed option
                    .ChkComplete.Caption = "Is the Operation/Activity Completed?"
                    
                    .ConfirmationText = ""
                    .cmboReasons = ""
                    .cmboReasons.Visible = True
                    .cmboReasons.Enabled = True
                    .lblReasons.Visible = True
                    
                    'For MISC PP orders ensure only confirmation option is available
                    fraYield.Visible = True
                    optFullYield.Visible = True
                    optFullYield.Enabled = False
                    optPartialYield.Visible = True
                    optPartialYield.Enabled = False
                    optConfirmationOnly.Visible = True
                    optConfirmationOnly.Enabled = True
                    optConfirmationOnly.Value = True
                End With
            Else ' not MISC PP Order
                With frmBatchTrack
                    'Set default settings for the form
                    Call SetFormDefaults
                    'Now selectively overide form settings if required
                    'For non MISC PP orders disable Post Dated option
                    .ChkPostDated.Value = cnUnChecked
                    .ChkPostDated.Enabled = False
                    .ChkPostDated.Visible = False
                    '.ChkPostDated.Caption = "Is this a Post Dated or ""Catch Up"" Booking?"
                    
                    'For non MISC PP orders disable PreOp option
                    .ChkPreOp.Value = cnUnChecked
                    .ChkPreOp.Visible = False
                    .ChkPreOp.Caption = "Was the Operation performed Out-of-Sequence? (Pre-Op)"
                    
                    'For non MISC PP orders enable but hide the completed option
                    .ChkComplete.Caption = "Is the Operation/Activity Completed?"
                    .ChkComplete.Enabled = True
                    .ChkComplete.Visible = False
                    
                    .ConfirmationText = ""
                    .cmboReasons = ""
                    .cmboReasons.Visible = True
                    .cmboReasons.Enabled = True
                    .lblReasons.Visible = True
                    
                    'For non MISC PP orders ensure all options are available
                    fraYield.Visible = True
                    optFullYield.Visible = True
                    optFullYield.Enabled = True
                    optPartialYield.Visible = True
                    optPartialYield.Enabled = True
                    optConfirmationOnly.Visible = True
                    optConfirmationOnly.Enabled = True
                End With
                
            End If
        Else 'Order cat <> '10'
            With frmBatchTrack
            
                'Set default settings for the form
                Call SetFormDefaults
                'Now selectively overide form settings if required
                .ChkPostDated.Value = cnUnChecked
                '.ChkPostDated.Caption = "Is this a Post Dated or ""Catch Up"" Booking?"
                .ChkPreOp.Value = cnUnChecked
                .ChkComplete.Caption = "Is the Operation/Activity Completed?"
                .ChkPreOp.Caption = "Was the Operation performed Out-of-Sequence? (Pre-Op)"
                .ChkPostDated.Enabled = False
                .ConfirmationText = ""
                .ChkPostDated.Visible = True
                
                .cmboReasons = ""
                .cmboReasons.Visible = True
                .cmboReasons.Enabled = True
                .lblReasons.Visible = True
                
                If gblOrder.Type = "ZS01" Then
                    With frmBatchTrack
                        .ChkPreOp.Enabled = False
                    End With
                End If
            End With
        End If

      Case cnMultiBook

         With frmBatchTrack
            'Set default settings for the form
            Call SetFormDefaults
            '.ChkPostDated.Value = cnUnChecked
            .ChkPreOp.Value = cnUnChecked
            .ChkComplete.Caption = "Are the Operations/Activities Completed?"
            .ChkPreOp.Caption = "Were Operations performed Out-of-Sequence? (Pre-Ops)"
            .ChkPostDated.Caption = "Are these Post Dated or ""Catch Up"" Bookings?"
            .ChkPostDated.Visible = True
            .ConfirmationText = ""
            .cmboReasons.Visible = False
            .cmboReasons.Enabled = False
            .lblReasons.Visible = False
         End With

   End Select

End Sub

Private Sub Form_Deactivate()

   Timer1.Enabled = False

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Timer1_Timer()

   If Not OK2DisplayForm(False, gblPlantCfg.FormTimeOut) Then
      Timer1.Enabled = False
      Me.Hide
   End If

End Sub
Private Sub SetFormDefaults()

'Set Completion checkbox regardless of function being performed
   If gblOrdTypeCfg.DefFinalConf = cnSAPTrue Then
      frmBatchTrack.ChkComplete.Enabled = True
      frmBatchTrack.ChkComplete.Visible = True
      frmBatchTrack.ChkComplete.Value = cnChecked
   Else
      frmBatchTrack.ChkComplete.Enabled = True
      frmBatchTrack.ChkComplete.Visible = True
      frmBatchTrack.ChkComplete.Value = cnUnChecked
   End If

   'Check user is authorised to perform final confirmations
   If gblUser.CanFinalConf <> cnSAPTrue Then
      frmBatchTrack.ChkComplete.Enabled = False
      frmBatchTrack.ChkComplete.Visible = True
      frmBatchTrack.ChkComplete.Value = cnUnChecked
      lblStatusMessage = _
         "PLEASE NOTE: User not authorised for Final Confirmations"
   End If

   'If required check that all previous ops have been finally confirmed
   If gblOrdTypeCfg.Chk4PrevCNF = cnSAPTrue Then
      'Disable "Op out of Sequence" checkbox as backflush is NOT running
      frmBatchTrack.ChkPreOp.Enabled = False
      frmBatchTrack.ChkPreOp.Visible = True
      'Check that all prev milestone ops are finally confirmed
      If gblOperation.PrevOpsCNF <> cnSAPTrue Then
         frmBatchTrack.ChkComplete.Enabled = False
         frmBatchTrack.ChkComplete.Visible = True
         frmBatchTrack.ChkComplete.Value = cnUnChecked
         lblStatusMessage = _
            "PLEASE NOTE: Previous Milestone Ops Not Finally Confirmed - Further Final Confirmations NOT Permitted"
      End If
   Else
      frmBatchTrack.ChkPreOp.Enabled = True
      frmBatchTrack.ChkPreOp.Visible = True
   End If
   
   'Set all PP fields
   fraYield.Visible = False
   optFullYield.Visible = False
   optFullYield.Enabled = False
   optPartialYield.Visible = False
   optPartialYield.Enabled = False
   optConfirmationOnly.Visible = False
   optConfirmationOnly.Enabled = False
   
End Sub
Private Sub LoadReasons()

   Dim i As Integer, Row As Integer, ItemCount As Integer
   Dim Code As String, Desc As String, Item As String

   'Clear the existing items in the combo box in readiness
   'for loading the new items appropriate to the order type
   'frmBatchTrack.cmboReasons.Refresh
   ItemCount = frmBatchTrack.cmboReasons.ListCount
   For i = 1 To ItemCount
      frmBatchTrack.cmboReasons.RemoveItem (0)
   Next i

   'loop thru Reasons Table and add items to Combo box List
   For i = 1 To ReasonsTable.RowCount
      Code = ReasonsTable.Value(i, "GRUND")
      Desc = ReasonsTable.Value(i, "GRDTX")
      If Code >= gblOrdTypeCfg.ReasonsLower And Code <= gblOrdTypeCfg.ReasonsUpper Then
         Row = Row + 1
         frmBatchTrack.cmboReasons.AddItem
         frmBatchTrack.cmboReasons.Column(0, Row - 1) = Code
         frmBatchTrack.cmboReasons.Column(1, Row - 1) = Desc
      End If
   Next i

End Sub
