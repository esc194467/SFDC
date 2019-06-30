VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRecordYield 
   Caption         =   "Declare Yield / Scrap"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7395
   Icon            =   "frmRecordYield.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtScrapRecordsSelected 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "0"
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox txtConfNo 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdYieldSelectAll 
      Height          =   375
      Left            =   360
      Picture         =   "frmRecordYield.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   " Select All "
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdYieldDeSelectAll 
      Height          =   375
      Left            =   360
      Picture         =   "frmRecordYield.frx":05CE
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   " De-Select All "
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox txtYieldRecordsSelected 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "0"
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox txtNextWorkCtr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtPartNumber 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtOrderNumber 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtOpDescription 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1080
      Width           =   4815
   End
   Begin VB.TextBox txtWorkCentre 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtOpNumber 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1080
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   8040
   End
   Begin VB.TextBox txtConfirmedYieldForOp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtMaxYieldForOp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtOrderScrapToDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtOriginalOrderQty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   8280
      Width           =   1335
   End
   Begin VB.TextBox txtConfirmedScrapQty 
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Text            =   "0"
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox txtConfirmedYieldQty 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "0"
      Top             =   7440
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flxScrap 
      Height          =   3375
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "Click on Serial numbers to select"
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid flxYield 
      Height          =   3375
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Click on Serial numbers to select"
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin MSForms.Label lblScrapCount 
      Height          =   255
      Left            =   4200
      TabIndex        =   38
      Top             =   6840
      Width           =   1455
      Caption         =   "Records Selected:"
      Size            =   "2566;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblConfNo 
      Caption         =   "Conf No"
      Height          =   255
      Left            =   4200
      TabIndex        =   37
      Top             =   120
      Width           =   1335
   End
   Begin MSForms.Label lblYieldCount 
      Height          =   255
      Left            =   1200
      TabIndex        =   32
      Top             =   6840
      Width           =   1455
      Caption         =   "Records Selected:"
      Size            =   "2566;450"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblNextWorkCtr 
      Caption         =   "Next Work Centre"
      Height          =   255
      Left            =   5760
      TabIndex        =   31
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblOrderConfirmedScrapToDate 
      Caption         =   "Confirmed Scrap To Date:"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblPartNo 
      Caption         =   "Part Number"
      Height          =   255
      Left            =   2040
      TabIndex        =   29
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblOrderNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Number"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   120
      Width           =   990
   End
   Begin VB.Label lblOpDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Op Description"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   2280
      TabIndex        =   26
      Top             =   840
      Width           =   1050
   End
   Begin VB.Label lblWorkCentre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Work Centre"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   1080
      TabIndex        =   25
      Top             =   840
      Width           =   900
   End
   Begin VB.Label lblOpNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Op No"
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   840
      Width           =   465
   End
   Begin VB.Label lblConfirmedYieldForOp 
      Caption         =   "Confirmed Yield To Date:"
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblMaxYieldForOp 
      Caption         =   "Maximum Yield Available:"
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblOpQtys 
      Caption         =   "Operation Qtys:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Line Line8 
      X1              =   7080
      X2              =   3840
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line7 
      DrawMode        =   16  'Merge Pen
      X1              =   3840
      X2              =   3840
      Y1              =   2880
      Y2              =   1680
   End
   Begin VB.Line Line6 
      DrawMode        =   16  'Merge Pen
      X1              =   3840
      X2              =   7080
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line5 
      X1              =   7080
      X2              =   7080
      Y1              =   2880
      Y2              =   1680
   End
   Begin VB.Line Line4 
      DrawMode        =   16  'Merge Pen
      X1              =   240
      X2              =   240
      Y1              =   1680
      Y2              =   2880
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   3360
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      X1              =   3360
      X2              =   3360
      Y1              =   2880
      Y2              =   1680
   End
   Begin VB.Line Line1 
      DrawMode        =   16  'Merge Pen
      X1              =   240
      X2              =   3360
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblExpectedOrderQty 
      Caption         =   "Original Qty:"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblConfirmScrapQty 
      Caption         =   "Confirm Scrap Qty:"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label lblConfirmYieldQty 
      Caption         =   "Confirm Yield Qty:"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label lblScrap 
      Caption         =   "Record Scrap:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblYield 
      Caption         =   "Record Yield:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblOrderQtys 
      Caption         =   "Order Qtys:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   1695
   End
End
Attribute VB_Name = "frmRecordYield"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dummy As Boolean

Dim ListPosition As Integer
Dim DialogResponse As Integer
Dim YieldRow As Integer
Dim ScrapRow As Integer
Dim YieldRowSelected As Boolean
Dim ScrapRowSelected As Boolean
Dim PPSerNoRowCount As Integer



Private Sub Image1_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblExpectedOpQty_Click()

End Sub

Private Sub cmdCancel_Click()

    gblYieldScrapFail = True
    Me.Hide

End Sub

Private Sub cmdContinue_Click()

    Dim YieldSelected As Boolean
    Dim ScrapSelected As Boolean
    Dim YieldCount As Integer
    Dim ScrapCount As Integer
    Dim RowIndex As Integer
    Dim IntPPSerNoRowCount As Integer
    
   If frmBookOff.optTimeOnly.Value = False Then '-- TCR7117 --
    If PCNF = cnSAPFalse Then
     If Val(txtMaxYieldForOp.Text) = Val(txtConfirmedYieldQty) + Val(txtConfirmedScrapQty) Then
          cmdNextBookOff = False
          MsgBox "Previous Milestone Ops Not Finally Confirmed - Further Final Confirmations NOT Permitted", _
            vbExclamation
          Exit Sub
     End If
    End If
   End If ' --TCR7117--
   
    'Reset StartTime
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    'Switch off the timer
    Timer1.Enabled = False
   
    YieldCount = 0
    YieldSelected = False
    
    'Check whether there are any Yield rows selected
    If frmRecordYield.flxYield.Rows > 1 Then
        
        For Row = 1 To (flxYield.Rows - 1)
            flxYield.Row = Row
            
            If frmRecordYield.flxYield.CellBackColor = vbYellow Then
                YieldSelected = True
                YieldCount = YieldCount + 1
            Else
                'Ignore value
            End If
        Next
        
    Else 'No yield rows selected
        'Do nothing
    End If
    
   ScrapCount = 0
   ScrapSelected = False
    'Check whether there are any Scrap rows selected
    If frmRecordYield.flxScrap.Rows > 1 Then
        
        For Row = 1 To (flxScrap.Rows - 1)
            flxScrap.Row = Row
            
            If frmRecordYield.flxScrap.CellBackColor = vbYellow Then
                ScrapSelected = True
                ScrapCount = ScrapCount + 1
            Else
                'Ignore value
            End If
        Next
        
    Else 'No scrap rows selected
        'Do nothing
    End If
   
    gblYieldScrapFail = False
    If sernp <> "" Then ' -- TCR7117 --
    'Check whether Yield and Scrap values have been selected
    If YieldSelected = False And ScrapSelected = False Then
        MsgBox "NO YIELD OR SCRAP SELECTED - PLEASE MAKE TIME ONLY/CONFIRMATION ONLY BOOKING", vbExclamation, cnDialogTitleLogon
        GoTo ExitOnError
    End If
   
    'Check whether Confirmed Yield value = Yield rows selected
    If YieldCount <> Val(txtConfirmedYieldQty) Then
        MsgBox "YIELD SERIAL NUMBERS SELECTED AND CONFIRMED YIELD QUANTITY ARE INCONSISTENT - PLEASE CORRECT.", vbExclamation, cnDialogTitleLogon
        GoTo ExitOnError
    End If
   
     'Check whether Confirmed Scrap value = Scrap rows selected
       If ScrapCount <> Val(txtConfirmedScrapQty) Then
        MsgBox "SCRAP SERIAL NUMBERS SELECTED AND CONFIRMED SCRAP QUANTITY ARE INCONSISTENT - PLEASE CORRECT.", vbExclamation, cnDialogTitleLogon
        GoTo ExitOnError
    End If
    End If
    
    'Check whether Yield and Scrap values are not greater than available yield
    If YieldCount + ScrapCount > Val(txtMaxYieldForOp.Text) Then
        MsgBox "YIELD PLUS SCRAP VALUES ARE GREATER THAN AVAILABLE YIELD - PLEASE CORRECT", vbExclamation, cnDialogTitleLogon
        GoTo ExitOnError
    End If
    If Val(txtConfirmedYieldQty) + Val(txtConfirmedScrapQty) > Val(txtMaxYieldForOp.Text) Then '--TCR7117 --
        MsgBox "YIELD PLUS SCRAP VALUES ARE GREATER THAN AVAILABLE YIELD - PLEASE CORRECT", vbExclamation, cnDialogTitleLogon
        GoTo ExitOnError
    End If  '--TCR7117 --
    'Set Final Conf if Yield and Scrap value = Available Yield
    If sernp <> "" Then '--TCR7117 --
    If YieldCount + ScrapCount = Val(txtMaxYieldForOp.Text) Then
        'Set Final Confirmation for BoBo processing
        If gblFunction = cnBookOnBookOff Then
            frmBookOff.ChkComplete = cnChecked
        End If
        'Set Final Confirmation for Zero Time Booking processing
        If gblFunction = cnZeroTimeConfirm Then
            frmBatchTrack.ChkComplete = cnChecked
        End If
    End If
    Else '--TCR7117 --
    If Val(txtConfirmedYieldQty) + Val(txtConfirmedScrapQty) = Val(txtMaxYieldForOp.Text) Then '--TCR7117 --
    If gblFunction = cnBookOnBookOff Then '--TCR7117 --
       frmBookOff.ChkComplete = cnChecked
    End If
    If gblFunction = cnZeroTimeConfirm Then '--TCR7117 --
       frmBatchTrack.ChkComplete = cnChecked
    End If
    If gblFunction = cnWorkBook Then '--TCR7117 --
       frmBatchTrack.ChkComplete = cnChecked
    End If
    End If
    End If '--TCR7117 --
    'Update yield and scrap values in confirmation for BoBo processing
    If gblFunction = cnBookOnBookOff Then
    If sernp <> "" Then '--TCR7117 --
      BoBoTable.Value(gblBoBoCurrentRow, "YIELD") = YieldCount
      BoBoTable.Value(gblBoBoCurrentRow, "SCRAP") = ScrapCount
      Else '--TCR7117 --
      BoBoTable.Value(gblBoBoCurrentRow, "YIELD") = Val(txtConfirmedYieldQty)
      BoBoTable.Value(gblBoBoCurrentRow, "SCRAP") = Val(txtConfirmedScrapQty)
      End If '--TCR7117 --
    End If
    'Update yield and scrap values in confirmation for Zero Time Booking processing
    If gblFunction = cnZeroTimeConfirm Then
    If sernp <> "" Then '--TCR7117 --
        frmBatchTrack.txtYieldQty = YieldCount
        frmBatchTrack.txtScrapQty = ScrapCount
    Else '--TCR7117 --
        frmBatchTrack.txtYieldQty = Val(txtConfirmedYieldQty)
        frmBatchTrack.txtScrapQty = Val(txtConfirmedScrapQty)
    End If '--TCR7117 --
    End If
  'Update yield and scrap values in confirmation for Work Booking processing
   If gblFunction = cnWorkBook Then '--TCR7117 --
   If sernp <> "" Then '--TCR7117 --
        frmBatchTrack.txtYieldQty = YieldCount
        frmBatchTrack.txtScrapQty = ScrapCount
    Else '--TCR7117 --
        frmBatchTrack.txtYieldQty = Val(txtConfirmedYieldQty)
        frmBatchTrack.txtScrapQty = Val(txtConfirmedScrapQty)
    End If
   End If '--TCR7117 --
   
    IntPPSerNoRowCount = IntPPSerNoLocnTable.RowCount
    
    'All Serial numbers are already in IntPPSerNoLocnTable but now need updating
    'Find rows in yield flexigrid that are selected
    For Row = 1 To (flxYield.Rows - 1)
        flxYield.Row = Row
            
        If frmRecordYield.flxYield.CellBackColor = vbYellow Then
            'For each selected row update matching record in IntPPSerNoLocnTable
            For RowIndex = 1 To IntPPSerNoRowCount
   
                If Format((IntPPSerNoLocnTable.Value(RowIndex, "AUFNR")), "############") = frmRecordYield.txtOrderNumber Then
                    If IntPPSerNoLocnTable.Value(RowIndex, "SERNR") = frmRecordYield.flxYield.Text Then
                        IntPPSerNoLocnTable.Value(RowIndex, "RUECK") = frmRecordYield.txtConfNo
                        IntPPSerNoLocnTable(RowIndex, "ARBPL") = frmRecordYield.txtNextWorkCtr.Text
                        IntPPSerNoLocnTable(RowIndex, "ZDATE") = gblSAP.SAPFormatDate
                        IntPPSerNoLocnTable(RowIndex, "ZTIME") = gblSAP.SAPFormatTime
                        
                        Exit For
                    End If
                End If
            Next RowIndex
                
                
        Else
            'Ignore value
        End If
    Next
    
    'All Serial numbers are already in IntPPSerNoLocnTable but now need updating
    'Find rows in scrap flexigrid that are selected
    For Row = 1 To (flxScrap.Rows - 1)
        flxScrap.Row = Row
            
        If frmRecordYield.flxScrap.CellBackColor = vbYellow Then
            'For each selected row update matching record in IntPPSerNoLocnTable
            For RowIndex = 1 To IntPPSerNoRowCount
   
                If Format((IntPPSerNoLocnTable.Value(RowIndex, "AUFNR")), "############") = frmRecordYield.txtOrderNumber Then
                    If IntPPSerNoLocnTable.Value(RowIndex, "SERNR") = frmRecordYield.flxScrap.Text Then
                        IntPPSerNoLocnTable.Value(RowIndex, "RUECK") = frmRecordYield.txtConfNo
                        IntPPSerNoLocnTable(RowIndex, "ARBPL") = frmRecordYield.txtWorkCentre.Text
                        IntPPSerNoLocnTable(RowIndex, "SCRAPIND") = "X"
                        IntPPSerNoLocnTable(RowIndex, "ZDATE") = gblSAP.SAPFormatDate
                        IntPPSerNoLocnTable(RowIndex, "ZTIME") = gblSAP.SAPFormatTime
                        Exit For
                    End If
                End If
            Next RowIndex
                
                
        Else
            'Ignore value
        End If
    Next
    
    Me.Hide
    
    'Reset StartTime for Display Period
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    'Disable the timer
    Timer1.Enabled = True
    
    Exit Sub
    
   
ExitOnError:
    'Switch on the timer
    Timer1.Enabled = True
    Exit Sub
   
ErrorHandler:
    MsgBox "Error:" & Err.Description & " in " & Err.Source
    Exit Sub

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdYieldDeSelectAll_Click()

Dim SelectedRow As Integer

    'Reset StartTime for Display Period
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    'Disable the timer
    Timer1.Enabled = False

    
    For Row = 1 To (flxYield.Rows - 1)
        flxYield.Row = Row
        If frmRecordYield.flxYield.CellBackColor = vbYellow Then
            frmRecordYield.flxYield.CellBackColor = flxYield.BackColor
            frmRecordYield.txtYieldRecordsSelected.Text = frmRecordYield.txtYieldRecordsSelected.Text - 1
        End If
    Next
    

    'Reset StartTime for Display Period
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    'Disable the timer
    Timer1.Enabled = True
    
    Exit Sub
    
ExitOnError:
   Timer1.Enabled = True
      Exit Sub
    

End Sub

Private Sub cmdYieldSelectAll_Click()

Dim SelectedRow As Integer

    'Reset StartTime for Display Period
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    'Disable the timer
    Timer1.Enabled = False

    'Check whether there are any Scrap rows selected
    For Row = 1 To (flxYield.Rows - 1)
        flxScrap.Row = Row
        If flxScrap.CellBackColor = vbYellow Then
            'Do nothing
        Else
            RowSelected = True
            flxYield.Row = Row
            If frmRecordYield.flxYield.CellBackColor <> vbYellow Then
                frmRecordYield.flxYield.CellBackColor = vbYellow
                frmRecordYield.txtYieldRecordsSelected.Text = frmRecordYield.txtYieldRecordsSelected.Text + 1
            End If
        End If
    Next
        

    'Reset StartTime for Display Period
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    'Disable the timer
    Timer1.Enabled = True
    
    Exit Sub
    
ExitOnError:
   Timer1.Enabled = True
      Exit Sub
    

End Sub


Private Sub flxScrap_Click()

Dim SelectedRow As Integer

    'Reset StartTime for Display Period
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    'Disable the timer
    Timer1.Enabled = False

    'Initialise the Selected Row
    SelectedRow = flxScrap.RowSel
    'Dissallow selection if same flxYield row has already been selected
    flxYield.Row = SelectedRow
    If flxYield.CellBackColor = vbYellow Then
        'Do nothing
    Else
        RowSelected = True
    
    
        
        If flxScrap.CellBackColor = vbYellow Then
            flxScrap.CellBackColor = flxScrap.BackColor
            txtScrapRecordsSelected.Text = txtScrapRecordsSelected.Text - 1
        Else
            flxScrap.CellBackColor = vbYellow
            txtScrapRecordsSelected.Text = txtScrapRecordsSelected.Text + 1
        End If
    End If
 
    'Reset StartTime for Display Period
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    'Disable the timer
    Timer1.Enabled = True
    
    Exit Sub
    
ExitOnError:
   Timer1.Enabled = True
      Exit Sub
End Sub

Private Sub flxYield_Click()

Dim SelectedRow As Integer

    'Reset StartTime for Display Period
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    'Disable the timer
    Timer1.Enabled = False

    'Initialise the Selected Row
    SelectedRow = flxYield.RowSel
    'Dissallow selection if same flxScrap row has already been selected
    flxScrap.Row = SelectedRow
    If flxScrap.CellBackColor = vbYellow Then
        'Do nothing
    Else
        RowSelected = True
    
    
        
        If flxYield.CellBackColor = vbYellow Then
            flxYield.CellBackColor = flxYield.BackColor
            txtYieldRecordsSelected.Text = txtYieldRecordsSelected.Text - 1
        Else
            flxYield.CellBackColor = vbYellow
            txtYieldRecordsSelected.Text = txtYieldRecordsSelected.Text + 1
        End If
    End If
 
    'Reset StartTime for Display Period
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    'Disable the timer
    Timer1.Enabled = True
    
    Exit Sub
    
ExitOnError:
   Timer1.Enabled = True
      Exit Sub
End Sub

Private Sub Form_Activate()

   
    'After the grid has been loaded set the Start Time and Enable the Timer
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    Timer1.Enabled = True
       
   Exit Sub
   
   
ErrorHandler:
    MsgBox "Error:" & Err.Description & " in " & Err.Source
    Exit Sub

End Sub

Private Sub Form_Deactivate()
    
    gblYieldScrapFail = True
    Timer1.Enabled = False

End Sub

Private Sub Form_GotFocus()

  frmRecordYield.txtConfirmedYieldQty.SetFocus
  
End Sub

Private Sub Form_Load()

Dim Row As Long

    'Initialise the Yield FlexGrid
    'MsgBox "Initialising FlexGrid"
    flxYield.Clear
    flxYield.Rows = 1
    flxYield.Row = 0
    flxYield.RowSel = 0
    RowSelected = False
    
    'Set the Colums and Headings for the FlexGrid
    'nb: heights and widths are in twips (1440 twips per inch)
    Row = 0
    flxYield.RowHeight(Row) = 500
    flxYield.Cols = 1
    flxYield.ColWidth(0) = 2050
    
        
    flxYield.TextMatrix(Row, 0) = "Serial No"
    
        
    'Initialise the Scrap FlexGrid
    'MsgBox "Initialising FlexGrid"
    flxScrap.Clear
    flxScrap.Rows = 1
    flxScrap.Row = 0
    flxScrap.RowSel = 0
    RowSelected = False
    
    'Set the Colums and Headings for the FlexGrid
    'nb: heights and widths are in twips (1440 twips per inch)
    Row = 0
    flxScrap.RowHeight(Row) = 500
    flxScrap.Cols = 1
    flxScrap.ColWidth(0) = 2050
    
    
    
    flxScrap.TextMatrix(Row, 0) = "Serial No"
    
    

End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    gblYieldScrapFail = True
    Timer1.Enabled = False

End Sub

Private Sub Timer1_Timer()
    
    If Not OK2DisplayForm(False, gblPlantCfg.FormTimeOut) Then
        Timer1.Enabled = False
        gblYieldScrapFail = True
        Me.Hide
    End If
End Sub

