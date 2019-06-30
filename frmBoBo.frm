VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBoBo 
   Caption         =   "SFDC - Book On / Book Off"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBoBo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3600
      Top             =   1200
   End
   Begin VB.CommandButton cmdRemovefromList 
      Caption         =   "Remove from List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdUpdateBoBo 
      BackColor       =   &H80000014&
      Caption         =   "Update  Bookings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Currently Booked On Jobs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid flxBoBo 
         Height          =   3975
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "CLICK on Job to Select/Deselect for BOOK OFF"
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
         BackColorBkg    =   12632256
         WordWrap        =   -1  'True
         AllowBigSelection=   -1  'True
         FocusRect       =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox ConfirmationNumber 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox CheckNumber 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox PersonName 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   9120
      Picture         =   "frmBoBo.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2340
   End
   Begin VB.Label C 
      AutoSize        =   -1  'True
      Caption         =   "BOOK ON Confirmation No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   1080
      Width           =   1905
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Check Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   5
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "frmBoBo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bbOrder As OrderInfo
Dim bbOp As OpInfo
Dim bbOrdTypeCfg As OrderTypeCfgInfo

Dim Dummy As Boolean

Dim ListPosition As Integer
Dim DialogResponse As Integer
Dim Row As Integer
Dim RowSelected As Boolean
Dim BoBoRowCount As Integer
Dim WarningFlagSet As Boolean

Private Sub cmdRemovefromList_Click()

   'Reset StartTime and disable timer
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   Timer1.Enabled = False

   'Validate the selected items
   If flxBoBo.Rows = 1 Then
      MsgBox "NO BOOKINGS ARE AVAILABLE FOR REMOVAL", vbExclamation
      GoTo ExitOnError
   ElseIf BoBoTable.Value(flxBoBo.RowSel, "ACTION") <> cnBookOnAction Then
      MsgBox "CAN ONLY REMOVE NEW ROWS", vbExclamation
      GoTo ExitOnError
   ElseIf RowSelected = False Then
      MsgBox "PLEASE SELECT AN ITEM FROM THE LIST", vbExclamation
      GoTo ExitOnError
   ElseIf flxBoBo.Rows = 2 Then
      RowSelected = False
      flxBoBo.Rows = 1
      flxBoBo.Row = 0
      flxBoBo.RowSel = 0
      'Remove corresponding row from BoBo Table
      BoBoTable.DeleteRow (1)
      Exit Sub
   ElseIf flxBoBo.RowSel - flxBoBo.Row <> 0 Then
      MsgBox "ONLY ONE ITEM CAN BE SELECTED", vbExclamation
      GoTo ExitOnError
   End If

   ListPosition = flxBoBo.RowSel
   'Remove row from flxBoBo
   flxBoBo.RemoveItem (ListPosition)
   'Remove corresponding row from BoBo Table
   BoBoTable.DeleteRow (ListPosition)

   'Enable the timer and reset the time on successful exit
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   Timer1.Enabled = True

   Exit Sub

ExitOnError:
   Timer1.Enabled = True
   Exit Sub

End Sub

Private Sub cmdUpdateBoBo_Click()

   Dim tmpBoBoTable As SAPTableFactoryCtrl.Table
   Dim tmpSerNoLocnTable As SAPTableFactoryCtrl.Table

  ' On Error GoTo ErrorHandler

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   'Switch off the timer
   Timer1.Enabled = False

   'Check that there are rows to process
   If flxBoBo.Rows < 2 Then
      MsgBox "THERE ARE NO BOOKINGS TO UPDATE", vbExclamation
      GoTo ExitOnError
   End If

   'Test SAP connection before proceeding
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
      GoTo ExitOnError
   End If

   'Loop thru the BoBo table to determine if there any
   'Long Running Open Bookings LROBs (rows highlighted in red)
   'and if there are and the action is not set to Book Off
   'generate error message and exit
   If WarningFlagSet Then
      For Row = 1 To BoBoTable.RowCount
         If BoBoTable.Value(Row, "WARNING") = cnSAPTrue And _
            BoBoTable.Value(Row, "ACTION") <> cnBookOffAction Then
            MsgBox "PLEASE BOOK-OFF ALL JOBS MARKED IN RED" & _
               Chr(10) & "CAREFULLY CONSIDER THE BOOK-OFF DATE/TIME", vbExclamation
            GoTo ExitOnError
         End If
      Next Row
   End If

   'Loop thru the BoBo table to determine if there any Book Offs
   'and if there are call BookOff form passing the starting row
   'in a global variable
   gblBoBoCurrentRow = 0
   For Row = 1 To BoBoTable.RowCount
      If BoBoTable.Value(Row, "ACTION") = cnBookOffAction Then
         gblBoBoCurrentRow = Row
         Exit For
      End If
   Next Row

   If gblBoBoCurrentRow > 0 Then
      'Clear the internal serial number table (for PP Orders)
      IntPPSerNoLocnTable.FreeTable
      
      frmBookOff.Show vbModal
      'Ensure timer in BookOff form is disabled on return
      frmBookOff.Timer1.Enabled = False
   Else
      gblBatchTrackOK = True
   End If

   'Test that all BookOffs were updated successfully and exit if not
   If gblBatchTrackOK = False Then
      GoTo ExitOnError
   End If

   'Set export parameters fo SAPMakeBoBo
   SAPMakeBoBo.Exports("PERSON") = gblUser.ClockNumber
   SAPMakeBoBo.Exports("TIMEZONE") = gblPlantCfg.TimeZone
   SAPMakeBoBo.Exports("SAP_DEBUG") = gblInstallation.SAPDebug

   'Define a temp table variable
   Set tmpBoBoTable = SAPMakeBoBo.Tables.Item("BOBO")
   'Clear the temp table
   tmpBoBoTable.FreeTable
   'Copy the data from the public table to the temp table
   tmpBoBoTable.data = BoBoTable.data
   
   'Define a temp table variable for the serial numbers
   Set tmpSerNoLocnTable = SAPMakeBoBo.Tables.Item("ZSERNOLOCN")
   'Clear the temp table
   tmpSerNoLocnTable.FreeTable
   
   'Copy the data from the public table to the temp table
   If IntPPSerNoLocnTable.RowCount <> 0 Then
      tmpSerNoLocnTable.data = IntPPSerNoLocnTable.data
   End If
   
   'Set the hourglass on and make the call
   Screen.MousePointer = vbHourglass

   'Call the RFC
   If SAPMakeBoBo.Call = False Then
      'set pointer to standard
      Screen.MousePointer = vbDefault
      MsgBox SAPMakeBoBo.Exception, vbExclamation, cnDialogTitleLogon
      GoTo ExitOnError
   Else
      Screen.MousePointer = vbDefault
      frmBoBo.Timer1.Enabled = False
      frmBoBo.Hide
      MsgBox "SAP UPDATES COMPLETED SUCCESSFULLY", vbExclamation
      
   End If

   Exit Sub

ExitOnError:
   'Switch on the timer
   Timer1.Enabled = True
   Exit Sub

ErrorHandler:
   MsgBox "Error:" & Err.Description & " in " & Err.Source
   Exit Sub

End Sub


Private Sub ConfirmationNumber_KeyPress(KeyAscii As Integer)


    If KeyAscii = 13 Then
    
        'Reset StartTime
        Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
        'Disable the Timer
        Timer1.Enabled = False
    
        'set focus back in this control
        ConfirmationNumber.SetFocus
    
        'test SAP connection
        If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
            Exit Sub
        End If
        
        'Test for Long Running Book Ons before Booking On any other Jobs
        If WarningFlagSet Then
            MsgBox "PLEASE BOOK OFF ALL JOBS MARKED IN RED" & Chr(10) & _
                   "BEFORE BOOKING ON ANY OTHER JOBS" & _
                    Chr(10) & "CAREFULLY CONSIDER THE BOOK-OFF DATE/TIME", vbExclamation
            GoTo ExitOnError
        End If
        
        'Check for existing row with this Confirmation No
        If flxBoBo.Rows > 1 Then
            For Row = 1 To flxBoBo.Rows - 1
                If Val(flxBoBo.TextMatrix(Row, 0)) = Val(ConfirmationNumber) Then
                    MsgBox "CANNOT BOOK ON TO THE SAME ACTIVITY TWICE", vbExclamation
                    GoTo ExitOnError
                End If
            Next Row
        End If
        
        'Set the Confirmation Number in the Op UDT
        bbOp.ConfNo = ConfirmationNumber
        'then call the function
        If Not ReadOrderOp(bbOrder, bbOp, bbOrdTypeCfg) Then
            GoTo ExitOnError
        End If
        
        'Check Op can be confirmed via Status checks on Order and Op
        If Not OpisOK2Confirm(bbOrdTypeCfg, bbOp) Then
            GoTo ExitOnError
        End If
             
        'Test for Op Plant not the same as Person Plant
        'and warn user
        ' 1/11/99 change to error - requested by Andy Dickinson
        If gblUser.Plant <> bbOp.Plant Then
            DialogResponse = MsgBox("PLANNED WORKCENTRE IS NOT IN YOUR PLANT", _
            vbExclamation, cnDialogTitleWorkBook)
            GoTo ExitOnError
        End If
        

        'Add a row to the flex grid nb. rows indexed from zero
        'row zero is the column headings
        Row = flxBoBo.Rows
        flxBoBo.AddItem ("")
        flxBoBo.TextMatrix(Row, 0) = Format(bbOp.ConfNo, "##########")
        flxBoBo.TextMatrix(Row, 1) = Format(bbOrder.Number, "############")
        flxBoBo.TextMatrix(Row, 2) = bbOp.Number
        flxBoBo.TextMatrix(Row, 3) = bbOp.WorkCentre
        flxBoBo.TextMatrix(Row, 4) = bbOp.Desc
        
        'Add a row to the BoBo table with appropriate values
        'nb. rows indexed from 1
        BoBoTable.AppendRow
        Row = BoBoTable.RowCount
        BoBoTable.Value(Row, "MANDT") = gblSAP.Client
        BoBoTable.Value(Row, "PERS_NO") = gblUser.ClockNumber
        BoBoTable.Value(Row, "CONF_NO") = bbOp.ConfNo
        BoBoTable.Value(Row, "UN_WORK") = bbOp.WorkUnits
        BoBoTable.Value(Row, "ACTION") = cnBookOnAction
        BoBoTable.Value(Row, "ORDER_NO") = bbOrder.Number
        BoBoTable.Value(Row, "ACTIVITY") = bbOp.Number
        BoBoTable.Value(Row, "APPLICATION") = bbOrder.Category
        BoBoTable.Value(Row, "AUFPL") = bbOrder.Aufpl
        BoBoTable.Value(Row, "APLZL") = bbOp.Aplzl
        BoBoTable.Value(Row, "PLANT") = bbOp.Plant
        BoBoTable.Value(Row, "WORK_CNTR") = bbOp.WorkCentre
        BoBoTable.Value(Row, "ACT_TYPE") = bbOp.ActivityType
        BoBoTable.Value(Row, "ARBID") = bbOp.Arbid
        BoBoTable.Value(Row, "ORD_DESC") = bbOrder.Desc
        BoBoTable.Value(Row, "ORD_TYPE") = bbOrder.Type
        
       
        ' reset confirmation number
        ConfirmationNumber = ""
        
    End If
    
    'Enable the timer and reset the time on successful exit
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    Timer1.Enabled = True
    
    Exit Sub
    
ExitOnError:
    ConfirmationNumber = ""
    Timer1.Enabled = True
    Exit Sub
            

End Sub

Private Sub flxBoBo_Click()

Dim SelectedRow As Integer

    'Reset StartTime for Display Period
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    'Disable the timer
    Timer1.Enabled = False
    
    'If there no rows in the list then exit
    If BoBoTable.RowCount = 0 Then
        GoTo ExitOnError
    End If

    'Initialise the Selected Row
    SelectedRow = flxBoBo.RowSel
    RowSelected = True
    
    'Changes to the following IF clause for ver 3.0.9 in response to
    'HAESL PPR48 10/12/03 Steve Lynam

    If BoBoTable.Value(SelectedRow, "ACTION") = cnBookOnAction Then
        MsgBox "CANNOT SELECT NEW ROWS FOR BOOK OFF", vbExclamation
        GoTo ExitOnError
        
    ElseIf BoBoTable.Value(SelectedRow, "ACTION") = cnBookOffAction Then
        'Set the colour of the selected row to the default
        flxBoBo.CellBackColor = flxBoBo.BackColor
        'Set the Action of the table row to blank
        BoBoTable.Value(SelectedRow, "ACTION") = cnSAPFalse
    
    Else
        'Set the selected row to BookOff
        flxBoBo.CellBackColor = vbYellow
        'Set the action of the table row to "O" - Book Off
        BoBoTable.Value(SelectedRow, "ACTION") = cnBookOffAction
    
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

    Dim RowIndex As Integer
    Dim Column As Integer
    Dim bbRow As BoBoRow
    
    'Initialise the flag for BOBO rows with the Warning Flag set
    WarningFlagSet = False
    
    'Populate field values as appropriate
    With frmBoBo
        .CheckNumber = gblUser.ClockNumber
        .PersonName = gblUser.PersName
    End With
    
    'Reset the Flexgrid
    flxBoBo.Rows = 1
    flxBoBo.Row = 0
    flxBoBo.RowSel = 0
    RowSelected = False
    

   'Set export parameters for SAPReadbobo (reading WIP values)
   SAPReadBoBo.Exports("PERSON") = gblUser.ClockNumber
   SAPReadBoBo.Exports("TIMEZONE") = gblPlantCfg.TimeZone
   SAPReadBoBo.Exports("STATUS") = cnBoBoWIPStatus
   SAPReadBoBo.Exports("SAP_DEBUG") = gblInstallation.SAPDebug
   
   
   'Clear the table prior to calling func module
   BoBoTable.FreeTable
   'BoBoTable.Refresh
   

   'Call the RFC
   If SAPReadBoBo.Call = False Then
        'set pointer to standard
        Screen.MousePointer = vbDefault
        MsgBox SAPReadBoBo.Exception, vbExclamation, cnDialogTitleLogon
        Me.Hide
        Exit Sub
   End If
   
   
   BoBoRowCount = BoBoTable.RowCount
      
   For RowIndex = 1 To BoBoRowCount
   
      bbRow.ConfNo = Format(BoBoTable.Value(RowIndex, "CONF_NO"), "#########")
      bbRow.OrderNo = Format(BoBoTable.Value(RowIndex, "ORDER_NO"), "#########")
      bbRow.OpNumber = BoBoTable.Value(RowIndex, "ACTIVITY")
      bbRow.WorkCentre = BoBoTable.Value(RowIndex, "WORK_CNTR")
      bbRow.OnDate = BoBoTable.Value(RowIndex, "ON_DATE")
      bbRow.OnTime = BoBoTable.Value(RowIndex, "ON_TIME")
      bbRow.OpDesc = BoBoTable.Value(RowIndex, "OP_DESC")
      bbRow.Warning = BoBoTable.Value(RowIndex, "WARNING")
      bbRow.OrderCategory = BoBoTable.Value(RowIndex, "APPLICATION")
      
      Row = flxBoBo.Rows
      flxBoBo.AddItem ("")
      
      'WarningFlag = "X"
      If bbRow.Warning = cnSAPTrue Then
        WarningFlagSet = True
        flxBoBo.Row = Row
        For Column = 5 To 6
            flxBoBo.Col = Column
            flxBoBo.CellBackColor = vbRed
        Next Column
      End If
      
      flxBoBo.TextMatrix(Row, 0) = Format(bbRow.ConfNo, "##########")
      flxBoBo.TextMatrix(Row, 1) = Format(bbRow.OrderNo, "############")
      flxBoBo.TextMatrix(Row, 2) = bbRow.OpNumber
      flxBoBo.TextMatrix(Row, 3) = bbRow.WorkCentre
      flxBoBo.TextMatrix(Row, 4) = bbRow.OpDesc
      flxBoBo.TextMatrix(Row, 5) = bbRow.OnDate
      flxBoBo.TextMatrix(Row, 6) = bbRow.OnTime
      
      
       
   Next RowIndex
   
    'After the grid has been loaded set the Start Time and Enable the Timer
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
    Timer1.Enabled = True
       
   Exit Sub
   
   
ErrorHandler:
    MsgBox "Error:" & Err.Description & " in " & Err.Source
    Exit Sub

End Sub



Private Sub Form_GotFocus()

    frmBoBo.ConfirmationNumber.SetFocus

End Sub


Private Sub Form_Load()
    Dim Row As Long

    'Initialise the FlexGrid
    'MsgBox "Initialising FlexGrid"
    flxBoBo.Clear
    flxBoBo.Rows = 1
    flxBoBo.Row = 0
    flxBoBo.RowSel = 0
    RowSelected = False
    
    'Set the Colums and Headings for the FlexGrid
    'nb: heights and widths are in twips (1440 twips per inch)
    Row = 0
    flxBoBo.RowHeight(Row) = 500
    flxBoBo.Cols = 9
    flxBoBo.ColWidth(0) = 1400
    flxBoBo.ColWidth(1) = 1200
    flxBoBo.ColWidth(2) = 500
    flxBoBo.ColWidth(3) = 800
    flxBoBo.ColWidth(4) = 5000
    flxBoBo.ColWidth(5) = 1000
    flxBoBo.ColWidth(6) = 800
    flxBoBo.ColWidth(7) = 0  'Hide this Column
    flxBoBo.ColWidth(8) = 0  'Hide this Column
    
    flxBoBo.TextMatrix(Row, 0) = "Confirmation No"
    flxBoBo.TextMatrix(Row, 1) = "Order No"
    flxBoBo.TextMatrix(Row, 2) = "Op No"
    flxBoBo.TextMatrix(Row, 3) = "Work Centre"
    flxBoBo.TextMatrix(Row, 4) = "Op Description"
    flxBoBo.TextMatrix(Row, 5) = "BookOn Date"
    flxBoBo.TextMatrix(Row, 6) = "BookOn Time"
    flxBoBo.TextMatrix(Row, 7) = "Aufpl"
    flxBoBo.TextMatrix(Row, 8) = "Aplzl"
 
End Sub

Private Sub Timer1_Timer()
    
    If Not OK2DisplayForm(False, gblPlantCfg.FormTimeOut) Then
        Timer1.Enabled = False
        Me.Hide
    End If
    
End Sub

Private Sub Form_Deactivate()

    
    Timer1.Enabled = False
    
    'Unlock the appropriate rows in ZBOBO
    'Note this should be done by proc DoBoBo in frmLogon
    'however it is not working in the .exe
    SAPUnlockBoBo.Exports("PERSON") = gblUser.ClockNumber
    
        
    'Call the RFC
    If SAPUnlockBoBo.Call = False Then
         MsgBox SAPReadBoBo.Exception, vbExclamation, cnDialogTitleLogon
    End If

End Sub

