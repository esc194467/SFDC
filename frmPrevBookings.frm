VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrevBookings 
   Caption         =   "SFDC - Previous Bookings"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrevBookings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Previous Bookings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid flxPrevBookings 
         Height          =   4455
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Click on Booking to select for Cancellation"
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7858
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedCols       =   0
         BackColorBkg    =   12632256
         WordWrap        =   -1  'True
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
   Begin MSComCtl2.DTPicker ToTime 
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   62521346
      CurrentDate     =   36917
   End
   Begin MSComCtl2.DTPicker FromTime 
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   62521346
      CurrentDate     =   36917
   End
   Begin MSComCtl2.DTPicker ToDate 
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   62521345
      CurrentDate     =   36917
   End
   Begin MSComCtl2.DTPicker FromDate 
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   62521345
      CurrentDate     =   36917
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   8040
      Top             =   360
   End
   Begin VB.CommandButton cmdCancelBooking 
      Caption         =   "Cancel Booking"
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
      Left            =   7200
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdShowBookings 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Show Bookings"
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
      Left            =   4680
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox TotalTimeBooked 
      Enabled         =   0   'False
      Height          =   360
      Left            =   9960
      TabIndex        =   6
      Top             =   2100
      Width           =   1095
   End
   Begin VB.TextBox PersonName 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   480
      Width           =   4575
   End
   Begin VB.TextBox CheckNumber 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   9360
      Picture         =   "frmPrevBookings.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2340
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Bookings To"
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
      TabIndex        =   9
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label T 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Time Booked"
      Height          =   240
      Left            =   9960
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Hrs"
      Height          =   240
      Left            =   11160
      TabIndex        =   7
      Top             =   2160
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Person Name"
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
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CheckNumber"
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
      TabIndex        =   4
      Top             =   240
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bookings From"
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
      TabIndex        =   1
      Top             =   1080
      Width           =   1050
   End
End
Attribute VB_Name = "frmPrevBookings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Dummy As Boolean
Dim RowSelected As Boolean

Private Sub cmdCancelBooking_Click()

   Dim ListPosition As Integer
   Dim ConfNo As String
   Dim ConfCtr As String
   Dim ConfDate As Date
   Dim CancelledOK As String * 1
   Dim DialogResponse As Integer

   'Switch off the timer
   Timer1.Enabled = False

   'Validate the selected items
   If flxPrevBookings.Rows = 1 Then
      MsgBox "NO BOOKINGS HAVE BEEN DISPLAYED", vbExclamation, cnDialogTitleOpInfo
      'Reset StartTime for Display Period
      Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
      'Switch on the timer
      Timer1.Enabled = True
      Exit Sub
   ElseIf RowSelected = False Then
      MsgBox "PLEASE SELECT AN ITEM FROM THE LIST", vbExclamation, cnDialogTitleOpInfo
      'Reset StartTime for Display Period
      Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
      'Switch on the timer
      Timer1.Enabled = True
      Exit Sub
   ElseIf flxPrevBookings.RowSel - flxPrevBookings.Row <> 0 Then
      MsgBox "ONLY ONE ITEM CAN BE SELECTED", vbExclamation, cnDialogTitleOpInfo
      'Reset StartTime for Display Period
      Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
      'Switch on the timer
      Timer1.Enabled = True
      Exit Sub
   End If

   'Get appropriate values from SAP table
   ListPosition = flxPrevBookings.Row
   ConfNo = PrevBookingsTable.Value(ListPosition, "CONFIRM_NO")
   ConfCtr = PrevBookingsTable.Value(ListPosition, "CONF_CTR")
   ConfDate = PrevBookingsTable.Value(ListPosition, "POST_DATE")

   ' test booking being cancelled lies within the cancellation period
   If ConfDate < (Now - gblPlantCfg.CancelPeriod - 1) Then
      MsgBox "CANNOT CANCEL BOOKINGS MORE THAN " & gblPlantCfg.CancelPeriod & " DAYS OLD", vbExclamation, cnDialogTitleWorkBook
      'Reset StartTime for Display Period
      Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
      'Switch on the timer
      Timer1.Enabled = True
      Exit Sub
   End If

   DialogResponse = MsgBox("ARE YOU SURE YOU WISH TO CANCEL THE BOOKING?", vbYesNo, cnDialogTitleWorkBook)

   If DialogResponse <> vbYes Then
      'Reset StartTime for Display Period
      Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
      'Switch on the timer
      Timer1.Enabled = True
      Exit Sub
   End If

   'test SAP connection
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
      Exit Sub
   End If

   'Set pointer to hourglass while cancellation is processed by SAP
   Screen.MousePointer = vbHourglass

   'Set export parameters for SAPcancelBooking

   SAPCancelBooking.Exports("CONFIRMATION_NUMBER") = ConfNo
   SAPCancelBooking.Exports("CONFIRMATION_COUNTER") = ConfCtr
   SAPCancelBooking.Exports("SAP_DEBUG") = gblInstallation.SAPDebug

   If SAPCancelBooking.Call = False Then
      Screen.MousePointer = vbDefault
      MsgBox SAPCancelBooking.Exception, vbOKOnly, cnDialogTitleWorkBook
      'Reset StartTime for Display Period
      Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
      'Switch on the timer
      Timer1.Enabled = True
      Exit Sub
   Else
      Screen.MousePointer = vbDefault
      CancelledOK = SAPCancelBooking.Imports("SUCCESSFULLY_CANCELED")

      If CancelledOK = "X" Then
         MsgBox "CANCELLATION SUCCESSFUL", vbOKOnly, cnDialogTitleWorkBook
         'Refresh list by call to SAP function module
         cmdShowBookings_Click
         'Reset StartTime for Display Period
         Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
         'Switch on the timer
         Timer1.Enabled = True

      Else

         gblSAP.ErrorMessage = SAPCancelBooking.Imports("ERROR_MESSAGE")
         MsgBox "CANCELLATION FAILED - " & gblSAP.ErrorMessage, vbExclamation, cnDialogTitleWorkBook
         'Reset StartTime for Display Period
         Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
         'Switch on the timer
         Timer1.Enabled = True

      End If
   End If

End Sub

Private Sub cmdShowBookings_Click()

   Dim RowCount As Integer
   Dim i As Integer
   Dim Row As Long
   Dim ConfirmNo As String * 10
   Dim OrderNo As String * 12
   Dim OpNo As String * 5
   Dim OpDesc As String * 41
   Dim PostDate As String * 10
   Dim PostTime As String * 8
   Dim TimeBooked As String * 8
   Dim Final As String * 2
   Dim Units As String
   Dim BookingStr As String
   Dim SAPFromDate As Date
   Dim SAPtoDate As Date
   Dim TotalTime As Single
   Dim ThisBooking As Single
   Dim StartTime As String * 8
   Dim EndTime As String * 8
   Dim ConfInd As String * 1

   On Error GoTo ErrorHandler

   'Disable the timer
   Timer1.Enabled = False

   'Convert date fields into date variables
   SAPFromDate = CDate(FromDate)
   SAPtoDate = CDate(ToDate)

   'Test for valid time parameters
   If ToTime > "24:00" Or FromTime > "24:00" Then
      GoTo ExitOnError
   End If

   If SAPtoDate < SAPFromDate Then
      MsgBox "FROM DATE SHOULD BE LESS THAN TO DATE", vbExclamation, cnDialogTitleWorkBook
      GoTo ExitOnError
   End If

   If SAPtoDate = SAPFromDate And ToTime < FromTime Then
      MsgBox "FROM TIME SHOULD BE LESS THAN TO TIME", vbExclamation, cnDialogTitleWorkBook
      GoTo ExitOnError
   End If

   'test SAP connection
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
      GoTo ExitOnError
   End If

   'Set pointer to hourglass during SAP processing
   Screen.MousePointer = vbHourglass

   'initialise TotalTime
   TotalTime = 0

   'Set export parameters fo SAPReadPrevBookings
   StartTime = Format(FromTime, "HHMMSS")
   EndTime = Format(ToTime, "HHMMSS")
   SAPReadPrevBookings.Exports("I_PERSON") = CheckNumber
   SAPReadPrevBookings.Exports("I_START_DATE") = Format(SAPFromDate, "YYYYMMDD")
   SAPReadPrevBookings.Exports("I_END_DATE") = Format(SAPtoDate, "YYYYMMDD")
   SAPReadPrevBookings.Exports("I_END_TIME") = EndTime
   SAPReadPrevBookings.Exports("I_START_TIME") = StartTime

   PrevBookingsTable.FreeTable 'remove existing data in table

   'Clear the Flexgrid and reset the Row Selected flags
   flxPrevBookings.Rows = 1
   flxPrevBookings.Row = 0
   flxPrevBookings.RowSel = 0
   RowSelected = False

   'Call the RFC
   If SAPReadPrevBookings.Call = False Then
      'set pointer to standard
      Screen.MousePointer = vbDefault

      MsgBox SAPReadPrevBookings.Exception, vbExclamation, cnDialogTitleLogon
      'Exit Sub
   End If

   RowCount = PrevBookingsTable.RowCount

   For i = 1 To RowCount

      ConfirmNo = Format(PrevBookingsTable.Value(i, "CONFIRM_NO"), "#########")
      OrderNo = Format(PrevBookingsTable.Value(i, "ORDER_NO"), "#########")
      OpNo = PrevBookingsTable.Value(i, "OP_NO")
      OpDesc = PrevBookingsTable.Value(i, "OP_DESC")
      PostDate = PrevBookingsTable.Value(i, "POST_DATE")
      PostTime = PrevBookingsTable.Value(i, "END_TIME")
      RSet TimeBooked = Format(Val(PrevBookingsTable.Value(i, "WORKBKNG")), "0.00")
      Units = PrevBookingsTable.Value(i, "WORK_UNITS")
      Final = PrevBookingsTable.Value(i, "FINAL_CONF")
      ConfInd = PrevBookingsTable.Value(i, "CONF_DESC")

      'Check for 1 PicoSecond Bookings and convert to Zero Minutes
      'for display purposes
      If TimeBooked = 1 And Units = "PS" Then
         RSet TimeBooked = Format(0, "0.00")
         Units = "MIN"
      End If

      'Check for automatic final confirmation and remove the finally
      'confirmed flag
      If ConfInd = Chr(2) Then
         Final = ""
      End If

      Row = flxPrevBookings.Rows
      flxPrevBookings.AddItem ("")
      flxPrevBookings.TextMatrix(Row, 0) = ConfirmNo
      flxPrevBookings.TextMatrix(Row, 1) = OrderNo
      flxPrevBookings.TextMatrix(Row, 2) = OpNo
      flxPrevBookings.TextMatrix(Row, 3) = OpDesc
      flxPrevBookings.TextMatrix(Row, 4) = Final
      flxPrevBookings.TextMatrix(Row, 5) = PostDate
      flxPrevBookings.TextMatrix(Row, 6) = PostTime
      flxPrevBookings.TextMatrix(Row, 7) = TimeBooked
      flxPrevBookings.TextMatrix(Row, 8) = Units

      Select Case Units
         Case "MIN"
            ThisBooking = Val(TimeBooked) / 60
         Case "H"
            ThisBooking = Val(TimeBooked)
         Case "D"
            ThisBooking = Val(TimeBooked) * 24
         Case Else
            ThisBooking = 0
      End Select

      TotalTime = TotalTime + ThisBooking

   Next i

   TotalTimeBooked = Format(TotalTime, "00000.00")

   'Set pointer to standard
   Screen.MousePointer = vbDefault

   GoTo ExitOnError

   Exit Sub

ExitOnError:
   'Reset StartTime for Display Period
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   'Switch on the timer
   Timer1.Enabled = True
   Exit Sub

ErrorHandler:
   MsgBox "INVALID DATE", vbExclamation, cnDialogTitleLogon
   Exit Sub

End Sub

Private Sub flxPrevBookings_Click()

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub flxPrevBookings_Scroll()

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub flxPrevBookings_SelChange()

   RowSelected = True

End Sub

Private Sub Form_Activate()

   'Set the Start Time and Enable the Timer
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   Timer1.Enabled = True

   If gblUser.PrevBookingsEndTime = "00:00:00" Or gblUser.PrevBookingsEndTime = "24:00:00" Then
      gblUser.PrevBookingsEndTime = "23:59:59"
   End If
   
   If gblUser.PrevBookingsStartTime = "00:00:00" Then
      'frmPrevBookings.FromDate = Format(Now, "dd/mm/yy")
      'frmPrevBookings.FromDate = Format(Now, gblLocalDateFormat)
      frmPrevBookings.FromDate = DateValue(Now)
      'frmPrevBookings.FromDate = Date
   Else
      'frmPrevBookings.FromDate = Format(Now - 1, "dd/mm/yy")
      'frmPrevBookings.FromDate = Format(Now - 1, gblLocalDateFormat)
      frmPrevBookings.FromDate = DateValue(Now - 1)
      'frmPrevBookings.FromDate = (Date - 1)
   End If

   'Set appropriate field values
   With frmPrevBookings
      .FromTime = gblUser.PrevBookingsStartTime
      '.ToDate = Format(Now, "dd/mm/yy")
      '.ToDate = Format(Now, gblLocalDateFormat)
      .ToDate = DateValue(Now)
      '.ToDate = Date
      .ToTime = gblUser.PrevBookingsEndTime
      .PersonName = gblUser.PersName
      .CheckNumber = gblUser.ClockNumber
      .TotalTimeBooked = ""
   End With

   'Reset the FlexGrid
   flxPrevBookings.Rows = 1
   flxPrevBookings.Row = 0
   flxPrevBookings.RowSel = 0
   RowSelected = False

End Sub

Private Sub Form_Deactivate()

   'Disable the Timer
   Timer1.Enabled = False

End Sub

Private Sub Form_Load()

   Dim Row As Long

   'Initialise the FlexGrid
   flxPrevBookings.Clear
   flxPrevBookings.Rows = 1
   flxPrevBookings.Row = 0
   flxPrevBookings.RowSel = 0
   RowSelected = False

   'Set the Colums and Headings for the FlexGrid
   'nb: heights and widths are in twips (1440 twips per inch)
   Row = 0
   flxPrevBookings.RowHeight(Row) = 500
   flxPrevBookings.Cols = 9
   flxPrevBookings.ColWidth(0) = 1200
   flxPrevBookings.ColWidth(1) = 1200
   flxPrevBookings.ColWidth(2) = 500
   flxPrevBookings.ColWidth(3) = 4500
   flxPrevBookings.ColWidth(4) = 500
   flxPrevBookings.ColWidth(5) = 1000
   flxPrevBookings.ColWidth(6) = 700
   flxPrevBookings.ColWidth(7) = 1100
   flxPrevBookings.ColAlignment(7) = flexAlignRightCenter
   flxPrevBookings.ColWidth(8) = 600

   flxPrevBookings.TextMatrix(Row, 0) = "Confirmation No"
   flxPrevBookings.TextMatrix(Row, 1) = "Order No"
   flxPrevBookings.TextMatrix(Row, 2) = "Op No"
   flxPrevBookings.TextMatrix(Row, 3) = "Op Description"
   flxPrevBookings.TextMatrix(Row, 4) = "Final Conf"
   flxPrevBookings.TextMatrix(Row, 5) = "Post Date"
   flxPrevBookings.TextMatrix(Row, 6) = "Post Time"
   flxPrevBookings.TextMatrix(Row, 7) = "Time Booked"
   flxPrevBookings.TextMatrix(Row, 8) = "Units"
   
   'Disable Cancellation button if flag is set from the Plant Configuration
   If gblPlantCfg.CancellationRestricted = "X" Then
      frmPrevBookings.cmdCancelBooking.Enabled = False
      frmPrevBookings.flxPrevBookings.ToolTipText = ""
   Else
      frmPrevBookings.cmdCancelBooking.Enabled = True
      frmPrevBookings.flxPrevBookings.ToolTipText = "Click on Booking to select for Cancellation"
   End If
   

End Sub

Private Sub FromDate_DropDown()

   'Set the Start Time
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   'Disable Timer
   Timer1.Enabled = False

End Sub


Private Sub FromDate_LostFocus()

   'Enable the Timer
   Timer1.Enabled = True

End Sub

Private Sub FromTime_Change()

   'Reset the Start Time
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub FromTime_Click()

   'Reset the Start Time
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub FromTime_KeyPress(KeyAscii As Integer)

   'Reset the Start Time
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub Timer1_Timer()

   If Not OK2DisplayForm(False, gblPlantCfg.FormTimeOut) Then
      Timer1.Enabled = False
      Me.Hide
   End If

End Sub

Private Sub ToDate_DropDown()

   'Set the Start Time
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   'Disable Timer
   Timer1.Enabled = False

End Sub


Private Sub ToDate_LostFocus()

   'Enable the Timer
   Timer1.Enabled = True

End Sub

Private Sub ToTime_Change()

   'Reset the Start Time
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub ToTime_Click()

   'Reset the Start Time
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub ToTime_KeyPress(KeyAscii As Integer)

   'Reset the Start Time
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

