VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMultiBooking 
   Caption         =   " SFDC - Multi Bookings"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   11130
   Icon            =   "frmMultiBooking.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Operations Being Booked"
      Height          =   4575
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid flxMultiBookings 
         Height          =   3975
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   1
         Cols            =   6
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5280
      Top             =   1080
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
      Left            =   2160
      TabIndex        =   9
      Top             =   480
      Width           =   4335
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
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin MSMask.MaskEdBox TotalTime 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000.00"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox WorkUnits 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmMultiBooking.frx":0442
      Left            =   4200
      List            =   "frmMultiBooking.frx":044C
      TabIndex        =   2
      Text            =   "MIN"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox ConfirmationNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
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
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdMakeBookings 
      Caption         =   "Make Bookings"
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
      Left            =   4440
      TabIndex        =   3
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   8760
      Picture         =   "frmMultiBooking.frx":0458
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2340
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   2160
      TabIndex        =   11
      Top             =   240
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Check Number"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Units"
      Height          =   195
      Left            =   4200
      TabIndex        =   6
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Time"
      Height          =   195
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   750
   End
   Begin VB.Label C 
      AutoSize        =   -1  'True
      Caption         =   "Confirmation No"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1125
   End
End
Attribute VB_Name = "frmMultiBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbOp As OpInfo
Dim mbOrder As OrderInfo
Dim mbOrdTypeCfg As OrderTypeCfgInfo

Dim Dummy As Boolean

Dim Row As Integer
Dim RowSelected As Boolean

Private Sub cmdMakeBookings_Click()

   Dim Conf As ConfInfo

   Dim tmpPostDateTime As String * 18
   Dim PostDateTime As Date
   Dim Msg As String
   Dim oldOrderNo As String * 12
   Dim tmpOpTime As Double
   Dim tmpTotalTime As Double
   Dim Cnt As Integer
   Dim RowCount As Integer
   Dim Start As Single

   On Error GoTo ErrorHandler

   'Switch off the timer
   Timer1.Enabled = False

   'test SAP connection
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
      GoTo ExitOnError
   End If

   'check there are some bookings to make
   If flxMultiBookings.Rows = 1 Then
      MsgBox "THERE ARE NO BOOKINGS TO MAKE", vbExclamation, cnDialogTitleWorkBook
      GoTo ExitOnError
   End If

   'Establish time for individual bookings
   If WorkUnits = "H" Then
      Conf.WorkUnits = "MIN"
      tmpTotalTime = TotalTime * 60
   Else
      Conf.WorkUnits = WorkUnits
      tmpTotalTime = TotalTime
   End If
   tmpOpTime = tmpTotalTime / (flxMultiBookings.Rows - 1)
   Conf.ActWork = Format(tmpOpTime, "######.00")

   'Determine if all Ops should be set to completed
   If frmBatchTrack.ChkComplete = cnChecked Then
      Conf.Complete = cnSAPTrue
      Conf.FinalConf = cnSAPTrue
   Else
      Conf.Complete = cnSAPFalse
      Conf.FinalConf = cnSAPFalse
   End If

   RowCount = flxMultiBookings.Rows

   'Loop thru the bookings
   For Cnt = 1 To RowCount - 1

      'Select the first row in list
      flxMultiBookings.Row = 1
      flxMultiBookings.Col = 0
      flxMultiBookings.RowSel = 1
      flxMultiBookings.ColSel = 6

      'Determine Values to be passed to SAP from FlexGrid
      mbOrder.Number = Format(flxMultiBookings.TextMatrix(1, 1), "000000000000")
      mbOrder.Category = flxMultiBookings.TextMatrix(1, 6)
      mbOrder.Aufpl = flxMultiBookings.TextMatrix(1, 7)
      mbOp.Number = flxMultiBookings.TextMatrix(1, 2)
      mbOp.Aplzl = flxMultiBookings.TextMatrix(1, 8)
      Conf.WorkCentre = flxMultiBookings.TextMatrix(1, 3)
      Conf.ActType = flxMultiBookings.TextMatrix(1, 5)


      'Call Batch Tracking form for each row
      frmMultiBatchTrack.Show vbModal
      Me.Refresh

      'Test BatchTrack form was exited correctly
      If gblBatchTrackOK = False Then
         Exit Sub
      End If

      'Set Backflush flag
      If frmMultiBatchTrack.ChkPreOp = cnChecked Then
         Conf.BackFlush = cnSAPFalse
      Else
         Conf.BackFlush = cnSAPTrue
      End If
      
      'Set Final Confirmation flag
      If frmMultiBatchTrack.ChkComplete = cnChecked Then
         Conf.FinalConf = cnSAPTrue
      Else
         Conf.FinalConf = cnSAPFalse
      End If

      'Set the remaining Confirmation values
      With Conf
         .Plant = gblUser.Plant
         .ConfText = frmMultiBatchTrack.ConfirmationText
         .PostDate = Format(frmMultiBatchTrack.DtpPostDate, "YYYYMMDD")
         .EndDate = Format(frmMultiBatchTrack.DtpPostDate, "YYYYMMDD")
         .EndTime = Format(frmMultiBatchTrack.DtpPostTime, "HHMMSS")
      End With

      'delay Processing if the consecutive bookings are for the same Order
      'to avoid locking problems
      If mbOrder.Number = oldOrderNo And Cnt > 1 Then
         Start = Timer
         Do While Timer < Start + 4
         Loop
      End If

      'Call the function to make the bookings in SAP
      If MakeConfirmation(Conf, mbOrder, mbOp) = False Then
         GoTo ExitOnError
      Else
         If flxMultiBookings.Rows > 2 Then
            flxMultiBookings.RemoveItem (1)
         Else
            flxMultiBookings.Rows = 1
         End If

         'Adjust the time remaining
         tmpTotalTime = tmpTotalTime - Conf.ActWork
         TotalTime = Format(Str$(tmpTotalTime), "######.00")
         Me.Refresh

      End If

      'reset the previous order no
      oldOrderNo = mbOrder.Number

   Next Cnt

   flxMultiBookings.Row = 0
   flxMultiBookings.RowSel = 0

   MsgBox "ALL BOOKINGS SUCCESSFUL", vbOKOnly, cnDialogTitleWorkBook
   frmMultiBooking.Hide

   Exit Sub

ExitOnError:
   'Reset StartTime for Display Period
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   'Switch on the timer
   Timer1.Enabled = True
   Exit Sub

ErrorHandler:
   Msg = "Error: " & Err.Description
   MsgBox Msg, vbExclamation
   Resume Next

End Sub

Private Sub cmdRemovefromList_Click()

   Dim ListPosition As Integer

   'Switch off the timer
   Timer1.Enabled = False

   'Validate the selected items
   If flxMultiBookings.Rows = 1 Then
      MsgBox "NO CONFIRMATION NUMBERS HAVE BEEN ENTERED", vbExclamation, cnDialogTitleOpInfo
      GoTo ExitOnError
   ElseIf RowSelected = False Then
      MsgBox "PLEASE SELECT AN ITEM FROM THE LIST", vbExclamation, cnDialogTitleOpInfo
      GoTo ExitOnError
   ElseIf flxMultiBookings.Rows = 2 Then
      RowSelected = False
      flxMultiBookings.Rows = 1
      flxMultiBookings.Row = 0
      flxMultiBookings.RowSel = 0
      'MsgBox "CANNOT REMOVE LAST ITEM FROM LIST", vbExclamation, cnDialogTitleOpInfo
      GoTo ExitOnError
   ElseIf flxMultiBookings.RowSel - flxMultiBookings.Row <> 0 Then
      MsgBox "ONLY ONE ITEM CAN BE SELECTED", vbExclamation, cnDialogTitleOpInfo
      GoTo ExitOnError
   End If

   ListPosition = flxMultiBookings.Row
   flxMultiBookings.RemoveItem (ListPosition)

ExitOnError:
   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   'Switch on the timer
   Timer1.Enabled = True
   Exit Sub

End Sub

Private Sub ConfirmationNumber_KeyPress(KeyAscii As Integer)

   Dim DialogResponse As Integer

   If KeyAscii = 13 Then

      'set focus back in this control
      ConfirmationNumber.SetFocus

      'Switch off the timer
      Timer1.Enabled = False

      'test SAP connection
      If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
         GoTo ExitOnError
      End If

      'Set the Conf Number in the UDT and call the function
      'to read the Operation
      mbOp.ConfNo = ConfirmationNumber
      If ReadOrderOp(mbOrder, mbOp, mbOrdTypeCfg) = False Then
         GoTo ExitOnError
      End If

      'Check Op can be confirmed via Status checks on Order and Op
      If OpisOK2Confirm(mbOrdTypeCfg, mbOp) = False Then
         GoTo ExitOnError
      End If

      'Test for Op Plant not the same as Person Plant
      'and warn user
      ' 1/11/99 change to error - requested by Andy Dickinson
      If gblUser.Plant <> mbOp.Plant Then
         DialogResponse = MsgBox("PLANNED WORKCENTRE IS NOT IN YOUR PLANT", _
            vbExclamation, cnDialogTitleWorkBook)
         GoTo ExitOnError
      End If

      Row = flxMultiBookings.Rows
      flxMultiBookings.AddItem ("")
      flxMultiBookings.TextMatrix(Row, 0) = mbOp.ConfNo
      flxMultiBookings.TextMatrix(Row, 1) = Format(mbOrder.Number, "############")
      flxMultiBookings.TextMatrix(Row, 2) = mbOp.Number
      flxMultiBookings.TextMatrix(Row, 3) = mbOp.WorkCentre
      flxMultiBookings.TextMatrix(Row, 4) = mbOp.Desc
      flxMultiBookings.TextMatrix(Row, 5) = mbOp.ActivityType
      flxMultiBookings.TextMatrix(Row, 6) = mbOrder.Category
      flxMultiBookings.TextMatrix(Row, 7) = mbOrder.Aufpl
      flxMultiBookings.TextMatrix(Row, 8) = mbOp.Aplzl
      flxMultiBookings.TextMatrix(Row, 9) = mbOrder.Desc
      flxMultiBookings.TextMatrix(Row, 10) = mbOrder.Type
      flxMultiBookings.TextMatrix(Row, 11) = mbOrder.Plant

      GoTo ExitOnError

   End If

   Exit Sub

ExitOnError:

   'Reset confirmation number
   ConfirmationNumber = ""
   ConfirmationNumber.SetFocus

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

   'Switch on the timer
   Timer1.Enabled = True

   Exit Sub

End Sub

Private Sub DtpPostDate_DropDown()

   'Reset the Start Time
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   'Disable Timer
   Timer1.Enabled = False

End Sub

Private Sub DtpPostDate_LostFocus()

   'Enable the timer
   Timer1.Enabled = True

End Sub

Private Sub DtpPostTime_Change()

    'Reset StartTime
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub DtpPostTime_Click()

    'Reset StartTime
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub DtpPostTime_KeyPress(KeyAscii As Integer)

    'Reset StartTime
    Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub flxMultiBookings_SelChange()

   RowSelected = True
   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub Form_Activate()

   'Set the Start Time and Enable the Timer
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   Timer1.Enabled = True

   'Set Multi Booking form initial values
   With frmMultiBooking
      .PersonName = gblUser.PersName
      .CheckNumber = gblUser.ClockNumber
   End With

   gblSAP.TimeOnEntry = gblSAP.LocalDateTime

   With frmMultiBatchTrack
      .DtpPostDate = gblSAP.LocalDateTime
      .DtpPostTime = gblSAP.LocalDateTime
      .DtpPostDate.MaxDate = gblSAP.LocalDateTime
      .DtpPostDate.MinDate = gblSAP.LocalDateTime - gblPlantCfg.CancelPeriod
   End With


   'Initialise fields
   ConfirmationNumber = ""
   TotalTime = ""

End Sub

Private Sub Form_Deactivate()

   Timer1.Enabled = False

End Sub

Private Sub Form_Load()

   'Initialise the FlexGrid
   'MsgBox "Initialising FlexGrid"
   flxMultiBookings.Clear
   flxMultiBookings.Rows = 1
   flxMultiBookings.Row = 0
   flxMultiBookings.RowSel = 0
   RowSelected = False

   'Set the Colums and Headings for the FlexGrid
   'nb: heights and widths are in twips (1440 twips per inch)
   Row = 0
   flxMultiBookings.RowHeight(Row) = 500
   flxMultiBookings.Cols = 12
   flxMultiBookings.ColWidth(0) = 1400
   flxMultiBookings.ColWidth(1) = 1400
   flxMultiBookings.ColWidth(2) = 500
   flxMultiBookings.ColWidth(3) = 900
   flxMultiBookings.ColWidth(4) = 5200
   flxMultiBookings.ColWidth(5) = 600
   flxMultiBookings.ColWidth(6) = 0  'Hide this Column
   flxMultiBookings.ColWidth(7) = 0  'Hide this Column
   flxMultiBookings.ColWidth(8) = 0  'Hide this Column
   flxMultiBookings.ColWidth(9) = 0  'Hide this Column
   flxMultiBookings.ColWidth(10) = 0  'Hide this Column
   flxMultiBookings.ColWidth(11) = 0  'Hide this Column

   flxMultiBookings.TextMatrix(Row, 0) = "Confirmation No"
   flxMultiBookings.TextMatrix(Row, 1) = "Order No"
   flxMultiBookings.TextMatrix(Row, 2) = "Op No"
   flxMultiBookings.TextMatrix(Row, 3) = "Work Centre"
   flxMultiBookings.TextMatrix(Row, 4) = "Op Description"
   flxMultiBookings.TextMatrix(Row, 5) = "Act Type"
   flxMultiBookings.TextMatrix(Row, 6) = "Order Category"
   flxMultiBookings.TextMatrix(Row, 7) = "Aufpl"
   flxMultiBookings.TextMatrix(Row, 8) = "Aplzl"
   flxMultiBookings.TextMatrix(Row, 9) = "Order Description"
   flxMultiBookings.TextMatrix(Row, 10) = "Order Type"
   flxMultiBookings.TextMatrix(Row, 11) = "Plant"

End Sub

Private Sub Timer1_Timer()

   If Not OK2DisplayForm(False, gblPlantCfg.FormTimeOut) Then
      Timer1.Enabled = False
      TotalTime.SetFocus
      Me.Hide
   End If

End Sub

Private Sub TotalTime_KeyPress(KeyAscii As Integer)

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub TotalTime_LostFocus()

   If Val(TotalTime) = 0 And frmMultiBooking.Visible Then
      Timer1.Enabled = False
      MsgBox "PLEASE ENTER TIME SPENT", vbExclamation, cnDialogTitleWorkBook
      Timer1.Enabled = True
      TotalTime.SetFocus
   End If

End Sub

Private Sub WorkUnits_KeyPress(KeyAscii As Integer)

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

