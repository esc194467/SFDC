VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmWorkBook 
   Caption         =   "SFDC - Work Booking"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWorkBook.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DtpPostTime 
      Height          =   375
      Left            =   1320
      TabIndex        =   22
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   51118082
      CurrentDate     =   36917
   End
   Begin MSComCtl2.DTPicker DtpPostDate 
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   180
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "ddd dd MMM yyyy"
      Format          =   51118080
      CurrentDate     =   36917
   End
   Begin VB.TextBox Plant 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   3
      Top             =   2400
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   360
      Top             =   4320
   End
   Begin VB.TextBox PersonName 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   17
      Top             =   1560
      Width           =   4575
   End
   Begin VB.TextBox CheckNumber 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton MakeBooking 
      Caption         =   "Make Booking"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox WorkUnits 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Text4"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox ActivityType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   4
      Text            =   "T&M"
      Top             =   2400
      Width           =   975
   End
   Begin MSMask.MaskEdBox ActualWork 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "000000.00"
      PromptChar      =   "_"
   End
   Begin VB.TextBox OrderNumber 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox OpNumber 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox WorkCentre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   5040
      Picture         =   "frmWorkBook.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2340
   End
   Begin VB.Label Label12 
      Caption         =   "Posting Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Posting Date"
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
      Left            =   240
      TabIndex        =   19
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Plant"
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
      Left            =   4920
      TabIndex        =   18
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Check No"
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
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label5 
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
      Left            =   2040
      TabIndex        =   14
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Units"
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
      Left            =   4200
      TabIndex        =   13
      Top             =   3240
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Activity Type"
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
      Left            =   5760
      TabIndex        =   11
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Actual Work"
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
      Left            =   2880
      TabIndex        =   10
      Top             =   3240
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Order No"
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
      Left            =   720
      TabIndex        =   9
      Top             =   2160
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Op No"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   2160
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Work Centre"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   2160
      Width           =   900
   End
End
Attribute VB_Name = "frmWorkBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Dummy As Boolean

Private Sub ActivityType_KeyPress(KeyAscii As Integer)

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub ActualWork_KeyPress(KeyAscii As Integer)

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub ActualWork_LostFocus()

   If ActualWork = "" And frmWorkBook.Visible Then
      'disable the timer until user responds
      Timer1.Enabled = False
      MsgBox "PLEASE ENTER TIME SPENT", vbExclamation, cnDialogTitleWorkBook
      'Switch on the timer
      Timer1.Enabled = True
      ActualWork.SetFocus
   ElseIf Val(ActualWork) = 0 And frmWorkBook.Visible Then
      'disable the timer until user responds
      Timer1.Enabled = False
      MsgBox "FOR ZERO BOOKINGS" & Chr(10) & _
         "PLEASE USE ZERO TIME BOOKING BUTTON" & Chr(10) & _
         "ON THE FRONT SCREEN", vbExclamation, cnDialogTitleWorkBook
      'Switch on the timer
      Timer1.Enabled = True
      ActualWork.SetFocus
   ElseIf Val(ActualWork) > 9999 And frmWorkBook.Visible Then
      'disable the timer until user responds
      Timer1.Enabled = False
      MsgBox "MAXIMUM TIME IS 9999", vbExclamation, cnDialogTitleWorkBook
      'Switch on the timer
      Timer1.Enabled = True
      ActualWork.SetFocus
   End If

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


Private Sub Form_Activate()

   'Set the Start Time and Enable the Timer
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   Timer1.Enabled = True

   'Set Work Booking Form with initial values
   With frmWorkBook
      .CheckNumber = gblUser.ClockNumber
      .PersonName = gblUser.PersName
      .OrderNumber = Format(gblOrder.Number, "############")
      .OpNumber = gblOperation.Number
      .WorkCentre = gblOperation.WorkCentre
      .Plant = gblOperation.Plant
      .ActualWork = ""
      .WorkUnits = gblOperation.WorkUnits
      .DtpPostDate = gblSAP.LocalDateTime
      .DtpPostDate.MaxDate = gblSAP.LocalDateTime
      .DtpPostDate.MinDate = gblSAP.LocalDateTime - gblPlantCfg.CancelPeriod
      .DtpPostTime = gblSAP.LocalDateTime
      .ActivityType = gblOperation.ActivityType
   End With

   gblSAP.TimeOnEntry = gblSAP.LocalDateTime

End Sub

Private Sub Form_Deactivate()

   Timer1.Enabled = False

End Sub

Private Sub MakeBooking_Click()

   Dim Conf As ConfInfo
   Dim DialogResponse As Integer
   Dim tmpPostDateTime As String * 19
   Dim PostDateTime As Date

   'Switch off the timer
   Timer1.Enabled = False

   'test SAP connection
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
      GoTo ExitOnError
   End If

   'Check that a time has been entered
   If ActualWork = "" Then
      ActualWork_LostFocus
      GoTo ExitOnError
   End If

   'Test for Op Plant not the same as Person Plant
   'and warn user
   ' 1/11/99 change to error - requested by Andy Dickinson
   If gblUser.Plant <> gblOperation.Plant Then
      DialogResponse = MsgBox("PLANNED WORKCENTRE IS NOT IN YOUR PLANT", _
         vbExclamation, cnDialogTitleWorkBook)
      GoTo ExitOnError
   End If

   'Check for Post dated Booking
    PostDateTime = DateValue(DtpPostDate) + TimeValue(DtpPostTime)
'* -- TCR7117 -- *
        gdate = PostDateTime
'* -- TCR7117 -- *
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

  If gblOrder.Category = "10" And gblPlantCfg.BookingMode = cnPostEventBooking Then  ' TCR7117 Start ...
  'Call BookOff Tracking form
   frmBookOff.Show vbModal
   frmBookOff.Timer1.Enabled = False
   Me.Refresh
  Else ' TCR7117 Start ...
   'Call Batch Tracking form
   frmBatchTrack.Show vbModal
   frmBatchTrack.Timer1.Enabled = False
   Me.Refresh
   End If ' TCR7117 Start ...

   'Test BatchTrack form was exited correctly
   If gblBatchTrackOK = False Then
      GoTo ExitOnError
   End If

   If frmBatchTrack.ChkPostDated = cnChecked Then
      'Post Dated booking
      'Check that Post Date/Time combination is in the past
      If PostDateTime >= gblSAP.TimeOnEntry Then
         MsgBox ("THIS IS A POST DATED BOOKING - PLEASE CORRECT DATE/TIME")
         DtpPostDate.SetFocus
         GoTo ExitOnError
      End If

      Conf.PostDate = Format(DtpPostDate, "YYYYMMDD")
      Conf.EndDate = Format(DtpPostDate, "YYYYMMDD")
      Conf.EndTime = Format(DtpPostTime, "HHMMSS")
   Else
      'NON post dated booking so...
      'Set date/times as current SAP date/time
      Conf.PostDate = Format(gblSAP.LocalDateTime, "YYYYMMDD")
      Conf.EndDate = Format(gblSAP.LocalDateTime, "YYYYMMDD")
      Conf.EndTime = Format(gblSAP.LocalDateTime, "HHMMSS")
      Conf.WorkUnits = gblOperation.WorkUnits
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

   'Set remaining Confirmation values from the the form fields
   With Conf
      .ActType = frmWorkBook.ActivityType
      .ActWork = frmWorkBook.ActualWork
      .ConfText = frmBatchTrack.ConfirmationText
      .DevReason = frmBatchTrack.cmboReasons
      .WorkCentre = frmWorkBook.WorkCentre
      .Plant = frmWorkBook.Plant
      'Update yield and scrap values in confirmation for Work Booking processing
    If gblFunction = cnWorkBook Then '--TCR7117 --
      .Yield = frmBatchTrack.txtYieldQty
      .Scrap = frmBatchTrack.txtScrapQty
    End If '--TCR7117 --
   End With
   
   'Call the function to make the Confirmation in SAP
   If MakeConfirmation(Conf, gblOrder, gblOperation) = False Then
      GoTo ExitOnError
   Else
      'Inform user of successful booking and check for any more
      DialogResponse = MsgBox("BOOKING SUCCESSFUL - DO YOU WANT TO MAKE ANOTHER BOOKING?", _
         (vbYesNo + vbDefaultButton2), cnDialogTitleWorkBook)

      If DialogResponse = vbYes Then
         gblContinueBooking = True
         gblLastBookingTime = Now
         Timer1.Enabled = False
         frmWorkBook.Hide
      Else
         gblContinueBooking = False
         Timer1.Enabled = False
         frmWorkBook.Hide
      End If

      ' Set logon screen fields to null ready for new booking
      ResetLogonForm
      frmLogon.ConfirmationNumber.SetFocus

   End If

   Exit Sub

ExitOnError:
   'Reset StartTime for Display Period
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)
   'Switch on the timer
   Timer1.Enabled = True
   Exit Sub

End Sub

Private Sub Plant_LostFocus()

   Plant = UCase(Plant)

End Sub

Private Sub Timer1_Timer()

   If Not OK2DisplayForm(False, gblPlantCfg.FormTimeOut) Then
      Timer1.Enabled = False
      Me.Hide
   End If

End Sub

Private Sub WorkCentre_KeyPress(KeyAscii As Integer)

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

End Sub

Private Sub WorkCentre_LostFocus()
   WorkCentre = UCase(WorkCentre)
End Sub

