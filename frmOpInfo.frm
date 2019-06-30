VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOpInfo 
   Caption         =   "SFDC - Op Information"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   Icon            =   "frmOpInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtConfNo 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtSalesOrderNumber 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame FraComponents 
      Caption         =   "Components"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   10215
      Begin MSFlexGridLib.MSFlexGrid GrdComponentList 
         Height          =   1215
         Left            =   240
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "List of components for the activity / operation"
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   -2147483637
         GridColor       =   -2147483637
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         FormatString    =   $"frmOpInfo.frx":0442
      End
   End
   Begin VB.TextBox DDEbox 
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdViewDocument 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Open Document"
      Height          =   1095
      Left            =   8400
      Picture         =   "frmOpInfo.frx":04DF
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Click to open the hi-lighted document"
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox WorkCentre 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox OpNumber 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox OrderNumber 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame FraTools 
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   10215
      Begin MSFlexGridLib.MSFlexGrid GrdToolList 
         Height          =   1215
         Left            =   240
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "List of tools for the activity / operation"
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         BackColorBkg    =   -2147483637
         GridColor       =   -2147483637
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         FormatString    =   $"frmOpInfo.frx":0921
      End
   End
   Begin VB.Frame FraOpDesc 
      Caption         =   "Op Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   10215
      Begin RichTextLib.RichTextBox OpLongText1 
         Height          =   1575
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Activity / Operation Description"
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   2778
         _Version        =   393217
         BackColor       =   -2147483638
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         RightMargin     =   1
         TextRTF         =   $"frmOpInfo.frx":09CA
      End
      Begin VB.PictureBox OpLongText 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   240
         ScaleHeight     =   1575
         ScaleWidth      =   9855
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Activity / Operation description"
         Top             =   240
         Width           =   9855
      End
   End
   Begin VB.Frame FraDocuments 
      Caption         =   "Documents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   10215
      Begin MSFlexGridLib.MSFlexGrid GrdDocumentList 
         Height          =   1215
         Left            =   240
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "List of documents for the activity / operation"
         Top             =   360
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
         BackColor       =   16777215
         BackColorBkg    =   -2147483637
         GridColor       =   -2147483637
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   "          "
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Confirmation No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4920
      TabIndex        =   12
      Top             =   120
      Width           =   1425
   End
   Begin VB.Label Label4 
      Caption         =   "Sales Order No."
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
      Left            =   8520
      TabIndex        =   13
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Work Centre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Op No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Order No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmOpInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Dummy As Boolean

Private Sub cmdViewDocument_Click()

   Dim ListPosition As Integer
   Dim PackIndex As Integer
   Dim i As Integer
   Dim w As String
   Dim x As Integer
   Dim Y As Variant
   Dim myRetVal As Long
   Dim ShellStr As String
   Dim FileOpen As String
   Dim StartTime As Long
   Dim Tries As Integer
   Dim lpAppName As String
   Dim lpKeyName As String
   Dim lpString As String
   Dim lpFilename As String
   Dim ErrMsg As String

   On Error GoTo error_handler

   'Validate the selected items
   If GrdDocumentList.Text = "" Then
      MsgBox "No Documents have been associated with this Op", vbExclamation, cnDialogTitleOpInfo
      GoTo Exit_Sub
   End If

   'Reset StartTime
   Dummy = OK2DisplayForm(True, gblPlantCfg.FormTimeOut)

   'Check whether SAP is connected or STOP flag is set
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, False) Then
      Exit Sub
   End If

   ListPosition = GrdDocumentList.Row

   'Reading document info from selected row
   With gblDocument
      .Name = GrdDocumentList.TextMatrix(ListPosition, 0)
      .ApplnType = GrdDocumentList.TextMatrix(ListPosition, 2)
      .DocType = GrdDocumentList.TextMatrix(ListPosition, 3)
      .AppPath = GrdDocumentList.TextMatrix(ListPosition, 4)
      .FullFilePath = GrdDocumentList.TextMatrix(ListPosition, 5)
   End With

   Select Case gblDocument.DocType

      Case "REC"

         Select Case gblDocument.ApplnType

            Case "XLS" 'New code

               On Error Resume Next

               Dim Path As String * 250
               Dim nAnswer As Integer
               
               'Reset the executed attribute of all the Methods
               'to ensure that the method is re-executed with the new
               'parameters
               Call ResetMethods

               '------------------------------------------------------------------------------

               Set gblExcel = Nothing

               'Check if SalesOrderNumber is blank
               If gblOrder.SalesOrder = "" Then
                  MsgBox "Recording Document Template cannot be opened" _
                     & Chr(13) & Chr(13) & "Sales Order Number does not have a value" _
                     & Chr(13) & Chr(13) & "Please contact local support" _
                     , vbExclamation, cnDialogTitleOpInfo
                  GoTo Exit_Sub
               End If

               '=====================
               ' Check if file exists
               '=====================
               If Dir(gblDocument.FullFilePath) = "" Then
                  MsgBox "Recording Document Template cannot be opened" _
                     & Chr(13) & Chr(13) & "The linked document cannot be found/accessed" _
                     & Chr(13) & Chr(13) & "Please contact local support" _
                     , vbExclamation, cnDialogTitleOpInfo
                  GoTo Exit_Sub
               End If

               Load frmRecControl
               frmRecControl.Show

               '============================
               ' Make the form always on top
               '============================
               Call SetWindowPos(frmRecControl.hwnd, cnHWND_TOPMOST, 0, 0, 0, 0, cnSWP_NOMOVE Or cnSWP_NOSIZE)
               '=====================================================================================
               ' Set open document flag to force the timer on the form to open a document (szDocName)
               '=====================================================================================
               gblbOpenDoc = True
               frmOpInfo.Hide
               frmLogon.Hide

            Case "RDM"  'Recording Document Manager file

               'Check if SalesOrderNumber is blank
               If gblOrder.SalesOrder = "" Then
                  MsgBox "Recording Document cannot be opened" _
                     & Chr(13) & Chr(13) & "Sales Order Number does not have a value" _
                     & Chr(13) & Chr(13) & "Please contact local support" _
                     , vbExclamation, cnDialogTitleOpInfo
                  GoTo Exit_Sub
               End If

               'Check that temp dir exists
               w = Dir(cnTemporaryFileAddress, vbDirectory)
               If w = "" Then
                  MkDir cnTemporaryFileAddress
               End If

               'Attempt to write Sales Order Number to file and store at this location
               lpAppName = "General"
               lpKeyName = "SalesOrderNo"
               lpString = Trim(gblOrder.SalesOrder)
               lpFilename = cnTemporaryFileAddress + cnTemporaryFileName

               x = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFilename)

               If x = 0 Then 'Write operation failed
                  MsgBox "Unable to write to " & "'" & lpFilename & "'" _
                     & Chr(13) & Chr(13) & "Open request on the document cancelled" _
                     & Chr(13) & Chr(13) & "Please contact local support", _
                     vbExclamation, cnDialogTitleOpInfo
                  GoTo Exit_Sub
               End If

               'Execute application and file
               Y = ShellExecute(0, "open", gblDocument.FullFilePath, 0, "", 5)

               If Y > 32 Then
                  ' open operation successful
               Else
                  ' open operation unsuccesful
                  ErrMsg = ShellExErrMsg(Y)

                  MsgBox "Unable to open the file - " & "'" _
                     & gblDocument.FullFilePath & "'" & Chr(13) & Chr(13) _
                     & ErrMsg & Chr(13) & Chr(13) _
                     & "Please contact local support" _
                     , vbExclamation, cnDialogTitleOpInfo

                  'Delete temporary file
                  Kill (cnTemporaryFileAddress + cnTemporaryFileName)

                  GoTo Exit_Sub
               End If

            Case Else

               'Execute application and file
               If gblDocument.AppPath = "%auto%" Then

                  'Execute application and file
                  Y = ShellExecute(0, "open", gblDocument.FullFilePath, 0, "", 5)

                  If Y > 32 Then
                     ' open operation successful
                  Else
                     ' open operation unsuccesful
                     ErrMsg = ShellExErrMsg(Y)

                     MsgBox "Unable to open the file - " & "'" _
                        & gblDocument.FullFilePath & "'" & Chr(13) & Chr(13) _
                        & ErrMsg & Chr(13) & Chr(13) _
                        & "Please contact local support" _
                        , vbExclamation, cnDialogTitleOpInfo

                     GoTo Exit_Sub
                  End If

               Else
                  Y = Shell(gblDocument.AppPath & Space(1) & gblDocument.FullFilePath, vbNormalFocus)

                  If Y = 0 Then 'Execution failed
                     MsgBox "Unable to open the application associated with file - " & "'" _
                        & gblDocument.FullFilePath & "'" & Chr(13) & Chr(13) & "Please contact local support" _
                        , vbExclamation, cnDialogTitleOpInfo

                     GoTo Exit_Sub

                  End If

               End If
         End Select

      Case Else
         'Execute application and file
         If gblDocument.AppPath = "%auto%" Then

            'Execute application and file
            Y = ShellExecute(0, "open", gblDocument.FullFilePath, 0, "", 5)

            If Y > 32 Then
               ' open operation successful
            Else
               ' open operation unsuccesful
               ErrMsg = ShellExErrMsg(Y)

               MsgBox "Unable to open the file - " & "'" _
                  & gblDocument.FullFilePath & "'" & Chr(13) & Chr(13) _
                  & ErrMsg & Chr(13) & Chr(13) _
                  & "Please contact local support" _
                  , vbExclamation, cnDialogTitleOpInfo

               GoTo Exit_Sub
            End If

         Else
            Y = Shell(gblDocument.AppPath & Space(1) & gblDocument.FullFilePath, vbNormalFocus)

            If Y = 0 Then 'Execution failed
               MsgBox "Unable to open the application associated with file - " & "'" _
                  & gblDocument.FullFilePath & "'" & Chr(13) & Chr(13) & "Please contact local support" _
                  , vbExclamation, cnDialogTitleOpInfo

               GoTo Exit_Sub

            End If

         End If

   End Select

Exit_Sub:

   GrdDocumentList.SetFocus
   Exit Sub

error_handler:

   'debug.print "Error number is " & Err.Number

   Select Case Err.Number

      Case -2147467259 ' Can't find URL
         Err.Clear
         'frmWebBrowser.Hide
         frmOpInfo.Refresh
         GoTo Exit_Sub

      Case 31004 'Invalid Application or File Ref
         MsgBox "Invalid FILE reference in SAP" _
            & Chr(13) & "Please contact Local Support" _
            , vbExclamation, cnDialogTitleOpInfo

      Case 282 'Cannot Initiate DDE Conversation
         'For up to 3 tries - Wait 2 secs and try again
         Tries = Tries + 1
         If Tries = 3 Then Exit Sub

         StartTime = Timer()
         Do While Timer() < StartTime + 2
         Loop
         Resume

      Case Else
         MsgBox (Error$)
         GoTo Exit_Sub

   End Select

   Resume Next

End Sub

Private Sub DocumentList_DblClick()

   cmdViewDocument_Click

End Sub

Private Sub Form_Load()

   Me.Caption = App.Title & " - " & cnDialogTitleOpInfo

   'Set hidden columns in frmOpinfo Document List to hold parameters passed
   'when calling the document applications
   With frmOpInfo.GrdDocumentList
      .ColWidth(2) = 0 'ApplnType
      .ColWidth(3) = 0 'DocType
      .ColWidth(4) = 0 'AppPath
      .ColWidth(5) = 0 'FullFilePath
   End With

End Sub

Private Sub OpLongText1_OLEDragDrop(data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)

   data.GetData vbCFRTF

End Sub

Private Sub ResetMethods()

   Dim i As Integer
   Dim MaxMethIndx As Integer

   MaxMethIndx = UBound(Method, 1)
   
   For i = 1 To MaxMethIndx
      Method(i).Executed = False
   Next i

End Sub


