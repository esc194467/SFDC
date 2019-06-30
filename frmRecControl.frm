VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRecControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SFDC - Recordings Controller"
   ClientHeight    =   465
   ClientLeft      =   315
   ClientTop       =   780
   ClientWidth     =   8805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecControl.frx":0000
   LinkTopic       =   "System"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   8805
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer ShutdownTimer 
      Interval        =   500
      Left            =   120
      Top             =   480
   End
   Begin VB.Label labCurrentDoc 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCurrentDocText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Document =    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1950
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu FileExit 
         Caption         =   "&Return to OP Info screen"
      End
      Begin VB.Menu FileNull2 
         Caption         =   "-"
      End
      Begin VB.Menu FilePrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
      End
      Begin VB.Menu FileNull1 
         Caption         =   "-"
      End
      Begin VB.Menu SaveFile 
         Caption         =   "&Copy and Save As......."
      End
   End
   Begin VB.Menu RecordingsMenu 
      Caption         =   "&Recordings"
      Begin VB.Menu UpdateValues 
         Caption         =   "&Update Values"
      End
      Begin VB.Menu UpdateNull1 
         Caption         =   "-"
      End
      Begin VB.Menu EnableUpdate 
         Caption         =   "&Values Confirmed"
      End
   End
   Begin VB.Menu OpInfo 
      Caption         =   "&Op Details"
      WindowList      =   -1  'True
      Begin VB.Menu OrderNo 
         Caption         =   "&OrderNo"
      End
      Begin VB.Menu OpNo 
         Caption         =   "OpNo"
      End
      Begin VB.Menu WorkCentre 
         Caption         =   "WorkCentre"
      End
      Begin VB.Menu ConfNo 
         Caption         =   "ConfNo"
      End
      Begin VB.Menu SalesNo 
         Caption         =   "SalesNo"
      End
   End
   Begin VB.Menu ViewMenu 
      Caption         =   "&View"
      Visible         =   0   'False
      Begin VB.Menu ViewCurrentDocument 
         Caption         =   "&Current Document"
         Visible         =   0   'False
      End
      Begin VB.Menu ViewNull1 
         Caption         =   "-"
      End
      Begin VB.Menu ViewDocument 
         Caption         =   "Name"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmRecControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private locCurrentDoc                       As Integer

Private bInTimer                            As Boolean
Private bShutdown                           As Boolean
Private bInClose                            As Boolean

Private szDocName                           As String

Public WithEvents gblAppExcel               As Excel.Application
Attribute gblAppExcel.VB_VarHelpID = -1

Private Sub EnableUpdate_Click()

   On Error Resume Next

   frmRecControl.UpdateValues.Enabled = True
   frmRecControl.EnableUpdate.Enabled = False

End Sub

Private Sub FileClose_Click()

   On Error Resume Next

   '---------------------------------------------------------------------

   If Me.CurrentDocument <> -1 Then

      If Not bInClose Then
         bInClose = True

         Call CloseRecordingDocument(Me.CurrentDocument, Me)

         DoEvents
         DoEvents

         If CountDocumentsOpen() <> 0 Then
            Me.CurrentDocument = Me.CurrentDocument
         Else
            Unload Me
         End If

         bInClose = False
      End If
   End If

End Sub

Private Sub FileExit_Click()

   On Error Resume Next

   If UpdateValues.Enabled = True Then
      gblExit_File = True
      Call UpdateValues_Click
   Else
      If EnableUpdate.Enabled = True Then
         gblExit_File = True
         Call UpdateValues_Click
      End If
   End If
   Unload Me

End Sub

Private Sub FilePrint_Click()

   On Error Resume Next

   '---------------------------------------------------------------------

   If Me.CurrentDocument <> -1 Then
      Call PrintRecordingDocument(Me.CurrentDocument)
   End If

End Sub

Private Sub Form_Load()

   On Error Resume Next

   Dim i As Integer

   '------------------------------------------------------------------
   Me.Caption = App.Title & " - " & cnDialogTitleOpInfo & " - " & cnDialogTitleRecCntl

   OrderNo.Caption = "Order No.: " & frmLogon.OrderNumber
   OpNo.Caption = "Operation No.: " & frmLogon.OpNumber
   WorkCentre.Caption = "Work Centre: " & frmLogon.WorkCentre
   ConfNo.Caption = "Confirmation No.: " & frmLogon.ConfirmationNumber
   SalesNo.Caption = "Sales Order No.: " & gblOrder.SalesOrder

   szDocName = ""
   gblbOpenDoc = False

   bInClose = False
   bShutdown = False
   bInTimer = False

   ReDraw

   For i = 0 To 20
      If i <> 0 Then
         Load Me.ViewDocument(i)
      End If
      Me.ViewDocument(i).Visible = False
   Next i

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   On Error Resume Next

   Dim i         As Integer

   '------------------------------------------------------------------

   For i = 0 To gblDocCount - 1
      If Not CloseRecordingDocument(i, Me) Then
         Cancel = True
         Exit For
      End If
   Next i

End Sub

Private Sub Form_Resize()

   On Error Resume Next

   ReDraw

End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Error Resume Next

   Dim i                               As Integer

   '------------------------------------------------------------------

   Me.Hide

   For i = 20 To 1 Step -1
      Unload Me.ViewDocument(i)
   Next i

   Call ShutdownExcel(Me)

   frmLogon.Show
   frmOpInfo.Show
   frmOpInfo.SetFocus

End Sub

Public Sub SetDocument(Index As Integer, lpName As String, bVisible As Boolean)

   On Error Resume Next

   Me.ViewDocument(Index).Caption = lpName
   Me.ViewDocument(Index).Visible = bVisible

End Sub

Public Property Get CurrentDocument() As Integer

   CurrentDocument = locCurrentDoc

End Property

Public Property Let CurrentDocument(ByVal nValue As Integer)

   On Error Resume Next

   Dim szName                      As String

   '------------------------------------------------------------------

   locCurrentDoc = nValue

   If locCurrentDoc >= 0 Then
      Me.labCurrentDoc.Caption = ShortenFilenameLevels(gblDocuments(locCurrentDoc).Name, 1)
      szName = RemoveExtension(ShortenFilenameLevels(Me.labCurrentDoc.Caption, 1))

      'Me.FileClose.Caption = "&Close " + szName
      'Me.FileClose.Enabled = True
      Me.FilePrint.Caption = "&Print " + szName
      Me.FilePrint.Enabled = True

      If UCase(Trim(gblPlantCfg.CanUpdRecDocs)) = cnSAPTrue Then
         Me.UpdateValues.Caption = "&Update Values For " + szName
         Me.UpdateValues.Enabled = True
      Else
         Me.UpdateValues.Caption = "&Update Values For " + szName
         Me.UpdateValues.Enabled = False
      End If
      gblDocuments(locCurrentDoc).Doc.Activate
      gblExcel.Visible = True
   Else
      Me.labCurrentDoc.Caption = ""
      'Me.FileClose.Caption = "&Close"
      'Me.FileClose.Enabled = False
      Me.FilePrint.Caption = "&Print"
      Me.FilePrint.Enabled = False
      Me.UpdateValues.Caption = "&Update Values"
      Me.UpdateValues.Enabled = False

   End If

End Property

Public Function AddDocumentToMenu(nDoc As Integer) As Boolean

   '----------------------------------------------------------------

   Me.ViewDocument(nDoc).Caption = ShortenFilenameLevels(gblDocuments(nDoc).Name, 1)
   Me.ViewDocument(nDoc).Visible = True

End Function

Private Sub ReDraw()

   On Error Resume Next

   '---------------------------------------------------------------------------

   Me.labCurrentDoc.Width = Me.ScaleWidth - Me.labCurrentDoc.Left

End Sub

Private Sub gblAppExcel_WorkbookActivate(ByVal Wb As Excel.Workbook)

   On Error Resume Next

   Dim i               As Integer
   Dim nDoc            As Integer

   '--------------------------------------------------------------------

   nDoc = -1
   For i = 0 To gblDocCount - 1
      If gblDocuments(i).Name = Wb.FullName Then
         nDoc = i
         Exit For
      End If
   Next i

   Me.CurrentDocument = nDoc

End Sub

Private Sub gblAppExcel_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, Cancel As Boolean)

   On Error Resume Next

   Dim i               As Integer
   Dim nDoc            As Integer

   '--------------------------------------------------------------------

   Cancel = False
   If Not bInClose Then
      bInClose = True

      nDoc = -1
      For i = 0 To gblDocCount - 1
         If gblDocuments(i).Name = Wb.FullName Then
            nDoc = i
            Exit For
         End If
      Next i

      If nDoc <> -1 Then
         Call CloseRecordingDocument(nDoc, Me)
         Cancel = False
      End If

      bInClose = False
   End If

End Sub

Private Sub gblAppExcel_WorkbookDeactivate(ByVal Wb As Excel.Workbook)

   On Error Resume Next

   Dim i                               As Integer

   '---------------------------------------------------------------------

   If CountDocumentsOpen() = 0 Then
      bShutdown = True
   Else
      Me.CurrentDocument = -1
   End If

End Sub

Private Sub labCurrentDoc_Click()

   On Error Resume Next

   If gblDocuments(0).Enabled Then
      gblDocuments(0).Doc.Activate
      gblExcel.Visible = True

   End If
End Sub

Private Sub lblCurrentDocText_Click()

   On Error Resume Next

   If gblDocuments(0).Enabled Then
      gblDocuments(0).Doc.Activate
      gblExcel.Visible = True

   End If
End Sub

Private Sub picDisableTopMost_DblClick()

   'frmRecControl.Moveable = True
   'frmRecControl.ControlBox = True
   
   
End Sub

Private Sub SaveFile_Click()

   Dim nDoc As Integer
   Dim FileName As String
   Dim SalesOrder As String
   Dim ThisDir As String

   On Error GoTo ErrHandler

   'Then display the checkno/password dialog
   frmCheckPass.Show vbModal

   If gblPersonValidated Then
      'check person is authorised to save files
      If gblUser.CanSave = cnSAPTrue Then
         'save the Excel document to the appropriate folder
         nDoc = Me.CurrentDocument
         SalesOrder = Format(gblOrder.SalesOrder, "##########")
         With CommonDialog1
            .DefaultExt = ".xls"
            .FileName = SalesOrder & "_" & gblDocument.Name
            .InitDir = gblPlantCfg.SOPath & SalesOrder & "\"
            .Flags = cdlOFNOverwritePrompt
            .Filter = "Excel WorkBooks (*.xls)|*.xls"
            .FilterIndex = 2
            .DialogTitle = "Save Recording Document"
            .CancelError = True
         End With
         If nDoc <> -1 Then
            ChDir CommonDialog1.InitDir
            CommonDialog1.ShowSave
            FileName = CommonDialog1.FileName
            If FileName > vbNullString Then

               gblDocuments(nDoc).Doc.SaveCopyAs (FileName)
            End If
         End If

      Else
         'error message
         MsgBox "USER NOT AUTHORISED TO SAVE FILES", vbExclamation
      End If
      
   End If

ErrHandler:
   Select Case Err.Number
   
      Case 76
         ' Error triggered from CHDIR command, so...
         ' User cannot access the SO Data directory
         MsgBox "NETWORK USERID CANNOT ACCESS SALES ORDER SPECIFIC DATA DIRECTORY", vbExclamation
         Resume Next
         
      Case 32755
         'User pressed the Cancel button in the Save Dialogue so....
         Exit Sub

   End Select

End Sub

Private Sub ShutdownTimer_Timer()

   On Error Resume Next

   Dim nDoc                            As Integer

   '---------------------------------------------------------------------
   Screen.MousePointer = vbDefault
   '====================================================
   ' Retrieve document path and name from docinfo record
   '====================================================
   szDocName = gblDocument.FullFilePath
   '====================================================
   'gblbOpenDoc = True
   '================================================
   ' Check if the timers interval has been activated
   '================================================
   If Not bInTimer Then

      bInTimer = True

      '=====================================================
      ' Check if 1st time through and document needs opening
      '=====================================================
      If gblbOpenDoc Then
         'Screen.MousePointer = vbHourglass

         nDoc = OpenRecordingDocument(szDocName)
         If nDoc <> -1 Then
            Call AddDocumentToMenu(nDoc)
         Else
            If CountDocumentsOpen() = 0 Then
               Unload Me
            End If
         End If

         szDocName = ""
         gblbOpenDoc = False
         'Screen.MousePointer = vbDefault

      End If

      '==================================
      ' Check if excel has been shut down
      '==================================
      If bShutdown Then
         Unload Me
      End If

      bInTimer = False
   End If

End Sub

Private Sub UpdateValues_Click()

   On Error Resume Next

   Dim NameCount As Integer
   Dim OrderRowCount As Integer
   Dim j As Integer
   Dim n As Integer
   Dim l As Integer
   Dim bContinue As Boolean
   Dim AlreadyExists As Boolean
   Dim Response As Integer

   '===================================
   ' Determine No. of names in document
   '===================================
   NameCount = gblDocuments(Me.CurrentDocument).Doc.Names.Count
   ReDim gblStored(NameCount)
   '=========================================================================
   ' Obtain all names in document that begin RD and send to SAP for retrieval
   '=========================================================================
   j = 0
   SAPRetrieveRecValues.Tables.Item("CELL_INFO").FreeTable
   For n = 1 To NameCount
      If Mid(UCase(gblDocuments(Me.CurrentDocument).Doc.Names(n).Name), 1, 2) = "RD" Then
         j = j + 1
         gblStored(j).CellID = UCase(gblDocuments(Me.CurrentDocument).Doc.Names(n).Name)
         gblStored(j).RecValue = gblDocuments(Me.CurrentDocument).Doc.Names(n).RefersToRange.Value

         With SAPRetrieveRecValues.Tables.Item("CELL_INFO")
            .AppendRow
            .Value(j, "VBELN") = gblOrder.SalesOrder
            .Value(j, "AUFNR") = gblOrder.Number
            .Value(j, "RECID") = UCase(gblDocuments(Me.CurrentDocument).Doc.Names(n).Name)
         End With

      End If
   Next

   ReDim Preserve gblStored(j)
   '======================================
   ' Check if any user entered names exist
   '======================================
   If j = 0 Then
      'No values will be retrieved
      Exit Sub
   End If

   '=======================================
   ' Remove existing data from import table
   '=======================================
   RetrievedRecOrderValuesTable.FreeTable

   '====================
   ' Test SAP Connection
   '====================
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, True) Then
      Exit Sub
   End If

   '==============================================
   ' Send "CELL_INFO" table to SAP function module
   '================================================================
   ' Obtain "REC_ORDER_RETRIEVE" table back from SAP function module
   '================================================================
   If SAPRetrieveRecValues.Call = False Then
      MsgBox (SAPRetrieveRecValues.Exception)
      Exit Sub
   End If

   OrderRowCount = RetrievedRecOrderValuesTable.RowCount

   '=================================================================
   ' For each name in document see if row already exists on SAP table
   '=================================================================
   ReDim gblChanged(UBound(gblStored))
   j = 0
   For n = 1 To UBound(gblStored)
      AlreadyExists = False
      For l = 1 To OrderRowCount
         ' Check if cellname already exists on SAP table
         If UCase(gblStored(n).CellID) = UCase(Trim(RetrievedRecOrderValuesTable.Value(l, "RECID"))) Then
            ' Cell name exists on table so check if changed
            AlreadyExists = True
            If gblStored(n).RecValue = Trim(RetrievedRecOrderValuesTable.Value(l, "RECVALUE")) Then
               ' Value has not changed so do nothing

            Else
               ' Value has changed so record changes
               '========================================================
               ' Exit sub if any entered values are greater than 20 long
               ' (Table field RecValue is only a 20 char field)
               '========================================================
               If Len(gblStored(n).RecValue) > 20 Then
                  MsgBox "Update cancelled - Values have not been updated" _
                     & Chr(13) & Chr(13) & "Value entered in cell name " & gblStored(n).CellID & " is greater than 20 digits long" _
                     & Chr(13) & Chr(13) & "Please re-enter value or contact local support" _
                     , vbExclamation, cnDialogTitleRecCntl
                  Call gblDocuments(Me.CurrentDocument).Doc.ActiveSheet.Range(gblStored(n).CellID).Select
                  Exit Sub
               End If

               j = j + 1
               gblChanged(j).CellID = gblStored(n).CellID
               gblChanged(j).RecValue = gblStored(n).RecValue
               gblChanged(j).SalesNo = gblOrder.SalesOrder
               'MsgBox ("*" + ChangedSalesOrderNumber(j).CellID + "*" + " - " + "*" + gblChanged(j).RecValue + "*")

            End If
            Exit For
         End If

      Next
      If AlreadyExists = False Then
         ' Cell name is not yet on SAP table, check if a value has been input in the cell to be entered into the table
         If gblStored(n).RecValue <> "" Then
            ' Value has changed so record changes
            If Len(gblStored(n).RecValue) > 20 Then
               MsgBox "Update cancelled - Values have not been updated" _
                  & Chr(13) & Chr(13) & "Value entered in cell name " & gblStored(n).CellID & " is greater than 20 digits long" _
                  & Chr(13) & Chr(13) & "Please re-enter value or contact local support" _
                  , vbExclamation, cnDialogTitleRecCntl
               Call gblDocuments(Me.CurrentDocument).Doc.ActiveSheet.Range(gblStored(n).CellID).Select
               Exit Sub
            End If

            j = j + 1
            gblChanged(j).CellID = gblStored(n).CellID
            gblChanged(j).RecValue = gblStored(n).RecValue
            gblChanged(j).SalesNo = gblOrder.SalesOrder
            'MsgBox ("*" + gblChanged(j).CellID + "*" + " - " + "*" + gblChanged(j).RecValue + "*")

         End If

      End If
   Next

   '=========================================
   ' Check whether any changes have been made
   '=========================================
   If j = 0 Then
      If gblExit_File = True Then
         gblExit_File = False
         Exit Sub
      Else
         MsgBox "Values haven't changed - No updating performed.", vbInformation Or vbOKOnly, cnDialogTitleRecCntl
         Exit Sub
      End If
   Else
      If gblExit_File = True Then
         gblExit_File = False
         If EnableUpdate.Enabled = False Then
            Response = MsgBox("Values have changed and have not yet been updated in the database. Do you want to update the database?", vbYesNo + vbQuestion, cnDialogTitleRecCntl)
            If Response = vbNo Then
               Exit Sub
            End If
         Else
            Response = MsgBox("Values have changed and have not yet been updated in the database. This includes some values that were retrieved from a different Order/appendix. Do you want to update the database?" & Chr(13) & Chr(13) & "(Please check the retrieved values are valid before selecting Yes)", vbYesNo + vbQuestion, cnDialogTitleRecCntl)
            If Response = vbNo Then
               Exit Sub
            End If
         End If
      End If

   End If

   ReDim Preserve gblChanged(j)

   '============================================================
   ' Ask for checkno and password before data can be sent to SAP
   '============================================================
   frmCheckPass.Show vbModal
   If gblPersonValidated = False Then
      Exit Sub
   End If

   '========================================================
   ' Check whether person has authorisation to record values
   '========================================================
   If gblUser.CanRecordVals = cnSAPFalse Then
      MsgBox "You have no authorisation to update the recordings" & Chr(13) & Chr(13) & "No updates performed.", vbInformation Or vbOKOnly, cnDialogTitleRecCntl
      Exit Sub
   End If

   '=================================================
   ' Check whether the person is in the correct plant
   '=================================================
   If gblUser.Plant <> gblOrder.Plant Then
      MsgBox "You have no authorisation to update the recordings for plant " & gblOrder.Plant & Chr(13) & Chr(13) & "No updates performed.", vbInformation Or vbOKOnly, cnDialogTitleRecCntl
      Exit Sub
   End If

   SAPUpdateRecValues.Tables.Item("REC_UPDATE").FreeTable

   For j = 1 To UBound(gblChanged)
      gblChanged(j).UserName = gblUser.PersName
      gblChanged(j).UserNo = gblUser.ClockNumber

      With SAPUpdateRecValues.Tables.Item("REC_UPDATE")
         .AppendRow
         .Value(j, "VBELN") = gblOrder.SalesOrder
         .Value(j, "AUFNR") = gblOrder.Number
         .Value(j, "RECID") = UCase(gblChanged(j).CellID)
         .Value(j, "RECVALUE") = UCase(gblChanged(j).RecValue)
         .Value(j, "PERSNAME") = UCase(gblChanged(j).UserName)
         .Value(j, "PERSNUM") = UCase(gblChanged(j).UserNo)
      End With

   Next

   '====================
   ' Test SAP connection
   '====================
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, True) Then
      Exit Sub
   End If

   '==================================
   ' Tell user that update is occuring
   '==================================
   Load frmMsg
   frmMsg.labMsg.Caption = "Updating Values to database..."
   frmMsg.imgFileOpen.Visible = True
   frmMsg.ImgPrinter.Visible = False
   frmMsg.Show
   DoEvents

   '===============================================
   ' Send "Recordings" table to SAP function module
   '===============================================
   If SAPUpdateRecValues.Call = False Then
      MsgBox (SAPUpdateRecValues.Exception)
      Exit Sub
   End If

   '===================================================
   ' Retrieve values back from SAP to populate comments
   '===================================================
   Call RetrieveValuesFromSAP(gblDocuments(Me.CurrentDocument).Doc.ActiveSheet, Me.CurrentDocument)

   frmMsg.Hide
   Unload frmMsg
   Screen.MousePointer = 11
   Call gblDocuments(Me.CurrentDocument).Doc.ActiveSheet.Cells(1, 1).Select
   Screen.MousePointer = 0

End Sub

Private Sub ViewCurrentDocument_Click()

   On Error Resume Next

   '---------------------------------------------------------------------

   If Me.CurrentDocument <> -1 Then
      If gblDocuments(Me.CurrentDocument).Enabled Then
         gblDocuments(Me.CurrentDocument).Doc.Activate
         gblExcel.Visible = True
      End If
   End If

End Sub

Private Sub ViewDocument_Click(Index As Integer)

   On Error Resume Next

   '---------------------------------------------------------------------

   If gblDocuments(Index).Enabled Then
      gblDocuments(Index).Doc.Activate
      gblExcel.Visible = True
   End If

End Sub

