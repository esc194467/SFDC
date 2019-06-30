Attribute VB_Name = "basRecControl"
Option Explicit

Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFilename As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Function OpenRecordingDocument(lpInFilename As String) As Integer

   On Error Resume Next

   Dim nDoc                                    As Integer
   Dim bContinue                               As Boolean

   Dim szFilename                              As String
   Dim szCtrlFile                              As String

   '------------------------------------------------------------------------------

   OpenRecordingDocument = -1
   Screen.MousePointer = vbDefault
   bContinue = False

   '======================
   ' Obtain Sales Order No
   '======================
   bContinue = True

   Load frmMsg
   frmMsg.labMsg.Caption = "Opening Recording Document Template..."
   frmMsg.imgFileOpen.Visible = True
   frmMsg.ImgPrinter.Visible = False
   frmMsg.Show
   Call SetWindowPos(frmMsg.hwnd, cnHWND_TOPMOST, 0, 0, 0, 0, cnSWP_NOMOVE Or cnSWP_NOSIZE)

   DoEvents

   szFilename = lpInFilename

   If bContinue Then
      '==================================
      ' Check if document is already open
      '==================================
      If Not IsDocumentAlreadyOpen(szFilename) Then
         If StartupExcel() Then
            gblExcel.Visible = False
            nDoc = GetFirstFreeDocument()
            If nDoc <> -1 Then
               gblDocuments(nDoc).Enabled = True
               gblDocuments(nDoc).Template = lpInFilename
               gblDocuments(nDoc).Name = szFilename

               Set gblDocuments(nDoc).Doc = gblExcel.Workbooks.Open(szFilename)
               frmMsg.Show
               If Not gblDocuments(nDoc).Doc Is Nothing Then
                  frmRecControl.CurrentDocument = nDoc
                  OpenRecordingDocument = nDoc
                  ' Populate header information
                  gblDocuments(nDoc).Doc.Names("SpecificSalesOrderNo").RefersToRange.Value = gblOrder.SalesOrder
                  gblDocuments(nDoc).Doc.Names("SpecificCustomer").RefersToRange.Value = gblOrder.Customer
                  gblDocuments(nDoc).Doc.Names("SpecificEngineModuleType").RefersToRange.Value = gblOrder.EngineType
                  gblDocuments(nDoc).Doc.Names("SpecificEngNo").RefersToRange.Value = gblOrder.EngSerialNo
                  gblDocuments(nDoc).Doc.Names("SpecificModNo").RefersToRange.Value = gblOrder.ModuleNumber

                  Call RetrieveValuesFromSAP(gblDocuments(nDoc).Doc.ActiveSheet, nDoc)
                  Call ProcessSAPMethods(gblDocuments(nDoc).Doc.ActiveSheet, nDoc)
                  
               End If

            End If
            gblExcel.Visible = True

         End If
      Else
         MsgBox "This document is already open..."
      End If
   End If

   frmMsg.Hide
   Unload frmMsg

End Function

'*=================================================================*
'*
'*
'*
'*
'*=================================================================*

Public Function PrintRecordingDocument(nDoc As Integer) As Boolean

   On Error Resume Next

   '---------------------------------------------------------------

   PrintRecordingDocument = False

   If gblDocuments(nDoc).Enabled Then
      Load frmMsg
      frmMsg.labMsg.Caption = "Printing " + ShortenFilenameLevels(gblDocuments(nDoc).Name, 1)
      frmMsg.ImgPrinter.Visible = True
      frmMsg.imgFileOpen.Visible = False
      frmMsg.Show
      DoEvents
      Call gblDocuments(nDoc).Doc.PrintOut
      frmMsg.Hide
      Unload frmMsg

      PrintRecordingDocument = True
   End If

End Function

'*=================================================================*
'*
'*
'*
'*
'*=================================================================*

Public Function CloseRecordingDocument(nDoc As Integer, frm As Form) As Boolean

   On Error Resume Next

   Dim i                               As Integer
   Dim bClosed                         As Boolean

   '---------------------------------------------------------------

   CloseRecordingDocument = True

   bClosed = False
   If gblDocuments(nDoc).Enabled Then
      CloseRecordingDocument = True
      gblDocuments(nDoc).Enabled = False
      Call gblDocuments(nDoc).Doc.Close(False)
      frm.ViewDocument(nDoc).Visible = False
      bClosed = True
   End If

   If bClosed Then
      If frm.CurrentDocument = nDoc Then
         frm.CurrentDocument = -1
         For i = 0 To gblDocCount - 1
            If gblDocuments(i).Enabled Then
               frm.CurrentDocument = i
               Exit For
            End If
         Next i
      End If
   End If

End Function

'*=================================================================*
'*
'*
'*
'*
'*=================================================================*

Public Function GetFirstFreeDocument() As Integer

   On Error Resume Next

   Dim i                           As Integer
   Dim nDoc                        As Integer
   Dim nDocs                       As Integer

   '---------------------------------------------------------------------

   nDoc = -1
   Err.Clear
   nDocs = UBound(gblDocuments)
   If Err.Number = 0 Then
      For i = 0 To UBound(gblDocuments)
         If Not gblDocuments(i).Enabled Then
            nDoc = i
            Exit For
         End If
      Next i
   End If

   If nDoc = -1 Then
      If gblDocCount <= cnMAX_DOCUMENTS Then
         nDoc = gblDocCount
         Err.Clear
         ReDim Preserve gblDocuments(nDoc) As tagWorkBook
         If Err.Number = 0 Then
            gblDocCount = gblDocCount + 1
         Else
            nDoc = -1
         End If
      Else
         MsgBox "The maximum No. of open documents is " + Trim(Str(cnMAX_DOCUMENTS + 1))
      End If
   End If

   GetFirstFreeDocument = nDoc

End Function

'*=================================================================*
'*
'*
'*
'*
'*=================================================================*

Public Function StartupExcel() As Boolean

   On Error Resume Next

   '---------------------------------------------------------------------

   If gblExcel Is Nothing Then
      Set gblExcel = New Excel.Application
   Else
      Err.Clear
      gblExcel.Visible = gblExcel.Visible
      gblExcel.ScreenUpdating = True
      If Err.Number <> 0 Then
         Set gblExcel = Nothing
         Set gblExcel = New Excel.Application
      End If
   End If

   If Not gblExcel Is Nothing Then
      Set frmRecControl.gblAppExcel = gblExcel
      StartupExcel = True
   Else
      StartupExcel = False
   End If

End Function

'*=================================================================*
'*
'*
'*
'*
'*=================================================================*

Public Sub ShutdownExcel(frm As Form)

   On Error Resume Next

   '---------------------------------------------------------------------

   If Not frm Is Nothing Then
      If Not frm.gblAppExcel Is Nothing Then
         Set frm.gblAppExcel = Nothing
      End If
   End If

   If Not gblExcel Is Nothing Then
      gblExcel.Visible = False
      gblExcel.Quit
      Set gblExcel = Nothing
   End If

End Sub

'*=================================================================*
'*
'*
'*
'*
'*=================================================================*

Public Function StripTerminator(lpVar As String) As String

   Dim nPos As Integer

   nPos = InStr(lpVar, Chr(0))
   If nPos = 0 Then
      StripTerminator = lpVar
   Else
      StripTerminator = Left(lpVar, nPos - 1)
   End If

End Function

Public Function StripTerminator2(lpVar As String) As String

   Dim nPos As Integer

   nPos = InStr(lpVar, Chr(0) + Chr(0))
   If nPos = 0 Then
      StripTerminator2 = lpVar
   Else
      StripTerminator2 = Left(lpVar, nPos - 1)
   End If

End Function

'*=================================================================*
'*
'*
'*
'*
'*=================================================================*

Public Function Separate(ByVal cField As String, ByVal cSep As String, aryFields() As String) As Integer

   Dim nCsvCnt                         As Integer      ' CSV Counter
   Dim DrPos                           As Integer      ' Position Var

   ReDim aryFields(0) As String

   ' -------------------------------------------------------------------------

   On Error Resume Next

   nCsvCnt = 0
   Do While Len(cField) > 0
      Do While Left(cField, 1) = " "
         cField = Right(cField, Len(cField) - 1)
      Loop

      If Len(Trim(cField)) > 0 Then
         DrPos = InStr(cField, cSep)
         nCsvCnt = nCsvCnt + 1
         If DrPos = 0 Then
            ReDim Preserve aryFields(nCsvCnt) As String
            aryFields(nCsvCnt) = StripTerminator(cField)
            cField = ""
         Else
            ReDim Preserve aryFields(nCsvCnt) As String
            aryFields(nCsvCnt) = StripTerminator(Left(cField, DrPos - 1))
            cField = Right(cField, Len(cField) - DrPos)
         End If
      End If
   Loop

   Separate = nCsvCnt

End Function

'*=================================================================*
'*
'*
'*
'*
'*=================================================================*

Public Function FileExists(lpFile As String) As Boolean

   On Error Resume Next

   Dim szEntry                         As String

   '------------------------------------------------------------------

   FileExists = False
   szEntry = Dir(lpFile, vbDirectory)
   If Len(szEntry) <> 0 Then
      If Right(UCase(lpFile), Len(szEntry)) = UCase(szEntry) Then
         FileExists = True
      End If
   End If

End Function

'*=================================================================*
'*
'*
'*
'*
'*=================================================================*

Public Function ShortenFilename(lpFile As String) As String

   On Error Resume Next

   Dim i                               As Integer
   Dim szFilename                      As String

   '------------------------------------------------------------------

   szFilename = ""
   For i = Len(lpFile) To 1 Step -1
      If Mid(lpFile, i, 1) <> "\" Then
         szFilename = Mid(lpFile, i, 1) + szFilename
      Else
         Exit For
      End If
   Next i

   ShortenFilename = szFilename

End Function

'*=================================================================*
'*
'*
'*
'*
'*=================================================================*

Public Function ShortenFilenameLevels(lpFile As String, nLevel As Integer) As String

   On Error Resume Next

   Dim i                               As Integer
   Dim nCnt                            As Integer
   Dim szFilename                      As String

   '------------------------------------------------------------------

   nCnt = 0
   szFilename = ""
   For i = Len(lpFile) To 1 Step -1
      If Mid(lpFile, i, 1) <> "\" Then
         szFilename = Mid(lpFile, i, 1) + szFilename
      Else
         'szFilename = "\" + szFilename
         szFilename = szFilename
         nCnt = nCnt + 1
      End If

      If nCnt = nLevel Then
         Exit For
      End If
   Next i

   ShortenFilenameLevels = szFilename

End Function

'*=================================================================*
'*
'*
'*
'*
'*=================================================================*

Public Function RemoveExtension(lpFilename As String) As String

   On Error Resume Next

   Dim i                               As Integer
   Dim bExt                            As Boolean
   Dim szFilename                      As String

   '------------------------------------------------------------------

   szFilename = ""
   bExt = True
   For i = Len(lpFilename) To 1 Step -1
      If Not bExt Then
         szFilename = Mid(lpFilename, i, 1) + szFilename
      Else
         If Mid(lpFilename, i, 1) = "." Then
            bExt = False
         End If
      End If
   Next i

   RemoveExtension = szFilename

End Function

Public Function IsDocumentAlreadyOpen(lpFile As String) As Boolean

   On Error Resume Next

   Dim i                           As Integer
   Dim nDocs                       As Integer
   Dim bFnd                        As Boolean

   '----------------------------------------------------------------------------

   bFnd = False
   Err.Clear
   nDocs = UBound(gblDocuments)
   If Err.Number = 0 Then
      For i = 0 To UBound(gblDocuments)
         If gblDocuments(i).Enabled Then
            If lpFile = gblDocuments(i).Name Then
               bFnd = True
               Exit For
            End If
         End If
      Next i
   End If

   IsDocumentAlreadyOpen = bFnd

End Function

Public Function CountDocumentsOpen() As Integer

   On Error Resume Next

   Dim i                           As Integer
   Dim nDocs                       As Integer
   Dim nCnt                        As Integer

   '----------------------------------------------------------------------------

   nCnt = 0
   Err.Clear
   nDocs = UBound(gblDocuments)
   If Err.Number = 0 Then
      For i = 0 To UBound(gblDocuments)
         If gblDocuments(i).Enabled Then
            nCnt = nCnt + 1
         End If
      Next i
   End If

   CountDocumentsOpen = nCnt

End Function

Public Sub BringDocumentToFront(lpFile As String)

   On Error Resume Next

   Dim i                           As Integer
   Dim nDocs                       As Integer
   Dim bFnd                        As Boolean

   '----------------------------------------------------------------------------

   Err.Clear
   nDocs = UBound(gblDocuments)
   If Err.Number = 0 Then
      For i = 0 To UBound(gblDocuments)
         If gblDocuments(i).Enabled Then
            If lpFile = gblDocuments(i).Name Then
               frmRecControl.CurrentDocument = i
               Exit For
            End If
         End If
      Next i
   End If

End Sub

Sub RetrieveValuesFromSAP(wsSheet As Worksheet, DocNo As Integer)

   On Error Resume Next

   Dim NameCount As Integer
   Dim j As Integer
   Dim i As Integer
   Dim n As Integer
   Dim x As Integer
   Dim OrderRowCount As Integer
   Dim DateRowCount As Integer
   Dim DocumentTitle As String
   Dim TrimmedUserName As String
   Dim TrimmedPersNum As String
   Dim FirstName As String
   Dim LatestValuesRequired As Boolean
   Dim LastName As String
   Dim RDNameCount As Integer
   Dim DateRowRetrieved As Boolean
   Dim OrderRowExists As Boolean
   Dim OrderRowNumber As Integer
   Dim Response As Integer
   Dim Duplicate_Records_Exist As Boolean
   Dim DateRowNumber As Integer
   Dim CellName As String

   j = 0
   RDNameCount = 0
   '===================================
   ' Determine No. of names in document
   '===================================
   NameCount = gblDocuments(DocNo).Doc.Names.Count

   '==================================================================================
   ' For each cell that can be changed either by user input or auto calculation,
   ' check if the name is greater than 80 characters (table field is only 80 char long
   '==================================================================================
   For n = 1 To NameCount
      If Mid(UCase(gblDocuments(DocNo).Doc.Names(n).Name), 1, 2) = "RD" Then

         If Len(gblDocuments(DocNo).Doc.Names(n).Name) > 80 Then
            MsgBox "Recording Document Template contains named cells that are greater than 80 digits in length" _
               & Chr(13) & Chr(13) & "Recording Document Template will be closed" _
               & Chr(13) & Chr(13) & "Please contact local support" _
               , vbExclamation, cnDialogTitleRecCntl
            Call CloseRecordingDocument(DocNo, frmRecControl)
            Exit Sub
         End If
      End If
   Next

   ReDim Preserve gblRetrieval(NameCount)

   '====================================
   ' Ensure export table is free of data
   '====================================
   SAPRetrieveRecValues.Tables.Item("CELL_INFO").FreeTable

   '===================================================
   ' Obtain names of cells that require value retrieval
   '===================================================
   For n = 1 To NameCount

      CellName = gblDocuments(DocNo).Doc.Names(n).Name

      If Mid(UCase(CellName), 1, 6) = "RD_SAP" Then

            'the RecDoc cell references a SAP method cell
            'so put data into the table that is later sent to SAP
            j = j + 1
            ' Put each name into gblRetrieval array
            gblRetrieval(j).CellID = UCase(gblDocuments(DocNo).Doc.Names(n).Name)
            With SAPRetrieveRecValues.Tables.Item("CELL_INFO")
               .AppendRow
               .Value(j, "VBELN") = gblOrder.SalesOrder
               .Value(j, "AUFNR") = gblOrder.Number
               .Value(j, "RECID") = UCase(gblDocuments(DocNo).Doc.Names(n).Name)
            End With
         
      ElseIf Mid(UCase(CellName), 1, 2) = "RD" Then
      
          If Mid(gblDocuments(DocNo).Doc.Names(n).RefersToRange.Formula, 1, 1) <> "=" Then
            'No formula in this cell
            'so put data into the table that is later sent to SAP
            j = j + 1
            ' Put each name into gblRetrieval array
            gblRetrieval(j).CellID = UCase(gblDocuments(DocNo).Doc.Names(n).Name)
            With SAPRetrieveRecValues.Tables.Item("CELL_INFO")
               .AppendRow
               .Value(j, "VBELN") = gblOrder.SalesOrder
               .Value(j, "AUFNR") = gblOrder.Number
               .Value(j, "RECID") = UCase(gblDocuments(DocNo).Doc.Names(n).Name)
            End With
         Else
            'Formula in this cell and cell name does not begin RD_SAP
            'so do not send for retrieval
         End If
     
      End If
   Next

   RDNameCount = j
   ReDim Preserve gblRetrieval(j)
   '======================================
   ' Check if any user entered names exist
   '======================================
   If j = 0 Then
      ' No values will be retrieved
      Exit Sub
   End If

   '========================================
   ' Remove existing data from import tables
   '========================================
   RetrievedRecOrderValuesTable.FreeTable
   RetrievedRecDateValuesTable.FreeTable

   '====================
   ' Test SAP Connection
   '====================
   If Not SAPisConnected(gblSAP, gblInstallation, cnCheckRunning, True) Then
      Call CloseRecordingDocument(DocNo, frmRecControl)
      Exit Sub
   End If

   '==============================================
   ' Send "Cell_Info" table to SAP function module
   '==============================================
   ' Obtain tables back from SAP function module
   '============================================
   If SAPRetrieveRecValues.Call = False Then
      MsgBox (SAPRetrieveRecValues.Exception)
      Exit Sub
   End If
   '===================================================================================================
   ' Determine how many rows were obtained from the "Rec_Order_Retrieve" and "Rec_Date_Retrieve" tables
   '===================================================================================================
   OrderRowCount = RetrievedRecOrderValuesTable.RowCount
   DateRowCount = RetrievedRecDateValuesTable.RowCount

   j = 0
   ReDim gblRecordings(UBound(gblRetrieval))
   DateRowRetrieved = False
   For n = 1 To UBound(gblRetrieval)

      '=====================================================================
      ' Check whether name already has a value recorded for the order number
      '=====================================================================
      OrderRowExists = False
      For i = 1 To OrderRowCount
         If UCase(gblRetrieval(n).CellID) = Trim(RetrievedRecOrderValuesTable.Value(i, "RECID")) Then
            ' Values exist for the orderno so store them
            j = j + 1
            OrderRowExists = True
            ' Populate gblRecordings Array with data from OrderValuesTable
            gblRecordings(j).OrderNo = Trim(RetrievedRecOrderValuesTable.Value(i, "AUFNR"))
            gblRecordings(j).CellID = Trim(RetrievedRecOrderValuesTable.Value(i, "RECID"))
            gblRecordings(j).RecValue = Trim(RetrievedRecOrderValuesTable.Value(i, "RECVALUE"))
            gblRecordings(j).UserName = Trim(RetrievedRecOrderValuesTable.Value(i, "PERSNAME"))
            gblRecordings(j).UserNo = Trim(RetrievedRecOrderValuesTable.Value(i, "PERSNUM"))
            gblRecordings(j).Date = Trim(RetrievedRecOrderValuesTable.Value(i, "ZDATE"))
            gblRecordings(j).Time = Trim(RetrievedRecOrderValuesTable.Value(i, "ZTIME"))

            Exit For
         Else
         End If
      Next
      '================================================================================================
      ' Check whether any names that didn't have a value on this order had a value on a different order
      '================================================================================================

      If OrderRowExists = False Then
         For i = 1 To DateRowCount
            If UCase(gblRetrieval(n).CellID) = Trim(RetrievedRecDateValuesTable.Value(i, "RECID")) Then
               j = j + 1
               DateRowRetrieved = True
               ' Populate gblRecordings Array with data from DateValuesTable
               gblRecordings(j).OrderNo = Trim(RetrievedRecDateValuesTable.Value(i, "AUFNR"))
               gblRecordings(j).CellID = Trim(RetrievedRecDateValuesTable.Value(i, "RECID"))
               gblRecordings(j).RecValue = Trim(RetrievedRecDateValuesTable.Value(i, "RECVALUE"))
               gblRecordings(j).UserName = Trim(RetrievedRecDateValuesTable.Value(i, "PERSNAME"))
               gblRecordings(j).UserNo = Trim(RetrievedRecDateValuesTable.Value(i, "PERSNUM"))
               gblRecordings(j).Date = Trim(RetrievedRecDateValuesTable.Value(i, "ZDATE"))
               gblRecordings(j).Time = Trim(RetrievedRecDateValuesTable.Value(i, "ZTIME"))
               Exit For
            End If

         Next
      End If
   Next
   If DateRowRetrieved Then
      ' Display message to check values before updating and disable Update option
      MsgBox "Some of the values have been retrieved from a previous Order/Appendix" _
         & Chr(13) & Chr(13) & "Please check these values are valid before updating this Appendix" _
         , vbExclamation, cnDialogTitleRecCntl

      If UCase(Trim(gblPlantCfg.CanUpdRecDocs)) = cnSAPTrue Then
         frmRecControl.UpdateValues.Enabled = False
         frmRecControl.EnableUpdate.Enabled = True
      Else
         frmRecControl.UpdateValues.Enabled = False
         frmRecControl.EnableUpdate.Enabled = False

      End If
   Else
      If UCase(Trim(gblPlantCfg.CanUpdRecDocs)) = cnSAPTrue Then
         frmRecControl.UpdateValues.Enabled = True
         frmRecControl.EnableUpdate.Enabled = False

      Else
         frmRecControl.UpdateValues.Enabled = False
         frmRecControl.EnableUpdate.Enabled = False

      End If

   End If

   ReDim Preserve gblRecordings(j)
   If j = 0 Then
      ' nothing to update so do nothing
   Else
      Screen.MousePointer = 11
      Call gblDocuments(DocNo).Doc.ActiveSheet.Cells(1, 1).Select
      Err.Clear
      '=======================================================
      ' Unprotect sheet to allow data input to comments object
      '=======================================================
      Call gblDocuments(DocNo).Doc.ActiveSheet.Unprotect
      '=====================================================================
      ' For each row retrieved populate the related cell and comment objects
      '=====================================================================
      For n = 1 To UBound(gblRecordings)
         gblDocuments(DocNo).Doc.Names(gblRecordings(n).CellID).RefersToRange.Value = gblRecordings(n).RecValue
         TrimmedUserName = Trim(gblRecordings(n).UserName)
         LastName = Mid(TrimmedUserName, 1, (InStr(1, TrimmedUserName, " ") - 1))
         FirstName = Trim(Mid(TrimmedUserName, (InStr(1, TrimmedUserName, " ") + 1)))
         TrimmedPersNum = Format(gblRecordings(n).UserNo, "########")
         gblDocuments(DocNo).Doc.Names(gblRecordings(n).CellID).RefersToRange.Comment.Delete
         With gblDocuments(DocNo).Doc.Names(gblRecordings(n).CellID).RefersToRange.AddComment
            .Text "Last Updated By: " & FirstName & " " & LastName & " (" & TrimmedPersNum & ")" & Chr(10) & "Date: " & gblRecordings(n).Date & " " & gblRecordings(n).Time & " (Order No: " & Format(gblRecordings(n).OrderNo, "############") & ")"
            .Visible = False
            .Shape.Height = 25
            .Shape.Width = 250
         End With
      Next
      '==================
      ' Protect the sheet
      '==================
      Call gblDocuments(DocNo).Doc.ActiveSheet.Protect("", True, True, True)
      Screen.MousePointer = 0
   End If

End Sub
Public Function CheckRecValueChanges() As Boolean

   Dim l As Integer

   CheckRecValueChanges = False
   For l = 1 To UBound(gblStored)
      If gblStored(l).RecValue <> gblDocuments(0).Doc.Names(gblStored(l).CellID).RefersToRange.Value Then
         CheckRecValueChanges = True
      End If
   Next
End Function

Public Sub ProcessSAPMethods(wsSheet As Worksheet, nDoc As Integer)

   'Call the routine for executing the Methods (cell names beginning SAP......)
   Call LoadNamedCellArray(wsSheet, nDoc)
   'execute the methods for each cell as appropriate
   Call ProcessNamedCells
   'update the value in the named cells
   Call UpdateNamedCells(wsSheet, nDoc)

End Sub
