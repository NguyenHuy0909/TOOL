Attribute VB_Name = "Mod_UI"

Option Explicit

'==============================================================================
' UI: Browse template file
'==============================================================================

Public Sub BrowseTemplateFile()

    Dim fd As fileDialog
    Dim selectedPath As String
    Dim wsConfig As Worksheet
    Dim markerCell As Range

    Set wsConfig = ThisWorkbook.Sheets(1)
    Set markerCell = FindTextCell(wsConfig, "#TEMPLATE FILE PATH")
    Set fd = Application.fileDialog(msoFileDialogFilePicker)

    With fd
        .title = "Select Template File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb"

        If .Show = -1 Then
            selectedPath = .SelectedItems(1)
            markerCell.Offset(0, 1).value = selectedPath
        End If
    End With

    Set fd = Nothing

End Sub

'==============================================================================
' UI: Browse spec folders
'==============================================================================

Public Sub BrowseSpecFolders()

    Dim fd As fileDialog
    Dim wsConfig As Worksheet
    Dim markerCell As Range
    Dim targetRange As Range
    Dim targetCell As Range
    Dim selectedPath As String
    Dim answer As VbMsgBoxResult

    Set wsConfig = ThisWorkbook.Sheets(1)
    Set markerCell = FindTextCell(wsConfig, "#SPEC. FOLDER")
    Set targetRange = wsConfig.Range(markerCell.Offset(0, 1), markerCell.Offset(8, 1))
    Set fd = Application.fileDialog(msoFileDialogFolderPicker)

    Do
        Set targetCell = Nothing

        For Each targetCell In targetRange.Cells
            If Trim$(CStr(targetCell.value)) = "" Then Exit For
        Next targetCell

        If targetCell Is Nothing Or Trim$(CStr(targetCell.value)) <> "" Then
            MsgBox "Spec folder range is full.", vbExclamation
            Exit Do
        End If

        With fd
            .title = "Select Spec Folder"
            .AllowMultiSelect = False

            If .Show <> -1 Then Exit Do

            selectedPath = .SelectedItems(1)
        End With

        targetCell.value = selectedPath

        answer = MsgBox("Do you want to select another spec folder?", vbYesNo + vbQuestion, "Continue")
        If answer = vbNo Then Exit Do

    Loop

    Set fd = Nothing

End Sub

'==============================================================================
' Main export
'==============================================================================

Public Sub ExportResultWorkbooksByRpm()

    Dim rpmFolderNames() As String
    Dim rpmIndex As Long
    Dim specIndex As Long
    Dim resultIndex As Long
    Dim bodyIndex As Long

    Dim rpmFolderName As String
    Dim specFolderPath As String
    Dim specName As String
    Dim targetSheetName As String
    Dim resultName As String
    Dim bodyName As String
    Dim csvFilePath As String
    Dim outputWorkbookPath As String

    Dim wbOutput As Workbook
    Dim wbSourceCsv As Workbook
    Dim wsOutput As Worksheet
    Dim wsSourceCsv As Worksheet

    Dim resultStartCells() As String
    Dim nextResultPasteCols() As Long
    Dim resultMarkerCell As Range

    Dim sharedFreqData As Variant
    Dim resultBlockData As Variant
    Dim sharedFreqStartRow As Long
    Dim sharedFreqCol As Long
    Dim resultStartRow As Long
    Dim resultStartCol As Long
    Dim resultBlockWidth As Long

    Dim isSharedFreqPasted As Boolean
    Dim rpmFolderPath As String

    LoadConfig
    PrepareLogSheet

    Application.ScreenUpdating = False

    rpmFolderNames = DetectRpmFolderNames(gSpecFolderPaths)

    For rpmIndex = LBound(rpmFolderNames) To UBound(rpmFolderNames)

        rpmFolderName = rpmFolderNames(rpmIndex)
        outputWorkbookPath = BuildOutputWorkbookPath(gOutputFolderPath, rpmFolderName)

        Debug.Print "----------------------------------------"
        Debug.Print "RPM: " & rpmFolderName
        Debug.Print "Output: " & outputWorkbookPath

        Set wbOutput = CreateOutputWorkbook(gTemplatePath, outputWorkbookPath)

        For specIndex = LBound(gSpecFolderPaths) To UBound(gSpecFolderPaths)

            specFolderPath = gSpecFolderPaths(specIndex)
            specName = GetLastFolderName(specFolderPath)
            targetSheetName = gOutputSheetNames(specIndex)
            rpmFolderPath = EnsureTrailingSlash(specFolderPath) & rpmFolderName & "\"

            Set wsOutput = wbOutput.Worksheets(targetSheetName)

            If Not FolderExists(rpmFolderPath) Then
                AppendLogRow "Missing RPM Folder", specName, rpmFolderName, "", "Folder not found"
                GoTo NextSpec
            End If

            ReDim resultStartCells(LBound(gResultMarkers) To UBound(gResultMarkers))
            ReDim nextResultPasteCols(LBound(gResultNames) To UBound(gResultNames))

            For resultIndex = LBound(gResultMarkers) To UBound(gResultMarkers)
                Set resultMarkerCell = FindTextCell(wsOutput, gResultMarkers(resultIndex))
                resultStartCells(resultIndex) = resultMarkerCell.Offset(1, 0).Address(False, False)
            Next resultIndex

            sharedFreqStartRow = wsOutput.Range(resultStartCells(LBound(resultStartCells))).row
            sharedFreqCol = wsOutput.Range(resultStartCells(LBound(resultStartCells))).Column - 1

            'Call PrepareClearForOutputSheet(wsOutput, resultStartCells)

            For resultIndex = LBound(gResultNames) To UBound(gResultNames)
                resultStartRow = wsOutput.Range(resultStartCells(resultIndex)).row
                resultStartCol = wsOutput.Range(resultStartCells(resultIndex)).Column
                nextResultPasteCols(resultIndex) = resultStartCol
            Next resultIndex

            isSharedFreqPasted = False

            For resultIndex = LBound(gResultNames) To UBound(gResultNames)

                resultName = gResultNames(resultIndex)
                resultStartRow = wsOutput.Range(resultStartCells(resultIndex)).row

                For bodyIndex = LBound(gBodyNames) To UBound(gBodyNames)

                    bodyName = gBodyNames(bodyIndex)
                    csvFilePath = FindCsvByBodyName(specFolderPath, rpmFolderName, bodyName)

                    If csvFilePath = "" Then
                        AppendLogRow "Missing CSV", specName, rpmFolderName, bodyName, "CSV not found"
                        GoTo NextBody
                    End If

                    Set wbSourceCsv = Workbooks.Open(csvFilePath)
                    Set wsSourceCsv = wbSourceCsv.Worksheets(1)

                    If Not isSharedFreqPasted Then
                        sharedFreqData = BuildSharedFrequencyColumn(wsSourceCsv)
                        PasteSharedFrequencyColumn wsOutput, sharedFreqStartRow, sharedFreqCol, sharedFreqData
                        isSharedFreqPasted = True
                    End If

                    resultBlockData = BuildResultValueBlock(wsSourceCsv, resultName)

                    If Not IsEmpty(resultBlockData) Then
                        PasteArrayBlock wsOutput, resultStartRow, nextResultPasteCols(resultIndex), resultBlockData
                        resultBlockWidth = GetArrayColumnCount(resultBlockData)
                        nextResultPasteCols(resultIndex) = nextResultPasteCols(resultIndex) + resultBlockWidth
                    Else
                        AppendLogRow "Missing Result Columns", specName, rpmFolderName, bodyName, resultName
                    End If

                    wbSourceCsv.Close SaveChanges:=False

NextBody:
                Next bodyIndex

            Next resultIndex

NextSpec:
        Next specIndex

        wbOutput.Save
        wbOutput.Close SaveChanges:=True

    Next rpmIndex

    Application.ScreenUpdating = True
    Debug.Print "===== Export finished ====="

End Sub

'==============================================================================
' Optional clear
'==============================================================================

Public Sub PrepareClearForOutputSheet(ByVal wsOutput As Worksheet, ByVal resultStartCells As Variant)

    Dim i As Long

    For i = LBound(resultStartCells) To UBound(resultStartCells)
        ClearResultBlockArea wsOutput, CStr(resultStartCells(i)), 300, 25
    Next i

End Sub
