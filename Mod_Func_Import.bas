Attribute VB_Name = "Mod_Func_Import"

Option Explicit

'==============================================================================
' Path helpers
'==============================================================================

Public Function EnsureTrailingSlash(ByVal folderPath As String) As String
    If Right$(folderPath, 1) = "\" Then
        EnsureTrailingSlash = folderPath
    Else
        EnsureTrailingSlash = folderPath & "\"
    End If
End Function

Public Function GetLastFolderName(ByVal folderPath As String) As String

    Dim normalizedPath As String
    Dim parts() As String

    normalizedPath = folderPath
    If Right$(normalizedPath, 1) = "\" Then
        normalizedPath = Left$(normalizedPath, Len(normalizedPath) - 1)
    End If

    parts = Split(normalizedPath, "\")
    GetLastFolderName = parts(UBound(parts))

End Function

Public Function BuildOutputWorkbookPath(ByVal outputFolderPath As String, ByVal rpmFolderName As String) As String
    BuildOutputWorkbookPath = EnsureTrailingSlash(outputFolderPath) & BuildOutputFileNameByRule(rpmFolderName)
End Function

Public Function IsNumericFolderName(ByVal folderName As String) As Boolean

    Dim i As Long
    Dim ch As String

    If Len(folderName) = 0 Then Exit Function

    For i = 1 To Len(folderName)
        ch = Mid$(folderName, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    IsNumericFolderName = True

End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    FolderExists = (Len(Dir(folderPath, vbDirectory)) > 0)
End Function

Public Function DetectRpmFolderNames(ByRef specFolderPaths() As String) As String()

    Dim detected() As String
    Dim specIndex As Long
    Dim basePath As String
    Dim folderName As String
    Dim seen As Boolean
    Dim i As Long

    For specIndex = LBound(specFolderPaths) To UBound(specFolderPaths)

        basePath = EnsureTrailingSlash(specFolderPaths(specIndex))
        folderName = Dir(basePath & "*", vbDirectory)

        Do While folderName <> ""
            If folderName <> "." And folderName <> ".." Then
                If (GetAttr(basePath & folderName) And vbDirectory) = vbDirectory Then
                    If IsNumericFolderName(folderName) Then

                        seen = False
                        If IsStringArrayAllocated(detected) Then
                            For i = LBound(detected) To UBound(detected)
                                If detected(i) = folderName Then
                                    seen = True
                                    Exit For
                                End If
                            Next i
                        End If

                        If Not seen Then
                            AppendToStringArray detected, folderName
                        End If
                    End If
                End If
            End If
            folderName = Dir
        Loop

    Next specIndex

    DetectRpmFolderNames = detected

End Function

Public Function FindCsvByBodyName(ByVal specFolderPath As String, ByVal rpmFolderName As String, ByVal bodyName As String) As String

    Dim targetFolder As String
    Dim fileName As String

    targetFolder = EnsureTrailingSlash(specFolderPath) & rpmFolderName & "\"

    If Not FolderExists(targetFolder) Then Exit Function

    fileName = Dir(targetFolder & "*.csv")
    Do While fileName <> ""
        If InStr(1, UCase$(fileName), "_" & UCase$(bodyName) & "_", vbTextCompare) > 0 Then
            FindCsvByBodyName = targetFolder & fileName
            Exit Function
        End If
        fileName = Dir
    Loop

End Function

'==============================================================================
' CSV parsing helpers
'==============================================================================

Public Function FindResultColumnIndexes(ByVal wsCsv As Worksheet, ByVal resultName As String) As Variant

    Dim lastCol As Long
    Dim colIndex As Long
    Dim indexes() As Long
    Dim countFound As Long
    Dim groupHeader As String
    Dim itemHeader As String

    lastCol = wsCsv.Cells(2, wsCsv.Columns.Count).End(xlToLeft).Column

    For colIndex = 2 To lastCol

        groupHeader = NormalizeKeyText(CStr(wsCsv.Cells(2, colIndex).value))
        itemHeader = NormalizeKeyText(CStr(wsCsv.Cells(4, colIndex).value))

        If groupHeader = NormalizeKeyText(resultName) Then
            If InStr(1, itemHeader, "ALL PANEL", vbTextCompare) = 0 Then
                countFound = countFound + 1
                ReDim Preserve indexes(1 To countFound)
                indexes(countFound) = colIndex
            End If
        End If

    Next colIndex

    FindResultColumnIndexes = indexes

End Function

Public Function GetCsvLastRow(ByVal wsCsv As Worksheet) As Long
    GetCsvLastRow = wsCsv.Cells(wsCsv.Rows.Count, 1).End(xlUp).row
End Function

Public Function BuildResultValueBlock(ByVal wsCsv As Worksheet, ByVal resultName As String) As Variant

    Dim resultCols As Variant
    Dim lastRow As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim dataBlock() As Variant
    Dim rowIndex As Long
    Dim srcIndex As Long
    Dim outCol As Long

    resultCols = FindResultColumnIndexes(wsCsv, resultName)

    If IsEmpty(resultCols) Then Exit Function

    lastRow = GetCsvLastRow(wsCsv)
    rowCount = lastRow - 1
    colCount = UBound(resultCols)

    ReDim dataBlock(1 To rowCount, 1 To colCount)

    For srcIndex = LBound(resultCols) To UBound(resultCols)
        outCol = srcIndex
        For rowIndex = 2 To lastRow
            dataBlock(rowIndex - 1, outCol) = wsCsv.Cells(rowIndex, resultCols(srcIndex)).value
        Next rowIndex
    Next srcIndex

    BuildResultValueBlock = dataBlock

End Function

Public Function BuildSharedFrequencyColumn(ByVal wsCsv As Worksheet) As Variant

    Dim lastRow As Long
    Dim rowCount As Long
    Dim freqData() As Variant
    Dim rowIndex As Long

    lastRow = GetCsvLastRow(wsCsv)
    rowCount = lastRow - 1

    ReDim freqData(1 To rowCount, 1 To 1)

    For rowIndex = 2 To lastRow
        freqData(rowIndex - 1, 1) = wsCsv.Cells(rowIndex, 1).value
    Next rowIndex

    BuildSharedFrequencyColumn = freqData

End Function

'==============================================================================
' Array / paste helpers
'==============================================================================

Public Function GetArrayColumnCount(ByRef dataArr As Variant) As Long
    GetArrayColumnCount = UBound(dataArr, 2)
End Function

Public Function GetArrayRowCount(ByRef dataArr As Variant) As Long
    GetArrayRowCount = UBound(dataArr, 1)
End Function

Public Sub PasteArrayBlock(ByVal wsTarget As Worksheet, ByVal startRow As Long, ByVal startCol As Long, ByRef dataArr As Variant)
    wsTarget.Cells(startRow, startCol).Resize(GetArrayRowCount(dataArr), GetArrayColumnCount(dataArr)).value = dataArr
End Sub

Public Sub PasteSharedFrequencyColumn(ByVal wsTarget As Worksheet, ByVal startRow As Long, ByVal targetCol As Long, ByRef freqData As Variant)
    wsTarget.Cells(startRow, targetCol).Resize(UBound(freqData, 1), 1).value = freqData
End Sub

'==============================================================================
' Workbook helpers
'==============================================================================

Public Function CreateOutputWorkbook(ByVal templatePath As String, ByVal outputWorkbookPath As String) As Workbook

    Dim wb As Workbook

    Set wb = Workbooks.Open(templatePath)

    Application.DisplayAlerts = False
    wb.SaveAs fileName:=outputWorkbookPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True

    Set CreateOutputWorkbook = wb

End Function

'==============================================================================
' Clear helpers
'==============================================================================

Public Sub ClearResultBlockArea(ByVal wsTarget As Worksheet, ByVal startCell As String, ByVal clearCols As Long, ByVal clearRows As Long)
    wsTarget.Range(startCell).Resize(clearRows, clearCols).ClearContents
End Sub

'==============================================================================
' Log helpers
'==============================================================================

Public Function GetLogWorksheet() As Worksheet
    Set GetLogWorksheet = ThisWorkbook.Worksheets("log")
End Function

Public Sub PrepareLogSheet()

    Dim wsLog As Worksheet

    Set wsLog = GetLogWorksheet()
    wsLog.Cells.ClearContents

    wsLog.Range("A1:E1").value = Array("Type", "Spec", "RPM", "Body", "Message")

End Sub

Public Sub AppendLogRow(ByVal logType As String, ByVal specName As String, ByVal rpmFolderName As String, ByVal bodyName As String, ByVal messageText As String)

    Dim wsLog As Worksheet
    Dim nextRow As Long

    Set wsLog = GetLogWorksheet()
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).row + 1
    If nextRow < 2 Then nextRow = 2

    wsLog.Cells(nextRow, 1).value = logType
    wsLog.Cells(nextRow, 2).value = specName
    wsLog.Cells(nextRow, 3).value = rpmFolderName
    wsLog.Cells(nextRow, 4).value = bodyName
    wsLog.Cells(nextRow, 5).value = messageText

End Sub
Public Function BuildOutputFileNameByRule(ByVal rpmFolderName As String) As String

    Dim fileName As String
    Dim i As Long

    If IsStringArrayAllocated(gNamingRuleValues) Then
        For i = LBound(gNamingRuleValues) To UBound(gNamingRuleValues)
            If Trim$(gNamingRuleValues(i)) <> "" Then
                If fileName <> "" Then fileName = fileName & "_"
                fileName = fileName & Trim$(gNamingRuleValues(i))
            End If
        Next i
    End If

    If fileName <> "" Then
        fileName = fileName & "_" & rpmFolderName & "rpm"
    Else
        fileName = rpmFolderName & "rpm"
    End If

    BuildOutputFileNameByRule = fileName & ".xlsx"

End Function
