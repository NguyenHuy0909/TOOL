Attribute VB_Name = "Mod_Init"

Option Explicit

' ===== Global config =====
Public gTemplatePath As String
Public gSpecFolderPaths() As String
Public gBodyNames() As String
Public gResultNames() As String
Public gResultMarkers() As String
Public gOutputSheetNames() As String
Public gNamingRuleValues() As String
Public gOutputFolderPath As String

Public Sub LoadConfig()

    Dim wsConfig As Worksheet
    Dim markerCell As Range

    Set wsConfig = ThisWorkbook.Sheets(1)

    Set markerCell = FindTextCell(wsConfig, "#TEMPLATE FILE PATH")
    gTemplatePath = ReadCellToRight(markerCell)

    Set markerCell = FindTextCell(wsConfig, "#SPEC. FOLDER")
    gSpecFolderPaths = ReadColumnValuesBelow(markerCell, 0, 1, 11)

    Set markerCell = FindTextCell(wsConfig, "#BODY NAME")
    gBodyNames = ReadRowValuesToRight(markerCell)

    Set markerCell = FindTextCell(wsConfig, "#RESULT NAME")
    gResultNames = ReadRowValuesToRight(markerCell)

    Set markerCell = FindTextCell(wsConfig, "#RESULT MARKER")
    gResultMarkers = ReadRowValuesToRight(markerCell)

    Set markerCell = FindTextCell(wsConfig, "#SHEET GROUPS")
    gOutputSheetNames = ReadRowValuesToRight(markerCell)

    Set markerCell = FindTextCell(wsConfig, "#NAMING RULE")
    gNamingRuleValues = ReadRowValuesToRight(markerCell, 1)

    Set markerCell = FindTextCell(wsConfig, "#OUTPUT DIRECTORY")
    gOutputFolderPath = ReadCellToRight(markerCell)

    DebugPrintConfig

End Sub

Public Sub DebugPrintConfig()

    Debug.Print "===== Config ====="
    Debug.Print "Template Path: " & gTemplatePath
    Debug.Print "Output Folder: " & gOutputFolderPath

    DebugPrintStringArray "Spec Folder Paths", gSpecFolderPaths
    DebugPrintStringArray "Body Names", gBodyNames
    DebugPrintStringArray "Result Names", gResultNames
    DebugPrintStringArray "Result Markers", gResultMarkers
    DebugPrintStringArray "Output Sheet Names", gOutputSheetNames
    DebugPrintStringArray "Naming Rule Values", gNamingRuleValues

End Sub

Public Sub DebugPrintStringArray(ByVal title As String, ByRef arr() As String)

    Dim i As Long

    Debug.Print title & ":"

    If Not IsStringArrayAllocated(arr) Then
        Debug.Print "  (empty)"
        Exit Sub
    End If

    For i = LBound(arr) To UBound(arr)
        Debug.Print "  [" & i & "] " & arr(i)
    Next i

End Sub
