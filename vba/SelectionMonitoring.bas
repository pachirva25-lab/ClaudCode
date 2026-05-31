Attribute VB_Name = "SelectionMonitoring"
Option Explicit

' Проект: "Выборка из Мониторинга".
' Совместимость: Excel 2016 / 2021 / 2024, 32-bit и 64-bit.
'
' Назначение:
'   Макрос загружается в файл "Выборка из Мониторинга.xlsm" и формирует лист
'   "Отчет УЗРКС" из файлов мониторинга вида "2021 Мониторинг УЗРКС*.xls*" ...
'   "2026 Мониторинг УЗРКС*.xls*".
'
' Источники:
'   1) Предпочтительно: <папка книги>\Source\Мониторинг УЗРКС\
'   2) Резервно:      <папка книги>\
'
' Пользовательский сценарий:
'   Шаг 1 — ввод годов мониторинга: 2021, 2022, 2023, 2024, 2025, 2026.
'   Шаг 2 — ввод вариантов поиска: 1, 2, 3, 4.
'   Шаг 3 — последовательный ввод значений для выбранных вариантов поиска.
'
' Варианты поиска по ТЗ:
'   1 — по номеру закупки: колонки 3, 4, 12, 32 + годовые SAP-колонки.
'   2 — по наименованию закупки: колонка 11, частичное текстовое совпадение.
'   3 — по номеру ЦЗК: колонка 58, поиск без учета регистра и разделителей.
'   4 — по типу сделки: колонка 58, поиск без учета регистра и разделителей.

Private Const PROJECT_NAME As String = "Выборка из Мониторинга"
Private Const SOURCE_SUBFOLDER As String = "Source\Мониторинг УЗРКС"
Private Const OUTPUT_SHEET As String = "Отчет УЗРКС"
Private Const SOURCE_SHEET As String = "Закупщик"
Private Const LOGS_FOLDER As String = "Logs"
Private Const HEADER_ROW As Long = 1
Private Const FIRST_DATA_ROW As Long = 2
Private Const MAX_YEAR As Long = 2026
Private Const MIN_YEAR As Long = 2021

Public Sub СформироватьВыборкуУЗРКС()
    Dim selectedYears As Variant
    Dim searchTypes As Variant
    Dim criteria(1 To 4) As Variant
    Dim sourceFolder As String
    Dim logFile As String
    Dim outputWs As Worksheet
    Dim yearItem As Variant
    Dim searchType As Variant
    Dim monitoringPath As String
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim outputHeadersReady As Boolean
    Dim outputRow As Long
    Dim totalScanned As Long
    Dim totalCopied As Long
    Dim totalErrors As Long
    Dim copiedKeys As Object

    On Error GoTo Fail

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    EnsureProjectFolders ThisWorkbook.Path
    sourceFolder = ResolveSourceFolder(ThisWorkbook.Path)
    logFile = CreateLogFile(ThisWorkbook.Path)

    selectedYears = PromptYears()
    If IsEmpty(selectedYears) Then GoTo CleanExit

    If Not ValidateSelectedFiles(selectedYears, sourceFolder, logFile) Then GoTo CleanExit

    searchTypes = PromptSearchTypes()
    If IsEmpty(searchTypes) Then GoTo CleanExit

    For Each searchType In searchTypes
        criteria(CLng(searchType)) = PromptCriteriaValues(CLng(searchType))
        If IsEmpty(criteria(CLng(searchType))) Then GoTo CleanExit
    Next searchType

    Set outputWs = EnsureWorksheet(ThisWorkbook, OUTPUT_SHEET)
    PrepareOutputSheet outputWs
    outputRow = FIRST_DATA_ROW
    outputHeadersReady = False
    Set copiedKeys = CreateObject("Scripting.Dictionary")

    For Each yearItem In selectedYears
        monitoringPath = FindMonitoringFile(sourceFolder, CLng(yearItem))

        On Error GoTo SourceOpenError
        Set sourceWb = Workbooks.Open(Filename:=monitoringPath, ReadOnly:=True, UpdateLinks:=False)
        On Error GoTo Fail

        If Not WorksheetExists(sourceWb, SOURCE_SHEET) Then
            LogIssue logFile, "В файле отсутствует лист '" & SOURCE_SHEET & "': " & monitoringPath, True
            totalErrors = totalErrors + 1
            GoTo CloseSource
        End If

        Set sourceWs = sourceWb.Worksheets(SOURCE_SHEET)
        ProcessSourceSheet sourceWs, outputWs, CLng(yearItem), searchTypes, criteria, outputHeadersReady, outputRow, totalScanned, totalCopied, copiedKeys, logFile

CloseSource:
        sourceWb.Close SaveChanges:=False
        Set sourceWb = Nothing
    Next yearItem

    If outputHeadersReady Then outputWs.Columns.AutoFit

    MsgBox "Выборка завершена." & vbCrLf & _
           "Обработано строк: " & totalScanned & vbCrLf & _
           "Перенесено строк: " & totalCopied & vbCrLf & _
           "Ошибок/предупреждений: " & totalErrors & vbCrLf & _
           "Журнал: " & logFile, vbInformation, PROJECT_NAME

CleanExit:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

SourceOpenError:
    LogIssue logFile, "Ошибка открытия файла: " & monitoringPath & ". " & Err.Description, True
    totalErrors = totalErrors + 1
    Err.Clear
    Resume CloseSource

Fail:
    MsgBox "Критическая ошибка: " & Err.Description, vbCritical, PROJECT_NAME
    Resume CleanExit
End Sub

' Обратная совместимость с предыдущей версией имени макроса.
Public Sub СформироватьВыборкуМониторинга()
    СформироватьВыборкуУЗРКС
End Sub

Private Sub ProcessSourceSheet( _
    ByVal sourceWs As Worksheet, _
    ByVal outputWs As Worksheet, _
    ByVal monitoringYear As Long, _
    ByVal searchTypes As Variant, _
    ByRef criteria() As Variant, _
    ByRef outputHeadersReady As Boolean, _
    ByRef outputRow As Long, _
    ByRef totalScanned As Long, _
    ByRef totalCopied As Long, _
    ByVal copiedKeys As Object, _
    ByVal logFile As String _
)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowIndex As Long
    Dim rowKey As String

    lastRow = LastUsedRow(sourceWs)
    lastCol = LastUsedColumn(sourceWs)

    If lastRow < FIRST_DATA_ROW Or lastCol = 0 Then
        LogIssue logFile, "На листе '" & SOURCE_SHEET & "' нет данных: " & sourceWs.Parent.Name, True
        Exit Sub
    End If

    If Not outputHeadersReady Then
        CopyOutputHeaders sourceWs, outputWs, lastCol
        outputHeadersReady = True
    End If

    For rowIndex = FIRST_DATA_ROW To lastRow
        totalScanned = totalScanned + 1
        rowKey = CStr(monitoringYear) & "|" & sourceWs.Parent.FullName & "|" & CStr(rowIndex)

        If Not copiedKeys.Exists(rowKey) Then
            If RowMatchesAnySearch(sourceWs, rowIndex, monitoringYear, searchTypes, criteria) Then
                CopyMatchedRowByHeaders sourceWs, outputWs, rowIndex, outputRow, monitoringYear
                copiedKeys.Add rowKey, True
                outputRow = outputRow + 1
                totalCopied = totalCopied + 1
            End If
        End If
    Next rowIndex
End Sub

Private Function RowMatchesAnySearch( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal monitoringYear As Long, _
    ByVal searchTypes As Variant, _
    ByRef criteria() As Variant _
) As Boolean
    Dim searchType As Variant

    For Each searchType In searchTypes
        Select Case CLng(searchType)
            Case 1
                If MatchPurchaseNumber(ws, rowIndex, monitoringYear, criteria(1)) Then RowMatchesAnySearch = True: Exit Function
            Case 2
                If MatchPurchaseName(ws, rowIndex, criteria(2)) Then RowMatchesAnySearch = True: Exit Function
            Case 3
                If MatchHistoryToken(ws, rowIndex, criteria(3)) Then RowMatchesAnySearch = True: Exit Function
            Case 4
                If MatchHistoryToken(ws, rowIndex, criteria(4)) Then RowMatchesAnySearch = True: Exit Function
        End Select
    Next searchType
End Function

Private Function MatchPurchaseNumber(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal monitoringYear As Long, ByVal values As Variant) As Boolean
    Dim columnsToCheck As Variant
    Dim colItem As Variant
    Dim valueItem As Variant
    Dim cellValue As String
    Dim expectedValue As String

    columnsToCheck = PurchaseNumberColumns(monitoringYear)

    For Each colItem In columnsToCheck
        If CLng(colItem) <= LastUsedColumn(ws) Then
            cellValue = NormalizeCode(CStr(ws.Cells(rowIndex, CLng(colItem)).Value))
            For Each valueItem In values
                expectedValue = NormalizeCode(CStr(valueItem))
                If Len(expectedValue) > 0 And cellValue = expectedValue Then
                    MatchPurchaseNumber = True
                    Exit Function
                End If
            Next valueItem
        End If
    Next colItem
End Function

Private Function MatchPurchaseName(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal values As Variant) As Boolean
    Dim valueItem As Variant
    Dim cellValue As String
    Const PURCHASE_NAME_COL As Long = 11

    If PURCHASE_NAME_COL > LastUsedColumn(ws) Then Exit Function

    cellValue = LCase$(Trim$(CStr(ws.Cells(rowIndex, PURCHASE_NAME_COL).Value)))

    For Each valueItem In values
        If Len(Trim$(CStr(valueItem))) > 0 Then
            If InStr(1, cellValue, LCase$(Trim$(CStr(valueItem))), vbTextCompare) > 0 Then
                MatchPurchaseName = True
                Exit Function
            End If
        End If
    Next valueItem
End Function

Private Function MatchHistoryToken(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal values As Variant) As Boolean
    Dim valueItem As Variant
    Dim cellValue As String
    Dim expectedValue As String
    Const HISTORY_COL As Long = 58

    If HISTORY_COL > LastUsedColumn(ws) Then Exit Function

    cellValue = NormalizeToken(CStr(ws.Cells(rowIndex, HISTORY_COL).Value))

    For Each valueItem In values
        expectedValue = NormalizeToken(CStr(valueItem))
        If Len(expectedValue) > 0 Then
            If InStr(1, cellValue, expectedValue, vbTextCompare) > 0 Then
                MatchHistoryToken = True
                Exit Function
            End If
        End If
    Next valueItem
End Function

Private Function PurchaseNumberColumns(ByVal monitoringYear As Long) As Variant
    Select Case monitoringYear
        Case 2021
            PurchaseNumberColumns = Array(3, 4, 12, 30, 32)
        Case 2022
            PurchaseNumberColumns = Array(3, 4, 12, 31, 32)
        Case 2023
            PurchaseNumberColumns = Array(3, 4, 12, 31, 32)
        Case 2024
            PurchaseNumberColumns = Array(3, 4, 12, 21, 32)
        Case 2025
            PurchaseNumberColumns = Array(3, 4, 12, 32)
        Case 2026
            PurchaseNumberColumns = Array(3, 4, 12, 21, 32)
        Case Else
            PurchaseNumberColumns = Array(3, 4, 12, 32)
    End Select
End Function

Private Sub CopyOutputHeaders(ByVal sourceWs As Worksheet, ByVal outputWs As Worksheet, ByVal lastCol As Long)
    sourceWs.Range(sourceWs.Cells(HEADER_ROW, 1), sourceWs.Cells(HEADER_ROW, lastCol)).Copy Destination:=outputWs.Cells(HEADER_ROW, 1)
    outputWs.Cells(HEADER_ROW, lastCol + 1).Value = "Год мониторинга"
    With outputWs.Range(outputWs.Cells(HEADER_ROW, 1), outputWs.Cells(HEADER_ROW, lastCol + 1))
        .Font.Bold = True
        .Interior.Color = RGB(33, 115, 70)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

Private Sub CopyMatchedRowByHeaders(ByVal sourceWs As Worksheet, ByVal outputWs As Worksheet, ByVal sourceRow As Long, ByVal outputRow As Long, ByVal monitoringYear As Long)
    Dim outputLastCol As Long
    Dim outputCol As Long
    Dim sourceCol As Long
    Dim headerName As String

    outputLastCol = LastUsedColumn(outputWs)

    For outputCol = 1 To outputLastCol - 1
        headerName = Trim$(CStr(outputWs.Cells(HEADER_ROW, outputCol).Value))
        sourceCol = FindHeaderColumn(sourceWs, headerName)
        If sourceCol > 0 Then
            sourceWs.Cells(sourceRow, sourceCol).Copy Destination:=outputWs.Cells(outputRow, outputCol)
        End If
    Next outputCol

    outputWs.Cells(outputRow, outputLastCol).Value = monitoringYear
End Sub

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim colIndex As Long

    If Len(headerName) = 0 Then Exit Function

    lastCol = LastUsedColumn(ws)
    For colIndex = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(HEADER_ROW, colIndex).Value)), headerName, vbTextCompare) = 0 Then
            FindHeaderColumn = colIndex
            Exit Function
        End If
    Next colIndex
End Function

Private Function PromptYears() As Variant
    Dim rawValue As String
    Dim values As Variant
    Dim result As Collection
    Dim item As Variant
    Dim yearNumber As Long

    rawValue = InputBox( _
        "Введите годы мониторинга через запятую, точку с запятой, пробел или перенос строки." & vbCrLf & _
        "Доступные годы: 2021, 2022, 2023, 2024, 2025, 2026", _
        PROJECT_NAME & " — шаг 1", _
        "2026")

    If Len(Trim$(rawValue)) = 0 Then Exit Function

    values = SplitValues(rawValue, True)
    Set result = New Collection

    For Each item In values
        If IsNumeric(item) Then
            yearNumber = CLng(item)
            If yearNumber >= MIN_YEAR And yearNumber <= MAX_YEAR Then AddUnique result, CStr(yearNumber)
        End If
    Next item

    If result.Count = 0 Then
        MsgBox "Не выбран ни один допустимый год.", vbExclamation, PROJECT_NAME
        Exit Function
    End If

    PromptYears = CollectionToArray(result)
End Function

Private Function PromptSearchTypes() As Variant
    Dim rawValue As String
    Dim values As Variant
    Dim result As Collection
    Dim item As Variant
    Dim typeNumber As Long

    rawValue = InputBox( _
        "Введите варианты поиска через запятую, точку с запятой или пробел:" & vbCrLf & _
        "1 — по номеру закупки" & vbCrLf & _
        "2 — по наименованию закупки" & vbCrLf & _
        "3 — по номеру ЦЗК" & vbCrLf & _
        "4 — по типу сделки", _
        PROJECT_NAME & " — шаг 2", _
        "1")

    If Len(Trim$(rawValue)) = 0 Then Exit Function

    values = SplitValues(rawValue, True)
    Set result = New Collection

    For Each item In values
        If IsNumeric(item) Then
            typeNumber = CLng(item)
            If typeNumber >= 1 And typeNumber <= 4 Then AddUnique result, CStr(typeNumber)
        End If
    Next item

    If result.Count = 0 Then
        MsgBox "Не выбран ни один допустимый вариант поиска.", vbExclamation, PROJECT_NAME
        Exit Function
    End If

    PromptSearchTypes = CollectionToArray(result)
End Function

Private Function PromptCriteriaValues(ByVal searchType As Long) As Variant
    Dim promptText As String
    Dim defaultValue As String
    Dim rawValue As String
    Dim splitBySpace As Boolean

    Select Case searchType
        Case 1
            promptText = "Введите номер ЗП, номер лота, номер SAP SRM или № извещения ЭТП. Можно указать несколько значений."
            defaultValue = ""
            splitBySpace = True
        Case 2
            promptText = "Введите наименование ЗП или ключевые слова. Для фраз лучше используйте перенос строки, запятую или точку с запятой как разделитель."
            defaultValue = ""
            splitBySpace = False
        Case 3
            promptText = "Введите номер ЦЗК, например: ЦЗК-316-25. Регистр и разделители не учитываются."
            defaultValue = "ЦЗК-316-25"
            splitBySpace = True
        Case 4
            promptText = "Введите тип сделки, например: 4552."
            defaultValue = "4552"
            splitBySpace = True
    End Select

    rawValue = InputBox(promptText, PROJECT_NAME & " — шаг 3", defaultValue)
    If Len(Trim$(rawValue)) = 0 Then Exit Function

    PromptCriteriaValues = SplitValues(rawValue, splitBySpace)
End Function

Private Function SplitValues(ByVal rawValue As String, ByVal splitBySpace As Boolean) As Variant
    Dim normalized As String
    Dim parts As Variant
    Dim result As Collection
    Dim item As Variant

    normalized = rawValue
    normalized = Replace(normalized, vbCrLf, ";")
    normalized = Replace(normalized, vbCr, ";")
    normalized = Replace(normalized, vbLf, ";")
    normalized = Replace(normalized, vbTab, ";")
    normalized = Replace(normalized, ",", ";")
    If splitBySpace Then normalized = Replace(normalized, " ", ";")

    parts = Split(normalized, ";")
    Set result = New Collection

    For Each item In parts
        If Len(Trim$(CStr(item))) > 0 Then AddUnique result, Trim$(CStr(item))
    Next item

    SplitValues = CollectionToArray(result)
End Function

Private Sub AddUnique(ByVal target As Collection, ByVal value As String)
    Dim item As Variant
    For Each item In target
        If StrComp(CStr(item), value, vbTextCompare) = 0 Then Exit Sub
    Next item
    target.Add value
End Sub

Private Function CollectionToArray(ByVal source As Collection) As Variant
    Dim result() As String
    Dim index As Long

    ReDim result(0 To source.Count - 1)
    For index = 1 To source.Count
        result(index - 1) = CStr(source(index))
    Next index

    CollectionToArray = result
End Function

Private Function ValidateSelectedFiles(ByVal selectedYears As Variant, ByVal sourceFolder As String, ByVal logFile As String) As Boolean
    Dim yearItem As Variant
    Dim missingFiles As String

    ValidateSelectedFiles = True

    For Each yearItem In selectedYears
        If Len(FindMonitoringFile(sourceFolder, CLng(yearItem))) = 0 Then
            missingFiles = missingFiles & vbCrLf & CStr(yearItem) & " Мониторинг УЗРКС*.xls*"
            LogIssue logFile, "Не найден файл мониторинга за " & CStr(yearItem) & " год в папке: " & sourceFolder, False
        End If
    Next yearItem

    If Len(missingFiles) > 0 Then
        ValidateSelectedFiles = False
        MsgBox "Не найдены выбранные файлы мониторинга:" & missingFiles & vbCrLf & vbCrLf & _
               "Поместите файлы в папку:" & vbCrLf & sourceFolder & vbCrLf & vbCrLf & _
               "После устранения несоответствия запустите макрос повторно.", vbExclamation, PROJECT_NAME
    End If
End Function

Private Function FindMonitoringFile(ByVal sourceFolder As String, ByVal monitoringYear As Long) As String
    Dim fileName As String
    Dim pattern As String

    pattern = CStr(monitoringYear) & " Мониторинг УЗРКС*.xls*"
    fileName = Dir(WithTrailingSlash(sourceFolder) & pattern)

    If Len(fileName) > 0 Then FindMonitoringFile = WithTrailingSlash(sourceFolder) & fileName
End Function

Private Function ResolveSourceFolder(ByVal projectPath As String) As String
    Dim preferredPath As String
    preferredPath = WithTrailingSlash(projectPath) & SOURCE_SUBFOLDER

    If FolderHasMonitoringFiles(preferredPath) Then
        ResolveSourceFolder = preferredPath
    ElseIf FolderHasMonitoringFiles(projectPath) Then
        ResolveSourceFolder = projectPath
    ElseIf FolderExists(preferredPath) Then
        ResolveSourceFolder = preferredPath
    Else
        ResolveSourceFolder = projectPath
    End If
End Function

Private Function FolderHasMonitoringFiles(ByVal folderPath As String) As Boolean
    FolderHasMonitoringFiles = Len(Dir(WithTrailingSlash(folderPath) & "20?? Мониторинг УЗРКС*.xls*")) > 0
End Function

Private Sub EnsureProjectFolders(ByVal projectPath As String)
    EnsureFolder WithTrailingSlash(projectPath) & "Source"
    EnsureFolder WithTrailingSlash(projectPath) & SOURCE_SUBFOLDER
    EnsureFolder WithTrailingSlash(projectPath) & "Output"
    EnsureFolder WithTrailingSlash(projectPath) & LOGS_FOLDER
    EnsureFolder WithTrailingSlash(projectPath) & "Skills"
    EnsureFolder WithTrailingSlash(projectPath) & "Documentation"
    EnsureFolder WithTrailingSlash(projectPath) & "Backup"
    EnsureFolder WithTrailingSlash(projectPath) & "Домашнее задание"
End Sub

Private Function CreateLogFile(ByVal projectPath As String) As String
    Dim logFolder As String
    logFolder = WithTrailingSlash(projectPath) & LOGS_FOLDER
    EnsureFolder logFolder
    CreateLogFile = WithTrailingSlash(logFolder) & "Ошибки_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
End Function

Private Sub LogIssue(ByVal logFile As String, ByVal message As String, ByVal notifyUser As Boolean)
    Dim fileNumber As Integer

    fileNumber = FreeFile
    Open logFile For Append As #fileNumber
    Print #fileNumber, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
    Close #fileNumber

    If notifyUser Then MsgBox message, vbExclamation, PROJECT_NAME
End Sub

Private Sub PrepareOutputSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
End Sub

Private Function EnsureWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If WorksheetExists(wb, sheetName) Then
        Set EnsureWorksheet = wb.Worksheets(sheetName)
    Else
        Set EnsureWorksheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheet.Name = sheetName
    End If
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim cell As Range
    Set cell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If cell Is Nothing Then LastUsedRow = 0 Else LastUsedRow = cell.Row
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim cell As Range
    Set cell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If cell Is Nothing Then LastUsedColumn = 0 Else LastUsedColumn = cell.Column
End Function

Private Function NormalizeCode(ByVal value As String) As String
    Dim result As String

    result = UCase$(Trim$(CStr(value)))
    result = Replace(result, Chr(160), "")
    result = Replace(result, " ", "")
    result = Replace(result, vbTab, "")

    If IsNumeric(result) Then result = TrimLeadingZeros(result)

    NormalizeCode = result
End Function

Private Function NormalizeToken(ByVal value As String) As String
    Dim result As String
    Dim symbols As Variant
    Dim symbol As Variant

    result = UCase$(Trim$(CStr(value)))
    symbols = Array(" ", vbTab, vbCr, vbLf, "-", "–", "—", "_", "/", "\", ".", ",", ";", ":", "№", "#", "(", ")")

    For Each symbol In symbols
        result = Replace(result, CStr(symbol), "")
    Next symbol

    NormalizeToken = result
End Function

Private Function TrimLeadingZeros(ByVal value As String) As String
    Do While Len(value) > 1 And Left$(value, 1) = "0"
        value = Mid$(value, 2)
    Loop
    TrimLeadingZeros = value
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    FolderExists = (Len(Dir(folderPath, vbDirectory)) > 0)
End Function

Private Sub EnsureFolder(ByVal folderPath As String)
    If Len(folderPath) = 0 Then Exit Sub
    If Not FolderExists(folderPath) Then MkDir folderPath
End Sub

Private Function WithTrailingSlash(ByVal folderPath As String) As String
    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        WithTrailingSlash = folderPath
    Else
        WithTrailingSlash = folderPath & Application.PathSeparator
    End If
End Function
