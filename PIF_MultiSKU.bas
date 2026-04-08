Attribute VB_Name = "modPIFKosmetikMultiSKU"
Option Explicit

'=========================================================
' PIF Kosmetik Multi-SKU Generator
' Template target : PIF_Template.docx
' Workbook sheets  : SKU_Master, Narrative_Text, Consumer_Info, Formula_Data
'
' Assumptions:
'  - SKU_Master / Narrative_Text / Consumer_Info each have SKU_ID in column A
'  - Header row is row 1
'  - Formula_Data has SKU_ID, SORT_ORDER, INGREDIENT_INCI, FUNCTION, PERCENT, REMARKS
'  - Word template uses placeholders like {{BRAND_NAME}}, {{INTENDED_USE}}, etc.
'  - Word template contains a single formula table with 4 columns
'=========================================================

Private Const WORD_TEMPLATE_FILE As String = "PIF_Template.docx"
Private Const OUTPUT_FOLDER As String = "PIF_Output"

Private Const WD_REPLACE_ALL As Long = 2
Private Const WD_FIND_CONTINUE As Long = 1
Private Const WD_FORMAT_XML_DOCUMENT As Long = 12

Public Sub PIF_Generation()
    GeneratePIFs
End Sub

Private Sub GeneratePIFs()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim basePath As String, templatePath As String, outPath As String
    Dim lastRow As Long, r As Long
    Dim skuId As String
    Dim activeFlag As String
    Dim processedCount As Long

    On Error GoTo FailHandler

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("SKU_Master")

    basePath = wb.Path
    If Len(basePath) = 0 Then
        MsgBox "Simpan workbook terlebih dahulu agar macro bisa menemukan template Word.", vbExclamation
        Exit Sub
    End If

    templatePath = basePath & Application.PathSeparator & WORD_TEMPLATE_FILE
    If Dir(templatePath) = "" Then
        MsgBox "Template Word tidak ditemukan:" & vbCrLf & templatePath, vbCritical
        Exit Sub
    End If

    outPath = basePath & Application.PathSeparator & OUTPUT_FOLDER
    EnsureFolder outPath

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Sheet SKU_Master belum berisi data SKU.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    processedCount = 0

    For r = 2 To lastRow
        skuId = Trim$(CStr(ws.Cells(r, 1).Value))
        activeFlag = UCase$(Trim$(CStr(ws.Cells(r, 2).Value)))

        If Len(skuId) > 0 Then
            If Not IsFalseFlag(activeFlag) Then
                GenerateOnePIF wb, skuId, templatePath, outPath
                processedCount = processedCount + 1
            End If
        End If
    Next r

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox "Selesai. " & processedCount & " file PIF berhasil dibuat di:" & vbCrLf & outPath, vbInformation
    Exit Sub

FailHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    MsgBox "Terjadi error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub GenerateOnePIF(ByVal wb As Workbook, ByVal skuId As String, ByVal templatePath As String, ByVal outPath As String)
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim data As Object
    Dim outFile As String
    Dim productName As String

    On Error GoTo FailHandler

    Set data = BuildDataDictionary(wb, skuId)
    If data Is Nothing Then Exit Sub

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    wdApp.DisplayAlerts = 0

    Set wdDoc = wdApp.Documents.Open( _
        FileName:=templatePath, _
        ReadOnly:=False, _
        AddToRecentFiles:=False, _
        Visible:=False _
    )

    ReplaceAllPlaceholders wdDoc, data
    FillFormulaTable wdDoc, wb.Worksheets("Formula_Data"), skuId
    SetDocumentText wdDoc

    productName = GetDictValue(data, "PRODUCT_NAME")
    If Len(productName) = 0 Then productName = GetDictValue(data, "CONSUMER_PRODUCT_NAME")
    If Len(productName) = 0 Then productName = skuId

    outFile = outPath & Application.PathSeparator & _
              SanitizeFileName(skuId & " - " & productName) & ".docx"

    If Len(Dir(outFile)) > 0 Then Kill outFile

    wdDoc.SaveAs2 FileName:=outFile, FileFormat:=WD_FORMAT_XML_DOCUMENT
    wdDoc.Close SaveChanges:=False
    wdApp.Quit

    Set wdDoc = Nothing
    Set wdApp = Nothing
    Exit Sub

FailHandler:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
    On Error GoTo 0
    MsgBox "Gagal membuat PIF untuk SKU " & skuId & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Function BuildDataDictionary(ByVal wb As Workbook, ByVal skuId As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' vbTextCompare

    LoadSheetFields wb.Worksheets("SKU_Master"), skuId, dict
    LoadOptionalSheetFields wb, "Narrative_Text", skuId, dict
    LoadOptionalSheetFields wb, "Consumer_Info", skuId, dict

    ' Helpful fallbacks for the template
    If Not dict.Exists("SECTION1_PRODUCT_NAME") Then dict("SECTION1_PRODUCT_NAME") = GetDictValue(dict, "PRODUCT_NAME")
    If Not dict.Exists("CONSUMER_PRODUCT_NAME") Then dict("CONSUMER_PRODUCT_NAME") = GetDictValue(dict, "PRODUCT_NAME")
    If Not dict.Exists("CONSUMER_ACTIVE_1") Then dict("CONSUMER_ACTIVE_1") = ""
    If Not dict.Exists("CONSUMER_BULLET_1") Then dict("CONSUMER_BULLET_1") = ""
    If Not dict.Exists("CONSUMER_BULLET_2") Then dict("CONSUMER_BULLET_2") = ""
    If Not dict.Exists("CONSUMER_BULLET_3") Then dict("CONSUMER_BULLET_3") = ""
    If Not dict.Exists("CONSUMER_BULLET_4") Then dict("CONSUMER_BULLET_4") = ""

    Set BuildDataDictionary = dict
End Function

Private Sub LoadOptionalSheetFields(ByVal wb As Workbook, ByVal sheetName As String, ByVal skuId As String, ByVal target As Object)
    If WorksheetExists(wb, sheetName) Then
        LoadSheetFields wb.Worksheets(sheetName), skuId, target
    End If
End Sub

Private Sub LoadSheetFields(ByVal ws As Worksheet, ByVal skuId As String, ByVal target As Object)
    Dim rowNum As Long
    Dim lastCol As Long, c As Long
    Dim header As String
    Dim valueText As String

    rowNum = FindSkuRow(ws, skuId)
    If rowNum = 0 Then Exit Sub

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        header = Trim$(CStr(ws.Cells(1, c).Value))
        If Len(header) > 0 Then
            valueText = CellToText(ws.Cells(rowNum, c).Value)
            If Not target.Exists(header) Then
                target.Add header, valueText
            Else
                target(header) = valueText
            End If
        End If
    Next c
End Sub

Private Function FindSkuRow(ByVal ws As Worksheet, ByVal skuId As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 1).Value)) = skuId Then
            FindSkuRow = r
            Exit Function
        End If
    Next r

    FindSkuRow = 0
End Function

Private Sub ReplaceAllPlaceholders(ByVal wdDoc As Object, ByVal data As Object)
    Dim key As Variant
    Dim token As String
    Dim valueText As String
    Dim story As Object
    Dim rng As Object

    For Each key In data.Keys
        token = "{{" & CStr(key) & "}}"
        valueText = CStr(data(key))

        For Each story In wdDoc.StoryRanges
            Set rng = story
            Do While Not rng Is Nothing
                ReplaceInRange rng, token, valueText
                Set rng = rng.NextStoryRange
            Loop
        Next story
    Next key
End Sub

Private Sub ReplaceInRange(ByVal rng As Object, ByVal findText As String, ByVal replaceText As String)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = WD_FIND_CONTINUE
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .Execute Replace:=WD_REPLACE_ALL
    End With
End Sub

Private Sub FillFormulaTable(ByVal wdDoc As Object, ByVal ws As Worksheet, ByVal skuId As String)
    Dim tbl As Object
    Dim formulaRows As Collection
    Dim item As Variant
    Dim i As Long
    Dim requiredRows As Long

    Set tbl = FindFormulaTable(wdDoc)
    If tbl Is Nothing Then
        MsgBox "Formula table tidak ditemukan pada template Word.", vbExclamation
        Exit Sub
    End If

    Set formulaRows = GetFormulaRows(ws, skuId)

    ' Keep header row; make sure data rows are enough.
    requiredRows = formulaRows.Count + 1
    Do While tbl.Rows.Count < requiredRows
        tbl.Rows.Add
    Loop

    ' Fill or clear each data row. Existing formatting is preserved.
    For i = 1 To tbl.Rows.Count - 1
        If i <= formulaRows.Count Then
            item = formulaRows(i)
            tbl.Cell(i + 1, 1).Range.Text = SafeText(item, 1)
            tbl.Cell(i + 1, 2).Range.Text = SafeText(item, 2)
            tbl.Cell(i + 1, 3).Range.Text = SafeText(item, 3)
            tbl.Cell(i + 1, 4).Range.Text = SafeText(item, 4)
        Else
            ClearWordRow tbl, i + 1
        End If
    Next i

    SetTableText tbl
End Sub

Private Function FindFormulaTable(ByVal wdDoc As Object) As Object
    Dim t As Object
    Dim headerText As String

    ' Prefer a 4-column table whose header looks like the formula table.
    For Each t In wdDoc.Tables
        If t.Columns.Count >= 4 Then
            headerText = LCase$( _
                GetCellText(t, 1, 1) & " " & _
                GetCellText(t, 1, 2) & " " & _
                GetCellText(t, 1, 3) & " " & _
                GetCellText(t, 1, 4) _
            )

            If InStr(headerText, "ingredient") > 0 Or _
               InStr(headerText, "function") > 0 Or _
               InStr(headerText, "remarks") > 0 Or _
               InStr(headerText, "percent") > 0 Or _
               InStr(headerText, "%") > 0 Then
                Set FindFormulaTable = t
                Exit Function
            End If
        End If
    Next t

    ' Fallback: first table with at least 4 columns.
    For Each t In wdDoc.Tables
        If t.Columns.Count >= 4 Then
            Set FindFormulaTable = t
            Exit Function
        End If
    Next t

    Set FindFormulaTable = Nothing
End Function

Private Function GetFormulaRows(ByVal ws As Worksheet, ByVal skuId As String) As Collection
    Dim rawRows As New Collection
    Dim sortedRows As Collection
    Dim lastRow As Long, r As Long
    Dim sortOrder As Double
    Dim inciText As String, functionText As String, percentText As String, remarksText As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 1).Value)) = skuId Then
            sortOrder = Val(CStr(ws.Cells(r, 2).Value))
            inciText = CellToText(ws.Cells(r, 3).Value)
            functionText = CellToText(ws.Cells(r, 4).Value)
            percentText = CellToText(ws.Cells(r, 5).Value)
            remarksText = CellToText(ws.Cells(r, 6).Value)

            rawRows.Add Array(sortOrder, inciText, functionText, percentText, remarksText)
        End If
    Next r

    Set sortedRows = SortFormulaRows(rawRows)
    Set GetFormulaRows = sortedRows
End Function

Private Function SortFormulaRows(ByVal rows As Collection) As Collection
    Dim result As New Collection
    Dim arr() As Variant
    Dim i As Long, j As Long
    Dim tmp As Variant
    Dim n As Long

    n = rows.Count
    If n = 0 Then
        Set SortFormulaRows = result
        Exit Function
    End If

    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = rows(i)
    Next i

    For i = 1 To n - 1
        For j = i + 1 To n
            If Val(CStr(arr(j)(0))) < Val(CStr(arr(i)(0))) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    For i = 1 To n
        result.Add arr(i)
    Next i

    Set SortFormulaRows = result
End Function

Private Sub ClearWordRow(ByVal tbl As Object, ByVal rowIndex As Long)
    Dim c As Long
    For c = 1 To tbl.Columns.Count
        On Error Resume Next
        tbl.Cell(rowIndex, c).Range.Text = ""
        On Error GoTo 0
    Next c
End Sub

Private Sub SetDocumentText(ByVal wdDoc As Object)
    Dim story As Object
    Dim rng As Object

    On Error Resume Next

    For Each story In wdDoc.StoryRanges
        Set rng = story
        Do While Not rng Is Nothing
            With rng.Font
                .Color = 0
            End With
            Set rng = rng.NextStoryRange
        Loop
    Next story

    On Error GoTo 0
End Sub

Private Sub SetTableText(ByVal tbl As Object)
    Dim r As Long, c As Long
    Dim cell As Object
    Dim rng As Object

    On Error Resume Next

    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            Set cell = tbl.Cell(r, c)
            Set rng = cell.Range

            ' Remove the end-of-cell marker from font formatting range.
            rng.End = rng.End - 1

            With rng.Font
                .Color = 0
            End With
        Next c
    Next r

    On Error GoTo 0
End Sub

Private Function GetCellText(ByVal tbl As Object, ByVal r As Long, ByVal c As Long) As String
    Dim s As String
    On Error GoTo FailHandler

    s = tbl.Cell(r, c).Range.Text
    If Len(s) >= 2 Then s = Left$(s, Len(s) - 2)
    GetCellText = Trim$(s)
    Exit Function

FailHandler:
    GetCellText = ""
End Function

Private Function SafeText(ByVal arr As Variant, ByVal idx As Long) As String
    On Error GoTo FailHandler
    SafeText = CStr(arr(idx))
    Exit Function

FailHandler:
    SafeText = ""
End Function

Private Function GetDictValue(ByVal d As Object, ByVal key As String) As String
    If Not d Is Nothing Then
        If d.Exists(key) Then
            GetDictValue = CStr(d(key))
            Exit Function
        End If
    End If
    GetDictValue = ""
End Function

Private Function CellToText(ByVal v As Variant) As String
    If IsError(v) Then
        CellToText = ""
    ElseIf IsNull(v) Then
        CellToText = ""
    Else
        CellToText = Trim$(CStr(v))
    End If
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Private Function IsFalseFlag(ByVal flagText As String) As Boolean
    Select Case UCase$(Trim$(flagText))
        Case "N", "NO", "0", "FALSE", "OFF"
            IsFalseFlag = True
        Case Else
            IsFalseFlag = False
    End Select
End Function

Private Sub EnsureFolder(ByVal folderPath As String)
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
End Sub

Private Function SanitizeFileName(ByVal fileName As String) As String
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In badChars
        fileName = Replace(fileName, CStr(ch), "_")
    Next ch
    SanitizeFileName = Trim$(fileName)
End Function
