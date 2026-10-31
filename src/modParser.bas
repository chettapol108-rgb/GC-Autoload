Option Explicit
'------------------------------------------------------------------------------
' TranslateRawInput_ToNormalizedRows
'
' Converts AutoLoad input into normalized production records.
'
' New processing model:
'
' 1. Column A identifies the SKU of each input row.
' 2. Rows are grouped by Order ID + SKU.
' 3. Qty values from all matching rows are summed.
' 4. Raw Input is parsed once per Order + SKU group.
' 5. Parsed quantity must equal the summed order quantity.
' 6. A group is written only when validation succeeds.
'
' Platform filters source rows before grouping; parsing grammar is unchanged.
'------------------------------------------------------------------------------
Public Sub TranslateRawInput_ToNormalizedRows(ByVal wb As Workbook)

    Dim cfg As AutoLoadConfig
    Dim wsInput As Worksheet
    Dim wsNorm As Worksheet

    Dim lastRow As Long
    Dim outRow As Long

    Dim groups As Object
    Dim groupKey As Variant
    Dim group As Object

    Dim tasks As Collection
    Dim task As Object

    Dim expectedQty As Long
    Dim parsedQty As Long

    Dim rawText As String
    Dim barcode As String

    Dim errors As Collection
    
    Dim declaredSkuQty As Long
    Dim skus As Object, sku As Variant

    On Error GoTo Failed
    Set wsNorm = wb.Worksheets("NormalizedInput")
    SetBatchProperty wsNorm, "GC_BatchReady", "NO"
    SetBatchProperty wsNorm, "GC_BatchLoaded", "NO"
    cfg = ReadConfig(wb)
    ValidateBatchInput wb, cfg

    Set wsInput = wb.Worksheets(cfg.InputSheetName)
    Set wsNorm = EnsureNormalizedSheet(wb, cfg)

    If wsNorm Is Nothing Then Exit Sub

    lastRow = FindLastInputRow(wsInput, cfg)

    If lastRow < cfg.FirstDataRow Then

        MsgBox _
            "No input data found.", _
            vbExclamation, _
            "GC AutoLoad"

        Exit Sub

    End If

    Set groups = BuildOrderSkuGroups(wsInput, cfg, lastRow)
    Set errors = New Collection
    Set skus = CreateObject("Scripting.Dictionary")
    skus.CompareMode = vbTextCompare
    For Each groupKey In groups.Keys
        Set group = groups(groupKey)
        skus(CStr(group("SKU"))) = True
    Next groupKey
    ' Resolve every SKU before replacing the previous normalized output.
    For Each sku In skus.Keys
        ReadSkuConfig wb, cfg, CStr(sku)
    Next sku

    '----------------------------------------------------------------------
    ' No rows match the selected platform.
    '----------------------------------------------------------------------
    If groups.Count = 0 Then

        MsgBox _
            "No input rows were found for platform " & cfg.ShopCode & ".", _
            vbCritical, _
            "GC AutoLoad - Platform filter"

        Exit Sub

    End If

    wsNorm.rows("2:" & wsNorm.rows.Count).ClearContents
    ClearAllInputErrors wsInput, cfg, lastRow
    outRow = 2

    '----------------------------------------------------------------------
    ' Process each Order + SKU group.
    '----------------------------------------------------------------------
    For Each groupKey In groups.Keys

        Set group = groups(groupKey)
        ReadSkuConfig wb, cfg, CStr(group("SKU"))

        barcode = CStr(group("Barcode"))
        rawText = CStr(group("RawText"))
        expectedQty = CLng(group("ExpectedQty"))

        '------------------------------------------------------------------
        ' Validate quantity declared in Raw Input against Qty column.
        '
        ' Example:
        '   Qty column total = 2
        '   Raw Input        = N006-E *3pcs || Name
        '
        ' Result:
        '   Error - quantities do not match.
        '------------------------------------------------------------------
        declaredSkuQty = GetDeclaredSkuQty(rawText, cfg.sku)
        
        If declaredSkuQty > 0 Then
        
            If declaredSkuQty <> expectedQty Then
        
                AddTranslationError _
                    errors, _
                    barcode, _
                    cfg.sku, _
                    "Quantity does not match." & vbCrLf & _
                    "Order quantity: " & expectedQty & " pcs" & vbCrLf & _
                    "Raw Input quantity: " & declaredSkuQty & " pcs"
        
                MarkGroupError _
                    wsInput, _
                    cfg, _
                    group("SourceRows")
        
                GoTo NextGroup
        
            End If
        
        End If

        '------------------------------------------------------------------
        ' Same Order + SKU contains different Raw Input values.
        '------------------------------------------------------------------
        If CBool(group("HasRawTextConflict")) Then

            AddTranslationError _
                errors, _
                barcode, _
                cfg.sku, _
                "The same order contains different Raw Input values."

            MarkGroupError wsInput, cfg, group("SourceRows")

            GoTo NextGroup

        End If

        '------------------------------------------------------------------
        ' Raw Input missing.
        '------------------------------------------------------------------
        If Trim$(rawText) = "" Then

            AddTranslationError _
                errors, _
                barcode, _
                cfg.sku, _
                "Raw Input is empty."

            MarkGroupError wsInput, cfg, group("SourceRows")

            GoTo NextGroup

        End If

        Set tasks = ParseRawInputToTasks(rawText, cfg.sku)

        '------------------------------------------------------------------
        ' SKU could not be found / parsed from Raw Input.
        '------------------------------------------------------------------
        If tasks.Count = 0 Then

            AddTranslationError _
                errors, _
                barcode, _
                cfg.sku, _
                "No name data for this SKU could be found in Raw Input."

            MarkGroupError wsInput, cfg, group("SourceRows")

            GoTo NextGroup

        End If

        '------------------------------------------------------------------
        ' Resolve quantities.
        '
        ' Example:
        '   One name, ExpectedQty 3
        '   -> Name Qty 3
        '
        '   A \ B \ C, ExpectedQty 3
        '   -> A=1, B=1, C=1
        '------------------------------------------------------------------
        ResolveSingleRowTaskQuantities tasks, expectedQty

        parsedQty = GetTotalTaskQty(tasks)

        '------------------------------------------------------------------
        ' Quantity validation.
        '------------------------------------------------------------------
        If parsedQty <> expectedQty Then

            AddTranslationError _
                errors, _
                barcode, _
                cfg.sku, _
                "Order quantity = " & expectedQty & _
                " pcs, but parsed name quantity = " & _
                parsedQty & " pcs."

            MarkGroupError wsInput, cfg, group("SourceRows")

            GoTo NextGroup

        End If

        '------------------------------------------------------------------
        ' Validation passed.
        ' Write normalized tasks.
        '------------------------------------------------------------------
        For Each task In tasks

            WriteNormalizedTask _
                wsNorm, _
                outRow, _
                task, _
                barcode

            outRow = outRow + 1

        Next task

NextGroup:
    Next groupKey

    '----------------------------------------------------------------------
    ' Final result.
    '----------------------------------------------------------------------
    If errors.Count > 0 Then

        MsgBox _
            BuildTranslationErrorMessage(errors), _
            vbExclamation, _
            "GC AutoLoad - Input Error"

    Else

        Set skus = NormalizedSkus(wsNorm)
        SetBatchProperty wsNorm, "GC_BatchConfig", BatchConfigKey(wb, cfg, skus)
        SetBatchProperty wsNorm, "GC_BatchReady", "YES"

        MsgBox _
            "Translate completed successfully." & vbCrLf & vbCrLf & _
            "Platform: " & cfg.ShopCode & "; SKUs: " & skus.Count & vbCrLf & _
            "Orders processed: " & groups.Count & vbCrLf & _
            "Rows created in NormalizedInput: " & (outRow - 2), _
            vbInformation, _
            "GC AutoLoad"

    End If
    Exit Sub
Failed:
    MsgBox "Translate stopped: " & Err.Description, vbExclamation, "GC AutoLoad"
End Sub

'------------------------------------------------------------------------------
' WriteNormalizedTask
'
' Writes one parsed task to the NormalizedInput sheet.
'
' OID (F) and Barcode (G) use Text format to preserve the original
' order number exactly.
'------------------------------------------------------------------------------
Private Sub WriteNormalizedTask( _
    ByVal wsNorm As Worksheet, _
    ByVal outputRow As Long, _
    ByVal task As Object, _
    ByVal rawBarcode As String _
)

    wsNorm.Cells(outputRow, "A").value = task("SKU")
    wsNorm.Cells(outputRow, "B").value = task("NameA")
    wsNorm.Cells(outputRow, "C").value = task("NameB")
    wsNorm.Cells(outputRow, "D").value = task("NameC")
    wsNorm.Cells(outputRow, "E").value = task("NameD")

    wsNorm.Cells(outputRow, "F").NumberFormat = "@"
    wsNorm.Cells(outputRow, "F").value = rawBarcode

    wsNorm.Cells(outputRow, "G").NumberFormat = "@"
    wsNorm.Cells(outputRow, "G").value = rawBarcode

    wsNorm.Cells(outputRow, "H").value = task("Qty")
End Sub
'------------------------------------------------------------------------------
' GetTotalTaskQty
'
' Returns the total quantity from all parsed tasks in a collection.
'
' Example:
'   Task 1 Qty = 2
'   Task 2 Qty = 3
'
' Result:
'   5
'------------------------------------------------------------------------------
Private Function GetTotalTaskQty(ByVal tasks As Collection) As Long

    Dim task As Object
    Dim totalQty As Long

    totalQty = 0

    For Each task In tasks
        totalQty = totalQty + CLng(task("Qty"))
    Next task

    GetTotalTaskQty = totalQty
End Function
'------------------------------------------------------------------------------
' FindLastInputRow
'
' Finds the last meaningful input row using SKU, Barcode, Qty and Raw Input.
'------------------------------------------------------------------------------
Private Function FindLastInputRow(ByVal ws As Worksheet, ByRef cfg As AutoLoadConfig) As Long

    Dim rSku As Long
    Dim rBarcode As Long
    Dim rQty As Long
    Dim rRaw As Long

    rSku = ws.Cells(ws.rows.Count, cfg.ColSku).End(xlUp).Row
    rBarcode = ws.Cells(ws.rows.Count, cfg.ColBarcode).End(xlUp).Row
    rQty = ws.Cells(ws.rows.Count, cfg.ColQty).End(xlUp).Row
    rRaw = ws.Cells(ws.rows.Count, cfg.ColRawInput).End(xlUp).Row

    FindLastInputRow = Application.WorksheetFunction.Max(rSku, rBarcode, rQty, rRaw)

End Function

'------------------------------------------------------------------------------
' BuildOrderSkuGroups
'
' Groups all input rows by Order ID + SKU.
'
' For all SKUs in the selected platform:
' - All rows belonging to the same Order ID are combined.
' - Qty values are summed into ExpectedQty.
' - Raw Input is stored once because duplicated order rows normally contain
'   the same Raw Input.
'
' Each group contains:
'   SKU
'   Barcode
'   ExpectedQty
'   RawText
'   SourceRows
'   HasRawTextConflict
'------------------------------------------------------------------------------
Private Function BuildOrderSkuGroups(ByVal wsInput As Worksheet, ByRef cfg As AutoLoadConfig, ByVal lastRow As Long) As Object

    Dim groups As Object
    Dim group As Object
    Dim rows As Collection

    Dim r As Long

    Dim rowSku As String
    Dim barcode As String
    Dim rawText As String
    Dim qty As Long
    Dim groupKey As String

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare

    For r = cfg.FirstDataRow To lastRow

        ' Excluded platforms must not affect validation or grouping.
        If MatchesPlatform(wsInput.Cells(r, cfg.ColPlatform).Value2, cfg.ShopCode) Then
        rowSku = Trim$(CStr(wsInput.Cells(r, cfg.ColSku).value))

        barcode = Trim$(CStr(wsInput.Cells(r, cfg.ColBarcode).Text))

        rawText = CStr(wsInput.Cells(r, cfg.ColRawInput).value)

        qty = CLng(Val(wsInput.Cells(r, cfg.ColQty).value))

            If rowSku = "" And (barcode <> "" Or rawText <> "") Then
                Err.Raise vbObjectError + 2111, , "Missing SKU on input row " & r
            End If
            If rowSku <> "" And barcode = "" Then
                Err.Raise vbObjectError + 2112, , "Missing Order ID on input row " & r
            End If

            If barcode <> "" Then

                If qty <= 0 Then qty = 1

                groupKey = UCase$(barcode) & "|||" & UCase$(rowSku)

                If Not groups.Exists(groupKey) Then

                    Set group = CreateObject("Scripting.Dictionary")
                    Set rows = New Collection

                    group("SKU") = rowSku
                    group("Barcode") = barcode
                    group("ExpectedQty") = 0
                    group("RawText") = ""
                    group("HasRawTextConflict") = False

                    Set group("SourceRows") = rows

                    groups.Add groupKey, group

                Else

                    Set group = groups(groupKey)

                End If

                ' Sum quantity from every row in this Order + SKU group.
                group("ExpectedQty") = CLng(group("ExpectedQty")) + qty

                group("SourceRows").Add r

                ' Store the first available Raw Input.
                If Trim$(rawText) <> "" Then

                    If Trim$(CStr(group("RawText"))) = "" Then

                        group("RawText") = rawText

                    ElseIf StrComp(NormalizeRawText(CStr(group("RawText"))), NormalizeRawText(rawText), vbBinaryCompare) <> 0 Then

                        ' The same Order + SKU contains different Raw Input.
                        group("HasRawTextConflict") = True

                    End If
                End If

            End If
        End If

    Next r

    Set BuildOrderSkuGroups = groups

End Function

'------------------------------------------------------------------------------
' ParseRawInputToTasks
'
' Converts raw input text into normalized task records for the selected SKU.
'
' Current input grammar:
'
'   ";"  separates multiple SKUs in the same header.
'   "\"  separates names/items belonging to the same SKU.
'   "\\" ends the current SKU block.
'   "||" separates the SKU/header from its data.
'   "*Np" indicates the quantity of an item.
'
' Examples:
'
'   L007-C || FARATOEY \ PEARKI \\
'
'   L007-J ; L007-N ; L007-N08 ; N006-E || XXX \\
'
'   N003-E *4p || YYY *2p \ ZZZ *2p \\
'------------------------------------------------------------------------------
Private Function ParseRawInputToTasks(ByVal rawText As String, ByVal selectedSku As String) As Collection
    Dim result As Collection
    Dim blocks As Variant
    Dim i As Long
    Dim blockText As String
    Dim blockTasks As Collection
    Dim task As Object
    
    Set result = New Collection
    
    rawText = NormalizeRawText(rawText)
    rawText = NormalizeMarkerSpaces(rawText)
    
    blocks = Split(rawText, "\\")
    
    For i = LBound(blocks) To UBound(blocks)
        blockText = Trim$(CStr(blocks(i)))
        
        If blockText <> "" Then
            Set blockTasks = ParseOneBlock(blockText, selectedSku)
            
            For Each task In blockTasks
                result.Add task
            Next task
        End If
    Next i
    
    Set ParseRawInputToTasks = result
End Function
'------------------------------------------------------------------------------
' NormalizeRawText
'
' Normalizes copied customer input before parsing.
'
' Only name-line markers are converted:
'   Thai line number 1 -> #1
'   Thai line number 2 -> #2
'   Thai line number 3 -> #3
'   Thai line number 4 -> #4
'
' Quantity units are intentionally NOT replaced.
' The quantity parser reads the number after "*" directly.
'
' Examples:
'   *3pcs  -> Qty 3
'   *3p    -> Qty 3
'
' This prevents normal names such as Thai names containing quantity words
' from being modified accidentally.
'------------------------------------------------------------------------------
Private Function NormalizeRawText(ByVal s As String) As String

    Dim lineWord As String
    Dim i As Long

    lineWord = ThaiWord_LineNumber()

    ' Normalize line breaks and tabs.
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")

    ' Normalize common Unicode spaces.
    s = Replace(s, ChrW$(160), " ")

    ' Convert Thai numbered-line prefixes to # markers.
    For i = 1 To 4

        s = Replace(s, lineWord & CStr(i), "#" & CStr(i), 1, -1, vbTextCompare)

        s = Replace(s, lineWord & " " & CStr(i), "#" & CStr(i), 1, -1, vbTextCompare)

    Next i

    ' Collapse repeated spaces.
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeRawText = Trim$(s)

End Function
'------------------------------------------------------------------------------
' NormalizeMarkerSpaces
'
' Standardizes numbered name markers.
'
' Examples:
'   "#1 ||" -> "#1||"
'   "#2  ||" -> "#2||"
'------------------------------------------------------------------------------
Private Function NormalizeMarkerSpaces(ByVal s As String) As String
    Dim i As Long

    For i = 1 To 4
        s = Replace(s, "#" & CStr(i) & "  ||", "#" & CStr(i) & "||")
        s = Replace(s, "#" & CStr(i) & " ||", "#" & CStr(i) & "||")
        s = Replace(s, "#" & CStr(i) & "|| ", "#" & CStr(i) & "||")
    Next i

    NormalizeMarkerSpaces = s
End Function

'------------------------------------------------------------------------------
' ParseOneBlock
'
' Parses one block of raw input.
'
' If the block has no SKU header, it assumes the block belongs to selectedSku.
' This is required for Condition 1 / 2 / 5 where customer input may contain
' only name fields and no SKU text.
'------------------------------------------------------------------------------
Private Function ParseOneBlock(ByVal blockText As String, ByVal selectedSku As String) As Collection
    Dim result As Collection
    Dim headerText As String
    Dim contentText As String
    Dim items As Variant
    Dim i As Long
    Dim itemText As String
    Dim task As Object
    
    Set result = New Collection
    
    ExtractHeaderAndContent blockText, selectedSku, headerText, contentText
    
    If headerText <> "" Then
        If Not HeaderContainsSku(headerText, selectedSku) Then
            Set ParseOneBlock = result
            Exit Function
        End If
    End If
    
    If contentText = "" Then contentText = blockText
    
    items = Split(contentText, "\")
    
    For i = LBound(items) To UBound(items)
        itemText = Trim$(CStr(items(i)))
        
        If itemText <> "" Then
            Set task = ParseOneItem(itemText, selectedSku)
            result.Add task
        End If
    Next i
    
    Set ParseOneBlock = result
End Function

'------------------------------------------------------------------------------
' ExtractHeaderAndContent
'
' Splits a block into header and content.
'
' Header exists only when the text before the first delimiter looks like SKU.
'
' Examples:
'   "L007-A || Pearl"          -> header = L007-A, content = Pearl
'   "L007-V1 1||A 2||B 3||C"  -> header = L007-V1, content = 1||A 2||B 3||C
'   "1||A 2||B 3||C"          -> header = "", content = full text
'------------------------------------------------------------------------------
Private Sub ExtractHeaderAndContent(ByVal blockText As String, ByVal selectedSku As String, ByRef headerText As String, ByRef contentText As String)
    Dim pDelim As Long
    Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long
    Dim pEnd As Long
    Dim candidateHeader As String
    
    pDelim = InStr(1, blockText, "||", vbTextCompare)
    p1 = InStr(1, blockText, "#1||", vbTextCompare)
    p2 = InStr(1, blockText, "#2||", vbTextCompare)
    p3 = InStr(1, blockText, "#3||", vbTextCompare)
    p4 = InStr(1, blockText, "#4||", vbTextCompare)
    
    pEnd = FirstPositiveMin(pDelim, p1, p2, p3, p4)
    
    If pEnd = 0 Then
        headerText = ""
        contentText = blockText
        Exit Sub
    End If
    
    candidateHeader = Trim$(Left$(blockText, pEnd - 1))
    
    If LooksLikeSkuHeader(candidateHeader, selectedSku) Then
        headerText = candidateHeader
        
        If Mid$(blockText, pEnd, 2) = "||" Then
            contentText = Trim$(Mid$(blockText, pEnd + 2))
        Else
            contentText = Trim$(Mid$(blockText, pEnd))
        End If
    Else
        headerText = ""
        contentText = blockText
    End If
End Sub

'------------------------------------------------------------------------------
' LooksLikeSkuHeader
'
' Determines whether the text before the data marker looks like an SKU header.
'
' The detection is intentionally generic and does not depend on specific
' product families such as L007, N015, M011, PRE-FLX, etc.
'
' Examples recognized as SKU headers:
'   L007-J
'   N015-01
'   M011-B
'   PRE-FLX-0011
'   L007-J ; L007-N02
'
' A valid SKU header must contain at least one hyphen and must not contain
' any name marker such as #1||, #2||, etc.
'------------------------------------------------------------------------------
Private Function LooksLikeSkuHeader( _
    ByVal candidateHeader As String, _
    ByVal selectedSku As String _
) As Boolean

    Dim cleanHeader As String
    Dim skuParts As Variant
    Dim i As Long
    Dim part As String

    cleanHeader = Trim$(candidateHeader)

    If cleanHeader = "" Then
        LooksLikeSkuHeader = False
        Exit Function
    End If

    ' Name markers indicate content, not SKU header.
    If InStr(1, cleanHeader, "#1||", vbTextCompare) > 0 Or _
       InStr(1, cleanHeader, "#2||", vbTextCompare) > 0 Or _
       InStr(1, cleanHeader, "#3||", vbTextCompare) > 0 Or _
       InStr(1, cleanHeader, "#4||", vbTextCompare) > 0 Then

        LooksLikeSkuHeader = False
        Exit Function
    End If

    ' Multiple SKUs may be separated by semicolons.
    skuParts = Split(cleanHeader, ";")

    For i = LBound(skuParts) To UBound(skuParts)

        part = Trim$(CStr(skuParts(i)))

        ' Remove optional quantity such as "*2p".
        part = RemoveTrailingQty(part)

        ' SKU format used in this system contains a hyphen.
        If part <> "" And InStr(1, part, "-", vbBinaryCompare) > 0 Then
            LooksLikeSkuHeader = True
            Exit Function
        End If

    Next i

    LooksLikeSkuHeader = False

End Function
'------------------------------------------------------------------------------
' HeaderContainsSku
'
' Checks whether the selected SKU exists in a header containing one or more
' SKU entries separated by semicolons.
'
' Each SKU entry may have its own quantity.
'
' Example:
'   L007-J *2p ; L007-C
'
' This function removes quantity from each SKU entry separately, instead of
' removing quantity from the entire header.
'------------------------------------------------------------------------------
Private Function HeaderContainsSku( _
    ByVal headerText As String, _
    ByVal selectedSku As String _
) As Boolean

    Dim skuParts As Variant
    Dim i As Long
    Dim currentSku As String

    ' Split the header first because each SKU may contain its own quantity.
    skuParts = Split(headerText, ";")

    For i = LBound(skuParts) To UBound(skuParts)

        ' Remove quantity only from the current SKU entry.
        currentSku = RemoveTrailingQty(Trim$(CStr(skuParts(i))))

        If StrComp( _
            Trim$(currentSku), _
            Trim$(selectedSku), _
            vbTextCompare _
        ) = 0 Then

            HeaderContainsSku = True
            Exit Function
        End If
    Next i

    HeaderContainsSku = False
End Function

'------------------------------------------------------------------------------
' ParseOneItem
'
' Parses one item into one task dictionary.
'
' Output fields:
'   SKU
'   NameA
'   NameB
'   NameC
'   NameD
'   Qty
'   ExplicitQty
'
' ExplicitQty is True only when the item contains a valid "*Np" quantity.
'------------------------------------------------------------------------------
Private Function ParseOneItem( _
    ByVal itemText As String, _
    ByVal selectedSku As String _
) As Object

    Dim task As Object
    Dim cleanItem As String

    Set task = CreateObject("Scripting.Dictionary")

    task("SKU") = selectedSku
    task("ExplicitQty") = HasExplicitTrailingQty(itemText)
    task("Qty") = ExtractTrailingQty(itemText)

    cleanItem = RemoveTrailingQty(itemText)

    If HasNameMarkers(cleanItem) Then

        task("NameA") = ExtractValueByMarker(cleanItem, "1")
        task("NameB") = ExtractValueByMarker(cleanItem, "2")
        task("NameC") = ExtractValueByMarker(cleanItem, "3")
        task("NameD") = ExtractValueByMarker(cleanItem, "4")

    Else

        task("NameA") = Trim$(cleanItem)
        task("NameB") = ""
        task("NameC") = ""
        task("NameD") = ""

    End If

    Set ParseOneItem = task
End Function

'------------------------------------------------------------------------------
' BuildBarcodeRowCounts
'
' Counts only input rows whose raw text contains tasks for the selected SKU.
' This keeps barcode occurrence counts aligned with rows that will actually
' be processed by the translator.
'------------------------------------------------------------------------------
Private Function BuildBarcodeRowCounts( _
    ByVal wsInput As Worksheet, _
    ByRef cfg As AutoLoadConfig, _
    ByVal lastRow As Long _
) As Object

    Dim result As Object
    Dim tasks As Collection
    Dim r As Long
    Dim barcodeKey As String
    Dim rawText As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    For r = cfg.FirstDataRow To lastRow

        barcodeKey = Trim$(CStr( _
            wsInput.Cells(r, cfg.ColBarcode).Text _
        ))

        rawText = CStr(wsInput.Cells(r, cfg.ColRawInput).value)

        If barcodeKey <> "" And Trim$(rawText) <> "" Then

            Set tasks = ParseRawInputToTasks(rawText, cfg.sku)

            If tasks.Count > 0 Then
                If result.Exists(barcodeKey) Then
                    result(barcodeKey) = CLng(result(barcodeKey)) + 1
                Else
                    result.Add barcodeKey, 1
                End If
            End If
        End If
    Next r

    Set BuildBarcodeRowCounts = result
End Function


'------------------------------------------------------------------------------
' ResolveSingleRowTaskQuantities
'
' Resolves task quantities when a barcode appears on only one input row.
'
' Rules:
'
' 1. One task without "*Np":
'    Use Qty from the input column.
'
' 2. One task with "*Np":
'    Keep the explicit quantity.
'
' 3. Multiple tasks:
'    Keep each task quantity.
'    Tasks without "*Np" remain Qty 1.
'
' Examples:
'
'   A, Input Qty 3
'   -> A Qty 3
'
'   A \ B, Input Qty 2
'   -> A Qty 1, B Qty 1
'
'   A *2p \ B *3p, Input Qty 5
'   -> A Qty 2, B Qty 3
'------------------------------------------------------------------------------
Private Sub ResolveSingleRowTaskQuantities( _
    ByVal tasks As Collection, _
    ByVal rawQty As Long _
)

    Dim task As Object

    If tasks.Count = 1 Then

        Set task = tasks(1)

        If Not CBool(task("ExplicitQty")) Then
            task("Qty") = rawQty
        End If

    Else

        For Each task In tasks

            If Not CBool(task("ExplicitQty")) Then
                task("Qty") = 1
            End If
        Next task

    End If
End Sub

'------------------------------------------------------------------------------
' HasNameMarkers
'
' Detects whether the item uses numbered name fields.
'------------------------------------------------------------------------------
Private Function HasNameMarkers(ByVal s As String) As Boolean
    HasNameMarkers = _
        InStr(1, s, "#1||", vbTextCompare) > 0 Or _
        InStr(1, s, "#2||", vbTextCompare) > 0 Or _
        InStr(1, s, "#3||", vbTextCompare) > 0 Or _
        InStr(1, s, "#4||", vbTextCompare) > 0
End Function

'------------------------------------------------------------------------------
' ExtractValueByMarker
'
' Extracts value after marker #1||, #2||, #3||, or #4||.
' Extraction stops at the next marker.
'------------------------------------------------------------------------------
Private Function ExtractValueByMarker( _
    ByVal itemText As String, _
    ByVal markerNo As String _
) As String

    Dim markerText As String
    Dim pStart As Long
    Dim pValue As Long
    Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long
    Dim pEnd As Long

    markerText = "#" & markerNo & "||"

    pStart = InStr(1, itemText, markerText, vbTextCompare)

    If pStart = 0 Then
        ExtractValueByMarker = ""
        Exit Function
    End If

    pValue = pStart + Len(markerText)

    p1 = InStr(pValue, itemText, "#1||", vbTextCompare)
    p2 = InStr(pValue, itemText, "#2||", vbTextCompare)
    p3 = InStr(pValue, itemText, "#3||", vbTextCompare)
    p4 = InStr(pValue, itemText, "#4||", vbTextCompare)

    pEnd = FirstPositiveMin(p1, p2, p3, p4)

    If pEnd = 0 Then
        ExtractValueByMarker = Trim$(Mid$(itemText, pValue))
    Else
        ExtractValueByMarker = Trim$(Mid$(itemText, pValue, pEnd - pValue))
    End If
End Function

'------------------------------------------------------------------------------
' ExtractTrailingQty
'
' Extracts trailing quantity from text.
'
' Examples:
'   "Pearl *3p"    -> 3
'   "Pearl *3pcs"  -> 3
'
' If no quantity is found, returns 1.
'------------------------------------------------------------------------------
Private Function ExtractTrailingQty(ByVal s As String) As Long
    Dim pStar As Long
    Dim qtyText As String
    Dim i As Long
    Dim ch As String
    
    ExtractTrailingQty = 1
    
    pStar = InStrRev(s, "*")
    If pStar = 0 Then Exit Function
    
    For i = pStar + 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            qtyText = qtyText & ch
        ElseIf qtyText <> "" Then
            Exit For
        End If
    Next i
    
    If qtyText <> "" Then
        ExtractTrailingQty = CLng(qtyText)
    End If
End Function

'------------------------------------------------------------------------------
' RemoveTrailingQty
'
' Removes trailing quantity such as "*1p", "*2p", or "*3pcs".
'------------------------------------------------------------------------------
Private Function RemoveTrailingQty(ByVal s As String) As String
    Dim pStar As Long
    
    pStar = InStrRev(s, "*")
    
    If pStar > 0 Then
        RemoveTrailingQty = Trim$(Left$(s, pStar - 1))
    Else
        RemoveTrailingQty = Trim$(s)
    End If
End Function

'------------------------------------------------------------------------------
' FirstPositiveMin
'
' Returns the smallest positive value from the given list.
' Returns 0 if none exists.
'------------------------------------------------------------------------------
Private Function FirstPositiveMin(ParamArray nums() As Variant) As Long
    Dim i As Long
    Dim n As Long
    Dim minVal As Long
    
    minVal = 0
    
    For i = LBound(nums) To UBound(nums)
        n = CLng(nums(i))
        If n > 0 Then
            If minVal = 0 Or n < minVal Then minVal = n
        End If
    Next i
    
    FirstPositiveMin = minVal
End Function

'------------------------------------------------------------------------------
'Thai Language Section
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' BuildUnicodeText
'
' Builds Unicode text from a list of character codes.
' This avoids typing Thai text directly in the VBA editor.
'------------------------------------------------------------------------------
Private Function BuildUnicodeText(ParamArray charCodes() As Variant) As String
    Dim i As Long
    Dim result As String

    For i = LBound(charCodes) To UBound(charCodes)
        result = result & ChrW$(CLng(charCodes(i)))
    Next i

    BuildUnicodeText = result
End Function

'------------------------------------------------------------------------------
' ThaiWord_LineNumber
'
' Returns the Thai text meaning "line number":
'------------------------------------------------------------------------------
Private Function ThaiWord_LineNumber() As String
    ThaiWord_LineNumber = BuildUnicodeText( _
        &HE1A, _
        &HE23, _
        &HE23, _
        &HE17, _
        &HE31, _
        &HE14, _
        &HE17, _
        &HE35, _
        &HE48 _
    )
End Function

'------------------------------------------------------------------------------
' AddTranslationError
'
' Adds a user-readable translation error to the error collection.
'------------------------------------------------------------------------------
Private Sub AddTranslationError( _
    ByVal errors As Collection, _
    ByVal barcode As String, _
    ByVal sku As String, _
    ByVal message As String _
)

    errors.Add _
        "Order: " & barcode & vbCrLf & _
        "SKU: " & sku & vbCrLf & _
        message

End Sub

'------------------------------------------------------------------------------
' BuildTranslationErrorMessage
'
' Builds a readable summary of all translation errors.
'------------------------------------------------------------------------------
Private Function BuildTranslationErrorMessage( _
    ByVal errors As Collection _
) As String

    Dim result As String
    Dim i As Long

    result = _
        "Some orders could not be translated." & vbCrLf & _
        "Please check the highlighted input rows." & _
        vbCrLf & vbCrLf

    For i = 1 To errors.Count

        result = result & _
            "----------------------------------------" & vbCrLf & _
            CStr(errors(i)) & vbCrLf

    Next i

    BuildTranslationErrorMessage = result

End Function

'------------------------------------------------------------------------------
' MarkGroupError
'
' Highlights all source rows belonging to an invalid Order + SKU group.
'------------------------------------------------------------------------------
Private Sub MarkGroupError( _
    ByVal wsInput As Worksheet, _
    ByRef cfg As AutoLoadConfig, _
    ByVal sourceRows As Collection _
)

    Dim rowIndex As Variant

    For Each rowIndex In sourceRows

        wsInput.Range( _
            cfg.ColSku & CStr(rowIndex) & ":" & _
            cfg.ColRawInput & CStr(rowIndex) _
        ).Interior.Color = RGB(255, 199, 206)

    Next rowIndex

End Sub
'------------------------------------------------------------------------------
' ClearAllInputErrors
'
' Removes previous error highlighting from input rows.
'------------------------------------------------------------------------------
Private Sub ClearAllInputErrors( _
    ByVal wsInput As Worksheet, _
    ByRef cfg As AutoLoadConfig, _
    ByVal lastRow As Long _
)

    If lastRow < cfg.FirstDataRow Then Exit Sub

    wsInput.Range( _
        cfg.ColSku & cfg.FirstDataRow & ":" & _
        cfg.ColRawInput & lastRow _
    ).Interior.Pattern = xlNone

End Sub

Private Function HasExplicitTrailingQty(ByVal s As String) As Boolean

    Dim starPos As Long
    Dim tail As String
    Dim i As Long
    Dim ch As String
    Dim numberPart As String
    Dim unitPart As String

    s = Trim$(s)

    starPos = InStrRev(s, "*")
    If starPos = 0 Then Exit Function

    tail = Trim$(Mid$(s, starPos + 1))
    If tail = "" Then Exit Function

    ' Read quantity number.
    For i = 1 To Len(tail)

        ch = Mid$(tail, i, 1)

        If ch >= "0" And ch <= "9" Then
            numberPart = numberPart & ch
        Else
            Exit For
        End If

    Next i

    If numberPart = "" Then Exit Function

    ' Read quantity unit.
    unitPart = LCase$(Trim$(Mid$(tail, Len(numberPart) + 1)))

    HasExplicitTrailingQty = _
        (unitPart = "p" Or unitPart = "pcs")

End Function

'------------------------------------------------------------------------------
' GetDeclaredSkuQty
'
' Returns the total quantity explicitly declared for the selected SKU
' inside Raw Input.
'
' Examples:
'
'   N006-E *3pcs || Name
'       -> 3
'
'   N006-E *2pcs || A
'   \\
'   N006-E *1pcs || B
'       -> 3
'
'   N006-E || Name
'       -> 0   (no explicit quantity declared)
'
' Only quantities belonging to the selected SKU are counted.
'------------------------------------------------------------------------------
Private Function GetDeclaredSkuQty( _
    ByVal rawText As String, _
    ByVal selectedSku As String _
) As Long

    Dim blocks As Variant
    Dim blockText As String

    Dim headerText As String
    Dim contentText As String

    Dim skuParts As Variant
    Dim skuPart As String
    Dim cleanSku As String

    Dim i As Long
    Dim j As Long
    Dim totalQty As Long

    rawText = NormalizeRawText(rawText)
    rawText = NormalizeMarkerSpaces(rawText)

    ' Separate SKU blocks.
    blocks = Split(rawText, "\\")

    For i = LBound(blocks) To UBound(blocks)

        blockText = Trim$(CStr(blocks(i)))

        If blockText <> "" Then

            ExtractHeaderAndContent _
                blockText, _
                selectedSku, _
                headerText, _
                contentText

            If headerText <> "" Then

                ' Multiple SKUs in one header are separated by ";".
                skuParts = Split(headerText, ";")

                For j = LBound(skuParts) To UBound(skuParts)

                    skuPart = Trim$(CStr(skuParts(j)))
                    cleanSku = RemoveTrailingQty(skuPart)

                    If StrComp( _
                        cleanSku, _
                        selectedSku, _
                        vbTextCompare _
                    ) = 0 Then

                        If HasExplicitTrailingQty(skuPart) Then

                            totalQty = totalQty + _
                                       ExtractTrailingQty(skuPart)

                        End If

                    End If

                Next j

            End If

        End If

    Next i

    GetDeclaredSkuQty = totalQty

End Function