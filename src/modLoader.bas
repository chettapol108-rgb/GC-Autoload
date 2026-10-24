Public Sub LoadData_ToSKUSheet(ByVal wb As Workbook)
    Dim cfg As AutoLoadConfig
    Dim wsOutput As Worksheet
    Dim wsNorm As Worksheet
    Dim lastRow As Long
    Dim totalCopiesPerSet As Long
    Dim i As Long, outRow As Long, dataSetCount As Long
    Dim startCol As Long, destCol As Long
    Dim tempValue As String
    Dim rawBarcode As String
    Dim encodedBarcode As String
    Dim j As Long
    Dim copyCount As Long
    Dim copyIdx As Long
    Dim colDATE As Long, colSHOP As Long, ColSku As Long, colFN As Long
    Dim rowNum As Long

    cfg = ReadConfig(wb)
    ' Calculate configuration values before validating them.
    totalCopiesPerSet = cfg.CopyNameA + cfg.CopyNameB + cfg.CopyNameC + cfg.CopyNameD + cfg.CopyBarcode

    If cfg.SetsPerRow <= 0 Then
        MsgBox "Sets per row in A3 must be greater than zero.", vbExclamation
        Exit Sub
    End If

    If totalCopiesPerSet <= 0 Then
        MsgBox "At least one output copy value in B3:F3 must be greater than zero.", vbExclamation
        Exit Sub
    End If
    Set wsNorm = EnsureNormalizedSheet(wb, cfg)
    If wsNorm Is Nothing Then Exit Sub
    
    If cfg.sku = "" Then
        MsgBox "SKU name in B1 is empty.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsOutput = wb.Sheets(cfg.sku)
    On Error GoTo 0
    
    If wsOutput Is Nothing Then
        MsgBox "Sheet '" & cfg.sku & "' does not exist.", vbCritical
        Exit Sub
    End If

    wsOutput.rows("2:" & wsOutput.rows.Count).ClearContents

    colDATE = GetHeaderColumn(wsOutput, "DATE")
    colSHOP = GetHeaderColumn(wsOutput, "SHOP")
    ColSku = GetHeaderColumn(wsOutput, "SKU")
    colFN = GetHeaderColumn(wsOutput, "FN")

    If colDATE = 0 Or colSHOP = 0 Or ColSku = 0 Or colFN = 0 Then
        MsgBox "Missing required output columns: DATE, SHOP, SKU, or FN.", vbCritical
        Exit Sub
    End If

    lastRow = wsNorm.Cells(wsNorm.rows.Count, "F").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No normalized data found. Please run Translate Input first.", vbExclamation
        Exit Sub
    End If

    outRow = 2
    dataSetCount = 0
        
    For i = 2 To lastRow
        If UCase$(Trim$(CStr(wsNorm.Cells(i, "A").value))) = UCase$(cfg.sku) Then
        
            copyCount = CLng(Val(wsNorm.Cells(i, "G").value))
            If copyCount <= 0 Then copyCount = 1
            
            rawBarcode = wsNorm.Cells(i, "F").Text
            
            If rawBarcode <> "" Then
                encodedBarcode = CODE128_LIBRE(rawBarcode)
            Else
                encodedBarcode = ""
            End If
            
            For copyIdx = 1 To copyCount
                startCol = (dataSetCount Mod cfg.SetsPerRow) * totalCopiesPerSet + 1

                If dataSetCount > 0 And (dataSetCount Mod cfg.SetsPerRow) = 0 Then
                    outRow = outRow + 1
                End If

                destCol = startCol

                tempValue = wsNorm.Cells(i, "B").value
                For j = 1 To cfg.CopyNameA
                    wsOutput.Cells(outRow, destCol).value = tempValue
                    destCol = destCol + 1
                Next j

                tempValue = wsNorm.Cells(i, "C").value
                For j = 1 To cfg.CopyNameB
                    wsOutput.Cells(outRow, destCol).value = tempValue
                    destCol = destCol + 1
                Next j

                tempValue = wsNorm.Cells(i, "D").value
                For j = 1 To cfg.CopyNameC
                    wsOutput.Cells(outRow, destCol).value = tempValue
                    destCol = destCol + 1
                Next j

                tempValue = wsNorm.Cells(i, "E").value
                For j = 1 To cfg.CopyNameD
                    wsOutput.Cells(outRow, destCol).value = tempValue
                    destCol = destCol + 1
                Next j

                For j = 1 To cfg.CopyBarcode
                    wsOutput.Cells(outRow, destCol).value = encodedBarcode
                    destCol = destCol + 1
                Next j

                dataSetCount = dataSetCount + 1
            Next copyIdx
        End If
    Next i

    For rowNum = 2 To outRow
        wsOutput.Cells(rowNum, colDATE).value = cfg.DateYMD
        wsOutput.Cells(rowNum, colSHOP).value = cfg.ShopCode
        wsOutput.Cells(rowNum, ColSku).value = cfg.skuCode
        wsOutput.Cells(rowNum, colFN).value = cfg.DateMMDD & "-" & cfg.ShopCode & "-" & cfg.sku & "-" & Format$(rowNum - 1, "00")
    Next rowNum

    MsgBox "Data loading completed to sheet '" & cfg.sku & "'." & vbCrLf & _
           "Total sets loaded: " & dataSetCount & vbCrLf & _
           "Rows written: " & (outRow - 1), vbInformation
End Sub


Public Sub ClearData_BySKU(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim skuName As String
    Dim cfg As AutoLoadConfig

    cfg = ReadConfig(wb)
    skuName = cfg.sku

    If skuName = "" Then
        MsgBox "SKU name in cell B1 is empty.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set ws = wb.Sheets(skuName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet '" & skuName & "' does not exist.", vbCritical
        Exit Sub
    End If

    ws.rows("2:" & ws.rows.Count).ClearContents

    MsgBox "Data from row 2 down in sheet '" & skuName & _
           "' has been cleared.", vbInformation
End Sub

Private Function GetHeaderColumn(ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim colIdx As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For colIdx = 1 To lastCol
        If UCase$(Trim$(CStr(ws.Cells(1, colIdx).value))) = UCase$(headerName) Then
            GetHeaderColumn = colIdx
            Exit Function
        End If
    Next colIdx
    
    GetHeaderColumn = 0
End Function