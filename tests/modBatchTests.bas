Attribute VB_Name = "modBatchTests"
Option Explicit

' Import with the source modules into an isolated development VBA project.
' Uses only synthetic data. Production MsgBox dialogs remain interactive.
' Leaves the unsaved synthetic workbook open for manual XML export checks.
Public Sub GC_TestBatchProcessing()
    Dim wb As Workbook, inputSheet As Worksheet, norm As Worksheet, db As Worksheet
    Dim cfg As AutoLoadConfig, skus As Object, text As String
    Dim blocked As Boolean, failure As String
    On Error GoTo Failed
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set inputSheet = wb.Worksheets(1)
    inputSheet.Name = "AutoLoad-Input"
    Set db = wb.Worksheets.Add(After:=inputSheet)
    db.Name = "Database"
    Set norm = wb.Worksheets.Add(After:=db)
    norm.Name = "NormalizedInput"
    WriteTestRow db, 1, Array("SKU", "Set", "Name A", "Name B", "Name C", "Name D", "OID", "Barcode")
    WriteTestRow db, 2, Array("TEST-A", 2, 1, 0, 0, 0, 1, 1)
    WriteTestRow db, 3, Array("TEST-B", 1, 2, 1, 0, 0, 0, 0)
    WriteTestRow norm, 1, Array("SKU", "NameA", "NameB", "NameC", "NameD", "OID", "Barcode", "Qty")
    WriteTestRow inputSheet, 3, Array("SKU", "Barcode (Order ID)", "PLATFORM", "Qty", "Raw Input")
    inputSheet.Range("A2").NumberFormat = "@"
    inputSheet.Range("A2").Value2 = "20261031"
    inputSheet.Range("B2").Value2 = "SP"
    WriteTestRow inputSheet, 1, Array("DATE", "PLATFORM")
    cfg = ReadConfig(wb)
    RequireTest cfg.FirstDataRow = 4 And cfg.DateYMD = "20261031" And cfg.ShopCode = "SP", "Compact input config"
    inputSheet.Range("A2").Value2 = "20261002(M)"
    cfg = ReadConfig(wb)
    RequireTest cfg.DateYMD = "20261002(M)" And cfg.DateMMDD = "1002(M)", "Text date suffix preserved"
    inputSheet.Range("A2").Value2 = "20261031"
    MakeTestOutput wb, "TEST-A", Array("N1-A1", "N1_OID", "N1_BC", "N2-A1", "N2-OID", "N2-BC", "DATE", "SHOP", "SKU", "FN")
    MakeTestOutput wb, "TEST-B", Array("N1-A1", "N1-A2", "N1-B1", "DATE", "SHOP", "SKU", "FN")
    ' Thai, Chinese, Korean, emoji surrogate pair, and all XML escape characters.
    text = ChrW(&HE01) & ChrW(&H4E2D) & ChrW(&HD55C) & ChrW(-10179) & ChrW(-8704) & " & < > "" '"
    WriteTestRow inputSheet, 4, Array("TEST-A", "00012345678901234567", " Shopee ", 1, "TEST-A *3pcs || " & text)
    WriteTestRow inputSheet, 5, Array("TEST-A", "00012345678901234567", "shopee", 2, "TEST-A *3pcs || " & text)
    WriteTestRow inputSheet, 6, Array("TEST-B", "00012345678901234567", "Shopee", 1, "#1|| BETA #2|| GAMMA")
    WriteTestRow inputSheet, 7, Array("TEST-A", "TT-001", "TikTok", 2, "ALPHA")
    WriteTestRow inputSheet, 8, Array("TEST-B", "FB-001", "Facebook", 1, "#1|| DELTA #2|| EPSILON")
    RequireTest MatchesPlatform(" SHOPEE ", "SP"), "Shopee comparison"
    RequireTest Not MatchesPlatform("Shopee", "TT"), "TT excludes Shopee"
    RequireTest MatchesPlatform("TikTok", "TT") And MatchesPlatform("Facebook", "TT"), "TT includes other platforms"
    RequireTest MatchesPlatform("Shopee", "ALL") And MatchesPlatform("TikTok", "ALL") And _
                MatchesPlatform("Facebook", "ALL") And MatchesPlatform("", "ALL"), "ALL includes every platform"
    RequireTest MatchesPlatform(CVErr(xlErrNA), "ALL"), "ALL does not inspect platform values"

    TranslateRawInput_ToNormalizedRows wb
    RequireTest GetBatchProperty(norm, "GC_BatchReady") = "YES", "SP translation"
    RequireTest norm.Cells(norm.Rows.Count, "A").End(xlUp).Row = 3, "Duplicate order+SKU grouped; different SKU kept"
    RequireTest norm.Range("H2").Value2 = 3, "Explicit quantity equals grouped quantity"
    RequireTest norm.Range("F2").Value2 = "00012345678901234567", "Long OID and leading zeros"
    RequireTest norm.Range("B2").Value2 = text, "Unicode unchanged"
    LoadData_ToSKUSheet wb
    RequireTest GetBatchProperty(norm, "GC_BatchLoaded") = "YES", "SP batch load"
    RequireTest wb.Worksheets("TEST-A").Range("A3").Value2 = text, "Sets wrap to next row"
    RequireTest wb.Worksheets("TEST-A").Range("D3").Value2 = "", "Partial final row"
    RequireTest wb.Worksheets("TEST-A").Range("B2").Value2 = "00012345678901234567", "Output plain-text OID"
    RequireTest wb.Worksheets("TEST-A").Range("H2").Value2 = "SP", "Output platform"
    RequireTest wb.Worksheets("TEST-B").Range("A2").Value2 = "BETA" And _
                wb.Worksheets("TEST-B").Range("B2").Value2 = "BETA" And _
                wb.Worksheets("TEST-B").Range("C2").Value2 = "GAMMA", "Per-SKU copy settings"

    ' A different B2 cannot consume the previous batch.
    inputSheet.Range("B2").Value2 = "TT"
    cfg = ReadConfig(wb)
    Set skus = NormalizedSkus(norm)
    On Error Resume Next
    RequireBatch wb, cfg, norm, skus, True
    blocked = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    RequireTest blocked, "Changed platform blocks old batch"
    TranslateRawInput_ToNormalizedRows wb
    RequireTest norm.Cells(norm.Rows.Count, "A").End(xlUp).Row = 3, "TT contains TikTok and Facebook only"
    RequireTest norm.Range("F2").Value2 = "TT-001" And norm.Range("H2").Value2 = 2, "No explicit quantity uses Qty"
    LoadData_ToSKUSheet wb
    RequireTest wb.Worksheets("TEST-B").Range("A2").Value2 = "DELTA", "Facebook routed to its SKU"

    inputSheet.Range("B2").Value2 = " all "
    TranslateRawInput_ToNormalizedRows wb
    RequireTest GetBatchProperty(norm, "GC_BatchReady") = "YES", "ALL translation"
    RequireTest norm.Cells(norm.Rows.Count, "A").End(xlUp).Row = 5, "ALL contains SP and TT groups"
    LoadData_ToSKUSheet wb
    RequireTest GetBatchProperty(norm, "GC_BatchLoaded") = "YES", "ALL batch load"
    RequireTest wb.Worksheets("TEST-A").Range("A4").Value2 = "ALPHA", "ALL retains TikTok after Shopee"
    RequireTest wb.Worksheets("TEST-B").Range("A2").Value2 = "BETA" And _
                wb.Worksheets("TEST-B").Range("A3").Value2 = "DELTA", "ALL loads Shopee and Facebook"
    RequireTest wb.Worksheets("TEST-A").Range("H2").Value2 = "ALL", "ALL output platform"

    ' One invalid destination must leave every SKU's old output untouched.
    wb.Worksheets("TEST-A").Range("A2").Value2 = "KEEP-A"
    wb.Worksheets("TEST-B").Range("D1").Value2 = "BROKEN-DATE"
    LoadData_ToSKUSheet wb
    RequireTest wb.Worksheets("TEST-A").Range("A2").Value2 = "KEEP-A", "Batch destination preflight"
    RequireTest GetBatchProperty(norm, "GC_BatchLoaded") = "NO", "Failed load blocks XML"
    wb.Worksheets("TEST-B").Range("D1").Value2 = "DATE"

    ' Mismatching explicit quantity and conflicting duplicate raw text.
    inputSheet.Range("B2").Value2 = "SP"
    inputSheet.Range("D5").Value2 = 1
    TranslateRawInput_ToNormalizedRows wb
    RequireTest GetBatchProperty(norm, "GC_BatchReady") = "NO", "Explicit quantity mismatch blocks batch"
    RequireTest inputSheet.Range("A4").Interior.Color = RGB(255, 199, 206), "Mismatch highlights source group"
    inputSheet.Range("D5").Value2 = 2
    inputSheet.Range("E5").Value2 = "TEST-A *3pcs || DIFFERENT"
    TranslateRawInput_ToNormalizedRows wb
    RequireTest GetBatchProperty(norm, "GC_BatchReady") = "NO", "Conflicting raw input blocks batch"
    inputSheet.Range("E5").Value2 = inputSheet.Range("E4").Value2

    ' Missing and duplicate Database SKUs fail before replacing old normalized rows.
    inputSheet.Range("A4:A5").Value2 = "MISSING-X"
    TranslateRawInput_ToNormalizedRows wb
    RequireTest GetBatchProperty(norm, "GC_BatchReady") = "NO", "Missing Database SKU blocks batch"
    inputSheet.Range("A4:A5").Value2 = "TEST-A"
    WriteTestRow db, 4, Array("TEST-A", 2, 1, 0, 0, 0, 1, 1)
    TranslateRawInput_ToNormalizedRows wb
    RequireTest GetBatchProperty(norm, "GC_BatchReady") = "NO", "Duplicate Database SKU blocks batch"
    db.Rows(4).ClearContents

    ' Final valid SP batch remains available for manual XML and Ribbon tests.
    TranslateRawInput_ToNormalizedRows wb
    LoadData_ToSKUSheet wb
    RequireTest GetBatchProperty(norm, "GC_BatchLoaded") = "YES", "Final valid batch"
    MakeTestOutput wb, "TEST-UNUSED", Array("N1-A1")
    wb.Worksheets("TEST-UNUSED").Range("A2").Value2 = "KEEP-UNUSED"
    WriteTestRow db, 4, Array("TEST-IDLE", 1, 1, 0, 0, 0, 0, 0)
    MakeTestOutput wb, "TEST-IDLE", Array("N1-A1", "DATE", "SHOP", "SKU", "FN")
    wb.Worksheets("TEST-IDLE").Range("A2").Value2 = "OLD-IDLE-DATA"
    ' Clear must work without a successful Translate or valid Date/Platform.
    SetBatchProperty norm, "GC_BatchReady", "NO"
    inputSheet.Range("A2").Value = CVErr(xlErrNA)
    inputSheet.Range("B2").Value = CVErr(xlErrNA)
    WriteTestRow db, 5, Array("MISSING-SHEET", 1, 1, 0, 0, 0, 0, 0)
    ClearData_BySKU wb
    RequireTest wb.Worksheets("TEST-IDLE").Range("A2").Value2 = "OLD-IDLE-DATA" And _
                wb.Worksheets("TEST-A").Range("A2").Value2 = text, "Clear preflights every destination"
    db.Rows(5).ClearContents
    ClearData_BySKU wb
    RequireTest wb.Worksheets("TEST-A").Range("A2").Value2 = "" And _
                wb.Worksheets("TEST-B").Range("A2").Value2 = "", "Clear all Database SKUs without a batch"
    RequireTest wb.Worksheets("TEST-IDLE").Range("A2").Value2 = "", "Clear includes SKUs absent from input"
    RequireTest wb.Worksheets("TEST-IDLE").Range("A1").Value2 = "N1-A1", "Clear preserves headers"
    RequireTest inputSheet.Range("A4").Value2 = "TEST-A" And norm.Range("A2").Value2 = "TEST-A" And _
                db.Range("A2").Value2 = "TEST-A", "Clear preserves supporting sheets"
    RequireTest GetBatchProperty(norm, "GC_BatchLoaded") = "NO", "Clear blocks export until reload"
    RequireTest wb.Worksheets("TEST-UNUSED").Range("A2").Value2 = "KEEP-UNUSED", "Clear preserves unrelated sheets"
    inputSheet.Range("A2").Value2 = "20261031"
    inputSheet.Range("B2").Value2 = "SP"
    TranslateRawInput_ToNormalizedRows wb
    LoadData_ToSKUSheet wb
    RequireTest GetBatchProperty(norm, "GC_BatchLoaded") = "YES", "Reload after clear"
    wb.Activate
    MsgBox "Batch assertions passed. Synthetic workbook left open for XML export checks. No file was saved.", vbInformation
    Exit Sub
Failed:
    failure = Err.Description
    MsgBox "BATCH TEST FAILED: " & failure, vbCritical
    Err.Raise vbObjectError + 2400, "GC_TestBatchProcessing", failure
End Sub

Private Sub WriteTestRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal values As Variant)
    Dim i As Long
    For i = 0 To UBound(values)
        If ws.Name = "AutoLoad-Input" And rowIndex >= 4 And i = 1 Then ws.Cells(rowIndex, i + 1).NumberFormat = "@"
        ws.Cells(rowIndex, i + 1).Value2 = values(i)
    Next i
End Sub

Private Sub MakeTestOutput(ByVal wb As Workbook, ByVal sku As String, ByVal headers As Variant)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sku
    WriteTestRow ws, 1, headers
End Sub

Private Sub RequireTest(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2401, "Batch assertion", message
End Sub
