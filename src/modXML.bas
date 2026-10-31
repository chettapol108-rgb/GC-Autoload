Option Explicit

Public Sub ExportXML_UTF8_BySKU(ByVal wb As Workbook)
    Dim cfg As AutoLoadConfig, wsNorm As Worksheet, ws As Worksheet
    Dim skus As Object, documents As Object, sku As Variant
    Dim picker As Object, fso As Object, stream As Object
    Dim folder As String, baseFolder As String, suffix As Long, written As Long
    Dim failure As String
    On Error GoTo Failed
    cfg = ReadConfig(wb)
    Set wsNorm = EnsureNormalizedSheet(wb, cfg)
    If wsNorm Is Nothing Then Exit Sub
    Set skus = NormalizedSkus(wsNorm)
    RequireBatch wb, cfg, wsNorm, skus, True
    Set documents = CreateObject("Scripting.Dictionary")
    ' Build every document before creating files; never scan unrelated SKU sheets.
    For Each sku In skus.Keys
        ValidateSkuFileName CStr(sku)
        Set ws = wb.Worksheets(CStr(sku))
        documents.Add CStr(sku), BuildSkuXML(ws)
    Next sku
    Set picker = Application.FileDialog(4) ' msoFileDialogFolderPicker
    picker.Title = "Select a folder for all SKU XML files"
    If picker.Show <> -1 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    baseFolder = fso.BuildPath(picker.SelectedItems(1), "AutoLoad-" & cfg.ShopCode & "-" & Format$(Now, "yyyymmdd-hhnnss"))
    folder = baseFolder
    Do While fso.FolderExists(folder) Or fso.FileExists(folder)
        suffix = suffix + 1
        folder = baseFolder & "-" & suffix
    Loop
    fso.CreateFolder folder
    For Each sku In skus.Keys
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2
        stream.Charset = "utf-8"
        stream.Open
        stream.WriteText documents(CStr(sku))
        stream.SaveToFile fso.BuildPath(folder, CStr(sku) & ".xml"), 1 ' Never overwrite.
        stream.Close
        Set stream = Nothing
        written = written + 1
    Next sku
    MsgBox "XML export completed. Platform: " & cfg.ShopCode & "; Files: " & written & vbCrLf & folder, vbInformation
    Exit Sub
Failed:
    failure = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    MsgBox "XML export stopped: " & failure & vbCrLf & "Files written: " & written & vbCrLf & folder, vbExclamation
End Sub

Private Sub ValidateSkuFileName(ByVal sku As String)
    Dim ch As Variant, stem As String
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        If InStr(sku, CStr(ch)) > 0 Then Err.Raise vbObjectError + 2300, , "Invalid XML filename for SKU: " & sku
    Next ch
    stem = UCase$(Split(sku, ".")(0))
    If stem = "CON" Or stem = "PRN" Or stem = "AUX" Or stem = "NUL" Or _
       stem Like "COM#" Or stem Like "LPT#" Then
        Err.Raise vbObjectError + 2301, , "Reserved XML filename for SKU: " & sku
    End If
End Sub

Private Function BuildSkuXML(ByVal ws As Worksheet) As String
    Dim lastRow As Long, lastCol As Long
    Dim rowIdx As Long, colIdx As Long
    Dim xmlContent As String
    Dim variableNames() As String
    Dim i As Long
    Dim cellValue As String
    Dim varName As String
    Dim names As Object, re As Object, lastCell As Range

    ' Find last used row and column
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlValues, _
                               LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If lastCell Is Nothing Then Err.Raise vbObjectError + 2302, , "Empty output sheet: " & ws.Name
    lastRow = lastCell.Row
    If lastRow < 2 Then Err.Raise vbObjectError + 2303, , "No loaded data for SKU: " & ws.Name
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Read variable names from Row 1
    ReDim variableNames(1 To lastCol)
    Set names = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z_][A-Za-z0-9_.-]*$"
    For i = 1 To lastCol
        variableNames(i) = Trim$(CStr(ws.Cells(1, i).value))
        If Not re.Test(variableNames(i)) Then Err.Raise vbObjectError + 2304, , "Invalid XML variable in " & ws.Name & ", column " & i
        If names.Exists(variableNames(i)) Then Err.Raise vbObjectError + 2305, , "Duplicate XML variable: " & variableNames(i)
        names.Add variableNames(i), True
    Next i

    ' Build XML header exactly like Illustrator export
    xmlContent = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
    xmlContent = xmlContent & "<!DOCTYPE svg PUBLIC ""-//W3C//DTD SVG 1.1//EN"" ""http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd"" [" & vbCrLf
    xmlContent = xmlContent & " <!ENTITY ns_flows ""http://ns.adobe.com/Flows/1.0/"">" & vbCrLf
    xmlContent = xmlContent & " <!ENTITY ns_extend ""http://ns.adobe.com/Extensibility/1.0/"">" & vbCrLf
    xmlContent = xmlContent & " <!ENTITY ns_graphs ""http://ns.adobe.com/Graphs/1.0/"">" & vbCrLf
    xmlContent = xmlContent & " <!ENTITY ns_vars ""http://ns.adobe.com/Variables/1.0/"">" & vbCrLf
    xmlContent = xmlContent & " <!ENTITY ns_custom ""http://ns.adobe.com/GenericCustomNamespace/1.0/"">" & vbCrLf
    xmlContent = xmlContent & " <!ENTITY ns_svg ""http://www.w3.org/2000/svg"">" & vbCrLf
    xmlContent = xmlContent & " <!ENTITY ns_ai ""http://ns.adobe.com/AdobeIllustrator/10.0/"">" & vbCrLf
    xmlContent = xmlContent & " <!ENTITY ns_sfw ""http://ns.adobe.com/SaveForWeb/1.0/"">" & vbCrLf
    xmlContent = xmlContent & " <!ENTITY ns_adobe_xpath ""http://ns.adobe.com/XPath/1.0/"">" & vbCrLf
    xmlContent = xmlContent & " <!ENTITY ns_imrep ""http://ns.adobe.com/ImageReplacement/1.0/"">" & vbCrLf
    xmlContent = xmlContent & "]>" & vbCrLf
    xmlContent = xmlContent & "<svg version=""1.1""" & vbCrLf
    xmlContent = xmlContent & "  xmlns=""http://www.w3.org/2000/svg"" xmlns:xlink=""http://www.w3.org/1999/xlink"" xmlns:a=""http://ns.adobe.com/AdobeSVGViewerExtensions/3.0/""" & vbCrLf
    xmlContent = xmlContent & "  x=""0px"" y=""0px"" width=""0px"" height=""0px"" viewBox=""0 0 0 0"" style=""overflow:visible;"" xml:space=""preserve"">" & vbCrLf
    xmlContent = xmlContent & "<variableSets  xmlns=""&ns_vars;"">" & vbCrLf
    xmlContent = xmlContent & " <variableSet  locked=""none"" varSetName=""binding1"">" & vbCrLf
    xmlContent = xmlContent & "     <variables>" & vbCrLf

    ' Write variable declarations
    For i = 1 To lastCol
        xmlContent = xmlContent & "         <variable trait=""textcontent"" varName=""" & variableNames(i) & """></variable>" & vbCrLf
    Next i

    xmlContent = xmlContent & "     </variables>" & vbCrLf
    xmlContent = xmlContent & "     <v:sampleDataSets  xmlns="""" xmlns:v=""&ns_vars;"" xmlns:custom=""&ns_custom;"">" & vbCrLf

    ' Write each sampleDataSet for each data row
    For rowIdx = 2 To lastRow
        xmlContent = xmlContent & "         <v:sampleDataSet  dataSetName=""Data Set " & (rowIdx - 1) & """>" & vbCrLf
        For colIdx = 1 To lastCol
            
            cellValue = CStr(ws.Cells(rowIdx, colIdx).value)
            varName = Trim$(CStr(variableNames(colIdx)))
            
            ' Remove unwanted control characters first
            cellValue = CleanXMLText(cellValue)
            
            ' Convert || to line break only for NON-barcode fields
            ' Barcode fields are identified by output variable names ending with "_BC"
            If Right$(UCase$(varName), 2) <> "BC" Then
                cellValue = Replace(cellValue, "||", vbLf)
            End If
            
            ' Escape XML special characters
            cellValue = Replace(cellValue, "&", "&amp;")
            cellValue = Replace(cellValue, "<", "&lt;")
            cellValue = Replace(cellValue, ">", "&gt;")
            cellValue = Replace(cellValue, """", "&quot;")
            cellValue = Replace(cellValue, "'", "&apos;")
            
            If varName <> "" Then
            
                xmlContent = xmlContent & "             <" & varName & ">" & vbCrLf
                xmlContent = xmlContent & "                 <p>" & cellValue & "</p>" & vbCrLf
                xmlContent = xmlContent & "             </" & varName & ">" & vbCrLf
            
            End If
        Next colIdx
        xmlContent = xmlContent & "         </v:sampleDataSet>" & vbCrLf
    Next rowIdx

    xmlContent = xmlContent & "     </v:sampleDataSets>" & vbCrLf
    xmlContent = xmlContent & " </variableSet>" & vbCrLf
    xmlContent = xmlContent & "</variableSets>" & vbCrLf
    xmlContent = xmlContent & "</svg>" & vbCrLf

    BuildSkuXML = xmlContent
End Function
'------------------------------------------------------------------------------
' CleanXMLText
'
' Removes control characters that should not be written to AutoLoad XML.
'
' Unicode characters are intentionally preserved unchanged.
' This allows names in Thai, Chinese, Japanese, Korean, Arabic, Cyrillic,
' emoji, and other Unicode scripts to pass through safely.
'
' Removed:
' - ASCII control characters 0-31
' - TAB, LF and CR are also removed to preserve existing AutoLoad behavior.
'
' XML special characters such as &, <, >, " and ' are escaped separately
' by ExportXML_UTF8_BySKU.
'------------------------------------------------------------------------------
Private Function CleanXMLText(ByVal strIn As String) As String

    Dim i As Long
    Dim codeUnit As Long
    Dim ch As String
    Dim strOut As String

    strOut = ""

    For i = 1 To Len(strIn)

        ch = Mid$(strIn, i, 1)
        codeUnit = AscW(ch)

        ' AscW may return negative values for Unicode characters above U+7FFF.
        ' These are valid Unicode characters, so preserve them.
        If codeUnit < 0 Then

            strOut = strOut & ch

        ' Keep normal printable characters.
        ElseIf codeUnit >= 32 Then

            strOut = strOut & ch

        End If

        ' Character codes 0-31 are intentionally discarded.

    Next i

    CleanXMLText = strOut

End Function
