Public Sub ExportXML_UTF8_BySKU(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rowIdx As Long, colIdx As Long
    Dim xmlContent As String
    Dim fileName As Variant
    Dim UTFStream As Object
    Dim variableNames() As String
    Dim i As Long
    Dim skuName As String
    Dim cfg As AutoLoadConfig
    Dim cellValue As String
    Dim varName As String

    cfg = ReadConfig(wb)

    ' Set input worksheet to read SKU name
    skuName = cfg.sku

    If skuName = "" Then
        MsgBox "SKU name in cell B1 is empty.", vbExclamation
        Exit Sub
    End If

    ' Set worksheet based on SKU name
    On Error Resume Next
    Set ws = wb.Sheets(skuName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & skuName & "' does not exist.", vbCritical
        Exit Sub
    End If

    ' Find last used row and column
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Read variable names from Row 1
    ReDim variableNames(1 To lastCol)
    For i = 1 To lastCol
        variableNames(i) = Trim$(CStr(ws.Cells(1, i).value))
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

    ' Prompt for file name (default to SKU name)
    fileName = Application.GetSaveAsFilename(skuName & ".xml", "XML Files (*.xml), *.xml")
    If fileName = False Then Exit Sub

    ' Create UTF-8 encoded file using ADODB.Stream
    Set UTFStream = CreateObject("ADODB.Stream")
    UTFStream.Type = 2 ' Text
    UTFStream.Charset = "utf-8"
    UTFStream.Open
    UTFStream.WriteText xmlContent
    UTFStream.SaveToFile fileName, 2
    UTFStream.Close

    MsgBox "Export completed from sheet '" & skuName & "'!" & vbCrLf & fileName, vbInformation
End Sub
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
