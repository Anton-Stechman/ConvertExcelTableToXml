'GitHub Repository: https://github.com/Anton-Stechman/ConvertExcelTableToXml
'VBA for Excel - Convert a Table into an xml file
Private filename As String
Private filepath As String
Private xmlStr As String

Sub RunXmlExport() 'Can be called from Button
	Call ExportData()
End Sub

Sub ExportData(Optional CustomPath As String = vbNullString, Optional CustomFile As String = vbNullString, _
Optional DataRange As String = vbNullString, Optional HeadRange As String = vbNullString)
    If CustomPath <> vbNullString Then: filepath = CustomPath
    If CustomFile <> vbNullString Then: filename = CustomFile
			If DataRange = vbNullString Then: DataRange = "SourceData[#All]"  'Replace With TableName Target Table Name e.g., Table1
				If HeadRange = vbNullString Then: HeadRange = "SourceData[#Headers]" 'Replace With TableName Target Table Name e.g., Table1
    Call Main(tblHeaders:=HeadRange, tblData:=DataRange)
End Sub
								
Sub Main(Optional tblHeaders As String = vbNullString, Optional tblData As String = vbNullString)
    On Error GoTo Error_Handle

    If filepath = vbNullString Then: filepath = ActiveWorkbook.Path & "\Data\" 'Change Path Here
    If DirExists(filepath) = False Then: MkDir (filepath)
    If filename = vbNullString Then
        Dim DateVal As String: DateVal = Format(Now, "YYYY-MM")
	filename = DateVal & "NewXmlExport.xml" 'Change filename here
    End If
    
    Call OptimiseVBA
    Call MsgBox("Exporting Data to xml..." & vbNewLine & "Click 'Ok' To Continue", vbOKOnly, "Begin xml Conversion")
    xmlStr = FormatForXml(HeaderRow:=Range(tblHeaders), TableRange:=Range(tblData))
    Call CreateNewXml(CStr(xmlStr))
    Call MsgBox("xml Export Complete!" & vbNewLine & "Press 'Ok' To Finish", vbOKOnly, "Success!")
    Call OptimiseVBA(True)
    Exit Sub
    
Error_Handle:
    Call MsgBox("An Error Occured!" & vbNewLine & "Error Number:" & Space(1) & _
    Err.Number & vbNewLine & "Description:" & Space(1) & Err.Description, vbOKOnly, "Error!")
    Call OptimiseVBA(True)
End Sub
Private Function FormatForXml(Optional HeaderRow As Range, Optional TableRange As Range, Optional MaxIterations As Integer = 1000) As String
    'Set Variables
    Dim str As String
    Dim Q As String: Q = Chr$(34)
    
    'Initiate xml Format
    str = "<?xml version=" & Q & "1.0" & Q & Space(1) & "encoding=" & Q & "UTF-8" & Q & "?>" & vbNewLine
    str = str & "<SourceDataTable>" & vbNewLine
    
    'Format Input Table for xml
    For i = 1 To TableRange.Rows.Count
        str = str & vbTab & "<SourceData>" & vbNewLine
        For Each h In HeaderRow
            Dim newHeader: newHeader = ReplaceChar(CStr(h.Value))
            With h.Offset(i, 0)
                If IsNumeric(.Value) = True Then
                    If IsEmpty(.Value) Then: v = "null": Else v = .Value
                Else
                    v = .Value
                End If
            End With
            str = str & vbTab & vbTab & "<" & newHeader & ">" & v & "</" & newHeader & ">" & vbNewLine
        Next h
        str = str & vbTab & "</SourceData>" & vbNewLine
        If i >= MaxIterations Then: Exit For
    Next i
    
    'Close off xml formatting
    str = str & "</SourceDataTable>" & vbNewLine
    str = Replace(str, "_>", ">")
    Debug.Print (str)
    FormatForXml = str
End Function
Private Sub OptimiseVBA(Optional switch As Boolean = False)
    Dim calcsettings As Variant
    If switch = True Then: calcsettings = xlAutomatic: Else: calcsettings = xlManual
    Application.ScreenUpdating = switch
    Application.EnableEvents = switch
    Application.Calculation = calcsettings
End Sub
Private Sub CreateNewXml(contents As String)
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "UTF-8"
    objStream.Open
    Call objStream.WriteText(contents)
    Call objStream.SaveToFile(filepath & filename, 2)
    objStream.Close
End Sub
Private Function ReplaceChar(str As String) As String
    ReplaceChar = Replace(str, " ", "_")
    ReplaceChar = Replace(ReplaceChar, ".000", vbNullString)
    For i = 1 To 47
        ReplaceChar = Replace(ReplaceChar, Chr$(i), vbNullString)
    Next i
    For i = 58 To 64
        ReplaceChar = Replace(ReplaceChar, Chr$(i), vbNullString)
    Next i
    If IsNumeric(Left(ReplaceChar, 1)) = True Then
        ReplaceChar = "n" & ReplaceChar
    End If
End Function
Private Function DirExists(Optional dirStr As String)
    If dirStr = vbNullString Then: dirstring = filepath & "\Data\"
    DirExists = Dir(dirStr, vbDirectory) <> vbNullString
End Function

