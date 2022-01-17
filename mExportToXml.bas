Attribute VB_Name = "mExportToXml"
'VBA for Excel - Convert a Table into an xml file
Private filename As String
Private filepath As String
Private xmlStr As String
Sub ExportData() 'Use to call from Button
    Call BeginXmlConversion
End Sub
Sub BeginXmlConversion(Optional tblHeaders As String = vbNullString, Optional tblData As String = vbNullString)
    On Error GoTo Error_Handle
    If tblHeaders = vbNullString Then: tblHeaders = "SourceData[#Headers]" 'Replace With TableName Target Table Name e.g., Table1
    If tblData = vbNullString Then: tblData = "SourceData[#All]" 'Replace With TableName Target Table Name e.g., Table1

    filepath = ActiveWorkbook.Path & "\" 'Change Path Here
    filename = "SourceData.xml" 'Change filename here

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
    ReplaceChar = Replace(str, " ", vbNullString)
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


