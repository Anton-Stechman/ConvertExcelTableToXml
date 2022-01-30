'GitHub Repository: https://github.com/Anton-Stechman/ConvertExcelTableToXml
'VBA for Excel - Convert a Table into an xml file
'Version 0.2.0

Private filename As String
Private filepath As String
Private xmlStr As String
Private HideFile As Boolean
Private CurAddrs As String
Private NoCellErrors As Boolean

'Initial Macro - Call From a Button
Sub RunXmlExport()
    'Add Optional Inputs Here, e.g., Call BeginMainLoop(CustomPath:="C:\Users\user\Documents\")
    Call BeginMainLoop
End Sub

'Used as a Helper to Set Variables Before Entering the MainLoop
Private Sub BeginMainLoop(Optional CustomPath As String = vbNullString, Optional CustomFile As String = vbNullString, _
Optional DataRange As String = vbNullString, Optional HeadRange As String = vbNullString)
    If CustomPath <> vbNullString Then: filepath = CustomPath
    If CustomFile <> vbNullString Then: filename = CustomFile
    If DataRange = vbNullString Then: DataRange = "TableName[#All]" 'Replace With TableName Target Table Name e.g., Table1[#All]
    If HeadRange = vbNullString Then: HeadRange = "TableName[#Headers]" 'Replace With TableName Target Table Name e.g., Table1[#Headers]
    HideFile = True
    Call MainLoop(tblHeaders:=HeadRange, tblData:=DataRange)
End Sub

'Main Loop
Private Sub MainLoop(Optional tblHeaders As String = vbNullString, Optional tblData As String = vbNullString)
    On Error GoTo Error_Handle
    If filepath = vbNullString Then: filepath = Range("TargetPath").Value 'Can Change Path Here
    If Right(filepath, 1) <> "\" Then: filepath = filepath & "\"
    If DirExists(filepath) = False Then: MkDir (filepath)
    If filename = vbNullString Then: DateVal = Format(Now, "YYYY-MM"): filename = DateVal & "_MBM_SourceData.xml" 'Can Change filename here
    Call OptimiseVBA
    xmlStr = FormatForXml(HeaderRow:=Range(tblHeaders), TableRange:=Range(tblData))
    Call CreateNewXml(CStr(xmlStr))
    Call SetAttr(filepath & filename, IIf(HideFile = True, vbHidden, vbNormal))
    Call MsgBox("xml Export Complete!" & vbNewLine & "Press 'Ok' To Finish", vbOKOnly, "Success!")
    Call OptimiseVBA(True)
    Debug.Print (xmlStr)
    Exit Sub
    
Error_Handle:
    Call MsgBox("An Error Occured!" & vbNewLine & "Error Number:" & Space(1) & _
    Err.Number & vbNewLine & "Description:" & Space(1) & Err.Description & _
    IIf(CurAddrs = vbNullString, vbNewLine & "Error in Cell:" _
    & Space(1) & CurAddrs, vbNullString), vbOKOnly, "Error!")
    Call OptimiseVBA(True)
End Sub

'Format Target Range Values For XML
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
            If newHeader = vbNullString Then: newHeader = "blank_field_Col" & h.Column
            With h.Offset(i, 0)
                CurAddrs = .Address
                If IsEmpty(.Value) Then: v = "null": Else v = .Value
            End With
            str = str & vbTab & vbTab & "<" & newHeader & ">" & v & "</" & newHeader & ">" & vbNewLine
        Next h
        str = str & vbTab & "</SourceData>" & vbNewLine
        If i >= MaxIterations Then: Exit For
    Next i
    
    'Close off xml formatting
    str = str & "</SourceDataTable>" & vbNewLine
    str = Replace(str, "_>", ">")
    str = Replace(str, "<>", "<blank_field>")
    str = Replace(str, "</>", "</blank_field>")
    FormatForXml = str
    NoCellErrors = True
End Function

'Generate XML File
Private Sub CreateNewXml(contents As String)
    If DirExists(filepath & filename) = False Then: Call SetAttr(filepath & filename, vbNormal)
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "UTF-8"
    objStream.Open
    Call objStream.WriteText(contents)
    Call objStream.SaveToFile(filepath & filename, 2)
    objStream.Close
End Sub

'Remove Illegal Characters
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

'Check Directory Exists
Private Function DirExists(Optional dirStr As String)
    If dirStr = vbNullString Then: dirstring = filepath & "\Data\"
    DirExists = Dir(dirStr, vbDirectory) <> vbNullString
End Function

'Turn VBA Optimisation On/Off
Private Sub OptimiseVBA(Optional switch As Boolean = False)
    Dim calcsettings As Variant
    If switch = True Then: calcsettings = xlAutomatic: Else: calcsettings = xlManual
    Application.ScreenUpdating = switch
    Application.EnableEvents = switch
    Application.Calculation = calcsettings
End Sub
