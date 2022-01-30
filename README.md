# ConvertExcelTableToXml
VBA Script to convert an Excel Table or Range to an xml file

VBA script to export either an Excel Table or an Excel Range to an xml file

How it works:
  the script loops the row count of the set data range (for i = 0 to data.rows.count-1)
  next it then loops through all the headers (for each h in HeaderRange) 
  from here it uses an offset from the each header using the value i to get the 
  resut = "<" headername ">" & h.offset(i,0).value & "</" headername ">"
  
  this way the loop takes each value in a row and assigns it to it's header, all of this is stored as a string value then written to an xml file upon completion
  The script is more or less "plug-and-play" and only requires minimal editing to work for any specific workbook.
  
  To work with any specific work book in the first method "RunXmlExport()" 
  add in optional inputs to the line: Call BeginMainLoop(CustomPath="PathHere",CustomFilename:="MyFile",DataRange:="A2:Z100",HeadRange:="A1:Z1")
  RunXmlExport() can be called from a button, shape, hotkey or on Workbook_Close().
  
  
