' Copyright (c) 2019 ActivePDF, Inc.
' ActivePDF DocConverter

Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate Object
Set oDC = CreateObject("APDocConverter.Object")

' Enable extra logging (logging should only be used while troubleshooting)
' C:\ProgramData\activePDF\Logs\
oDC.Debug = True

' Compresses eligible objects in the output PDF, which include page objects and
' fonts. Streams (including content, text, images, and data) are not affected.
' SetPDFCompression is only supported in PDF versions 1.5 and above. It is
' enabled by default.
oDC.SetPDFCompression True

' Convert the file to PDF
Set result = oDC.ConvertToPDF(_
        strPath & "DocConverter.Word.Input.doc",_
        strPath & "DocConverter.SetPDFCompression.pdf")

' Output result
WriteResult result

' Release Object
Set oDC = Nothing

' Process Complete
Wscript.Quit

Sub WriteResult(oResult)
  message = "Result Status: " & result.DocConverterStatus
  If result.DocConverterStatus = 0 Then
      message = message & ", Success!"
  Else
      message = message &", " & result.Details
  End If
  Wscript.Echo message
End Sub