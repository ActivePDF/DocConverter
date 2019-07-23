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

' Add a stamp image onto the output PDF
oDC.AddStampCollection "IMGimage"
oDC.AddStampImage strPath & "DocConverter.Input.jpg", 508.0, 16.0, 32.0, 32.0, True

' Set whether the stamp collection(s) appears in the background or foreground
oDC.StampBackground = 0

' Convert the file to PDF
Set result = oDC.ConvertToPDF(_
        strPath & "DocConverter.Word.Input.doc",_
        strPath & "DocConverter.AddStampImage.pdf")

' Output result
WriteResult result

' Release Object
Set oDC = Nothing

' Process Complete
Wscript.Quit result.DocConverterStatus

Sub WriteResult(oResult)
  message = "Result Status: " & result.DocConverterStatus
  If result.DocConverterStatus = 0 Then
      message = message & ", Success!"
  Else
      message = message &", " & result.Details
  End If
  Wscript.Echo message
End Sub