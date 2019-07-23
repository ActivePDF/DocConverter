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

' ActivePDF Service runs under the 'Local System' account, as such it will only
' see certificates located in HKEY_LOCAL_MACHINE. To use certificates located
' under a specific user set the impersonate user options in the GUI or API
oDC.InvisiblySignFile "localhost", "My", true, "Mission Viejo, CA", "Approval", "949-555-1212", 1

' Convert the file to PDF
Set result = oDC.ConvertToPDF(_
        strPath & "DocConverter.Word.Input.doc",_
        strPath & "DocConverter.DigitallySignPDF.pdf")

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