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

' Start a remote conversion by opening a connection Specify the IP address of
' the machine with the DocConverter installation If 0 is used as the port
' number DocConverter will use the default port
oDC.OpenRemoteConnection "10.1.11.88", 62625

' Convert the file to PDF
Set result = oDC.ConvertToPDF(_
        strPath & "DocConverter.Word.Input.doc",_
        strPath & "DocConverter.RemoteConversion.pdf")

' Close the remote connection when finished
docConverter.CloseRemoteConnection

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