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

' Settings specific to other formats created with from PDF
' conversions via Solid Documents SDK are set using the
' FromPDFOptions object
Set fPDF = CreateObject("APDocConverter.FromPDFOptions")

' To Image options
fPDF.ToImagePageDPI = 300

' Confirm the from PDF settings for conversion via Solid Documents SDK
oDC.SetFromPDFOptions fPDF

' Convert the Word document to PDF
Set result = oDC.ConvertFromPDF( _
        strPath & "DocConverter.PDF.Input.pdf", _
        7, _
        strPath & "DocConverter.ToImageJPEG.Output.jpg")

' Output result
WriteResult result

' Release Object
Set fPDF = Nothing
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