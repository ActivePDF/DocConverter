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

' Stamp Images and Text onto the output PDF
oDC.AddStampCollection "TXTinternal"
oDC.StampFillMode = 2
oDC.StampFont = "Helvetica"
oDC.StampFontSize = 108
oDC.StampFontTransparency = 0.3
oDC.StampRotation = 45.0
oDC.SetStampColor 255, 0, 0, 0
oDC.SetStampStrokeColor 100, 0, 0, 0
oDC.AddStampText 116.0, 156.0, "Internal Only"

' Set whether the stamp collection(s) appears in the background or foreground
oDC.StampBackground = 0

' Convert the file to PDF
Set result = oDC.ConvertToPDF(_
        strPath & "DocConverter.Word.Input.doc",_
        strPath & "DocConverter.AddStampText.pdf")

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