' Copyright (c) 2019 ActivePDF, Inc.
' ActivePDF DocConverter

Dim FSO, strPath, results

' Get current path
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
Set FSO = Nothing

' Instantiate DocCOnverter object
Set oDC = CreateObject("APDocConverter.Object")

' Enable extra logging (logging should only be used while troubleshooting)
' C:\ProgramData\activePDF\Logs\
oDC.Debug = True

' Settings specific to CAD conversions
Set oCAD = CreateObject("CADConverter.Object")

' Options available for CAD conversions
' Options that are 'NotSet' use the setting from the configuation
' manager
oCAD.ASCIIHexEncoding = False
oCAD.Color = 0
oCAD.EmbedFonts = 2
oCAD.ExportLayers = 2
oCAD.ExportLayouts = 0
oCAD.FlateCompression = False
oCAD.HatchToBitmapDPI = 150
oCAD.HiddenLineRemoval = False
oCAD.Lineweight = -1
oCAD.OtherHatchesSettings = 0
oCAD.PlotStyleTableName = "Example"
oCAD.SHXTextAsGeometry = False
oCAD.SimpleGeometryOptimization = False
oCAD.SolidHatchesSettings = 2
oCAD.TrueTypeAsGeometry = False
oCAD.ZoomToExtents = True

' Convert the file to PDF
' If the output parameter is not used the created PDF will use
' the input string substituting the filename extension to 'pdf'
Set result = oDC.ConvertToPDF( _
        strPath & "DocConverter.CAD.Input.dwg", _
        strPath & "DocConverter.ConvertCADToPDF.Options.pdf")
WriteResult result

' Release Object
Set oCAD = Nothing
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