// Copyright (c) 2019 ActivePDF, Inc.
// ActivePDF DocConverter

using System;

// Make sure to add the ActivePDF product .NET DLL(s) to your application.
// .NET DLL(s) are typically found in the products 'bin' folder.
namespace DocConverterExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            string strPath = System.AppDomain.CurrentDomain.BaseDirectory;

            // Instantiate DocConverter object
            APDocConverter.DocConverter docConverter =
                new APDocConverter.DocConverter();

            // Enable extra logging (logging should only be used while
            // troubleshooting) C:\ProgramData\activePDF\Logs\
            docConverter.Debug = true;

            // Instantiate the CAD settings object
            APDocConverter.CADConverter.CADConverter cadConverterSettings =
                new APDocConverter.CADConverter.CADConverter();

            // Options available for CAD conversions
            // Options that are 'NotSet' use the setting from the configuation
            // manager
            cadConverterSettings.ASCIIHexEncoding = false;
            cadConverterSettings.Color = APDocConverter.CADConverter.ColorMode.NotSet;
            cadConverterSettings.EmbedFonts = APDocConverter.CADConverter.EmbedFontsOptions.NotSet;
            cadConverterSettings.ExportLayers = APDocConverter.CADConverter.LayersOptions.NotSet;
            cadConverterSettings.ExportLayouts = APDocConverter.CADConverter.LayoutsOptions.NotSet;
            cadConverterSettings.FlateCompression = false;
            cadConverterSettings.HatchToBitmapDPI = 150;
            cadConverterSettings.HiddenLineRemoval = false;
            cadConverterSettings.Lineweight = APDocConverter.CADConverter.LineweightOptions.NotSet;
            cadConverterSettings.OtherHatchesSettings = APDocConverter.CADConverter.OtherHatchOptions.NotSet;
            cadConverterSettings.PlotStyleTableName = "";
            cadConverterSettings.SHXTextAsGeometry = false;
            cadConverterSettings.SimpleGeometryOptimization = false;
            cadConverterSettings.SolidHatchesSettings = APDocConverter.CADConverter.SolidHatchOptions.NotSet;
            cadConverterSettings.TrueTypeAsGeometry = false;
            cadConverterSettings.ZoomToExtents = true;

            // Convert the file to PDF
            // If the output parameter is not used the created PDF will use
            // the input string substituting the filename extension to 'pdf'
            DCDK.Results.DocConverterResult result =
                docConverter.ConvertToPDF(
                    $"{strPath}DocConverter.CAD.Input.dwg",
                    $"{strPath}DocConverter.ConvertCADToPDF.Options.pdf");

            WriteResult(result);

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        // Result Handling
        public static void WriteResult(DCDK.Results.DocConverterResult Result)
        {
            Console.WriteLine($"Result Origin: {Result.Origin.Class}.{Result.Origin.Function}");
            Console.WriteLine($"Result Status: {Result.DocConverterStatus}");
            if (!string.IsNullOrEmpty(Result.Details))
            {
                Console.WriteLine($"Result Details: {Result.Details}");
            }
            if (Result.ResultException != null)
            {
                Console.WriteLine("Exception caught during conversion.");
                Console.WriteLine($"Excpetion Details: {Result.ResultException.Message}");
                Console.WriteLine($"Exception Stack Trace: {Result.ResultException.StackTrace}");
            }
        }
    }
}
