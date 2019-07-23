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

            // Instantiate Object
            APDocConverter.DocConverter docConverter =
                new APDocConverter.DocConverter();

            // Enable extra logging (logging should only be used while
            // troubleshooting) C:\ProgramData\activePDF\Logs\
            docConverter.Debug = true;

            // Add stamp image onto the output PDF
            docConverter.AddStampCollection(collectionName: "IMGimage");
            docConverter.AddStampImage(
                ImageFile: $"{strPath}DocConverter.Input.jpg",
                x: 508.0f,
                y: 16.0f,
                Width: 32.0f,
                Height: 32.0f,
                PersistRatio: true);

            // Set whether the stamp collection(s) appears in the background or
            // foreground
            docConverter.StampBackground = 0;

            // Convert the file to PDF
            // If the output parameter is not used the created PDF will use
            // the input string substituting the filename extension to 'pdf'
            DCDK.Results.DocConverterResult result =
                docConverter.ConvertToPDF(
                    inputPath: $"{strPath}DocConverter.Word.Input.doc",
                    outputPDF: $"{strPath}DocConverter.AddStampImage.pdf");

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
