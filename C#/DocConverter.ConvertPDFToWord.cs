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

            // Settings specific to other formats created with from PDF
            // conversions via Solid Documents SDK are set using the
            // FromPDFOptions object
            APDocConverter.FromPDFOptions fromPDFOptions =
                new APDocConverter.FromPDFOptions();

            // To Word options
            fromPDFOptions.ToWordHeadersAndFootersMode =
                APDocConverter.ToWordHeadersAndFootersOptions.Detect;

            // Confirm the from PDF settings for conversion via Solid Documents
            // SDK
            docConverter.SetFromPDFOptions(fromPDFOptions);

            // Convert the Word document from PDF to another format using Solid
            // Documents SDK. The second parameter determines the output file
            // format
            DCDK.Results.DocConverterResult results =
                docConverter.ConvertFromPDF(
                    $"{strPath}DocConverter.PDF.Input.pdf",
                    APDocConverter.FromPDFFunction.ToWordDocX,
                    $"{strPath}DocConverter.ConvertPDFToWord.Output.docx");
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
