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

            // ActivePDF Service runs under the 'Local System' account, as such
            // it will only see certificates located in HKEY_LOCAL_MACHINE. To
            // use certificates located under a specific user set the
            // impersonate user options in the GUI or API
            docConverter.InvisiblySignFile(
                CommonName: "localhost",
                CertificateStore: "My",
                UseLocalMachine: true,
                Location: "Mission Viejo, CA",
                Reason: "Approval",
                ContactInfo: "949-555-1212",
                SignatureType: 1);

            // Convert the file to PDF
            // If the output parameter is not used the created PDF will use
            // the input string substituting the filename extension to 'pdf'
            DCDK.Results.DocConverterResult result =
                docConverter.ConvertToPDF(
                    inputPath: $"{strPath}DocConverter.Word.Input.doc",
                    outputPDF: $"{strPath}DocConverter.DigitallySignPDF.pdf");

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
