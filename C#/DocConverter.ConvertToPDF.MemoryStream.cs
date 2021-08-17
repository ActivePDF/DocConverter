using System;

namespace DocConverterExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            string strPath = System.AppDomain.CurrentDomain.BaseDirectory;
            string inputFile = $"{strPath}DocConverter.Word.Input.doc";
            byte[] inputFileBytes = System.IO.File.ReadAllBytes(inputFile);

            using (System.IO.MemoryStream inputStream = new System.IO.MemoryStream(inputFileBytes, 0, inputFileBytes.Length))
            {
                System.IO.MemoryStream outputStream = new System.IO.MemoryStream();
                
                // Instantiate Object
                APDocConverter.DocConverter docConverter =
                    new APDocConverter.DocConverter();

                // Enable extra logging (logging should only be used while
                // troubleshooting) C:\ProgramData\activePDF\Logs\
                docConverter.Debug = true;

                // Convert the file to a PDF MemoryStream
                DCDK.Results.DocConverterResult result =
                    docConverter.ConvertToPDF(
                        inputMemoryStream: inputStream,
                        APDocConverter.ToPDFFunction.FromDOC,
                        out outputStream);

                // Save the output to a local file.
                byte[] outputFile = outputStream.ToArray();
                outputStream.Close();
                System.IO.File.WriteAllBytes($"{strPath}output.pdf", outputFile);

                WriteResult(result);
            }

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
