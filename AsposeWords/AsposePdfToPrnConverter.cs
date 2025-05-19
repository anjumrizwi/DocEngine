//using System.Reflection.Metadata;
using System.Diagnostics;
using Aspose.Pdf;
using Aspose.Pdf.Devices;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Printing;
namespace AsposeWords
{
    public class AsposePdfToPrnConverter
    {

        /// <summary>
        /// Convert PDF to PRN Folder
        /// </summary>
        public static void ConvertAllPdfToPrnInFolder(string inputFolder, string outputFolder)
        {
            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");

            foreach (var file in pdfFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                string outputFile = Path.Combine(outputFolder, fileName + ".prn");

                try
                {
                    Console.WriteLine($"Converting: {fileName}.pdf → {fileName}.prn");
                    var stopwatch = Stopwatch.StartNew();
                    ConvertPdfToPrn(file, outputFile);
                    stopwatch.Stop();
                    Console.WriteLine($"[SUCCESS] {fileName} converted in {stopwatch.Elapsed.TotalSeconds:F2} seconds");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error converting {fileName}: {ex.Message}");
                }
            }

            Console.WriteLine("✅ PRN Batch conversion complete.");
        }

        public static void ConvertPdfToPrn(string sourceFile, string prnFile)
        {
            Document pdfDocument = new Document(sourceFile);

            PrinterSettings printerSettings = new PrinterSettings
            {
                PrinterName = "Microsoft Print to PDF",
                PrintToFile = true,
                PrintFileName = prnFile,
               

            };

            printerSettings.DefaultPageSettings.Margins = new Margins(20, 20, 20, 20);
            printerSettings.DefaultPageSettings.Landscape = false;
            printerSettings.DefaultPageSettings.Color = false;
            printerSettings.DefaultPageSettings.PrinterResolution.Kind = PrinterResolutionKind.High;
            printerSettings.DefaultPageSettings.PrinterResolution.X = 600;
            printerSettings.DefaultPageSettings.PrinterResolution.Y = 600;


            PdfViewer viewer = new PdfViewer(pdfDocument);
            viewer.PrintDocumentWithSettings(printerSettings);
            viewer.Close();

        }
    }
}