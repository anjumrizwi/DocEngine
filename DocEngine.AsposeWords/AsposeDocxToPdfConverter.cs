
using System.Diagnostics;
using Aspose.Words;
namespace DocEngine.AsposeWords
{
    /// <summary>
    /// Docs to PDF converter
    /// </summary>
    public class AsposeDocxToPdfConverter
    {
        public static void ConvertAllDocToPdfInFolder(string inputFolder, string outputFolder)
        {
            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var docxFiles = Directory.GetFiles(inputFolder, "*.docx");

            foreach (var file in docxFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                string outputFile = Path.Combine(outputFolder, fileName + ".pdf");

                try
                {
                    Console.WriteLine($"Converting: {fileName}.docx → {fileName}.pdf");
                    var stopwatch = Stopwatch.StartNew();
                    ConvertToPdf(file, outputFile);
                    stopwatch.Stop();
                    Console.WriteLine($"[SUCCESS] {fileName} converted in {stopwatch.Elapsed.TotalSeconds:F2} seconds");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error converting {fileName}: {ex.Message}");
                }
            }

            Console.WriteLine("✅Pdf Batch conversion complete.");
        }


        public static void ConvertToPdf(string sourceFile, string pdfFile)
        {
            var doc = new Document(sourceFile);

            // Save as PDF
            doc.Save(pdfFile, SaveFormat.Pdf);
        }
    }
}