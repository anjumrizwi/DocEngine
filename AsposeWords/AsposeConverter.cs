//using System.Reflection.Metadata;
using System.IO;
using Aspose.Words;
namespace AsposeWords
{
    public class AsposeConverter
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
                    Convert(file, outputFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error converting {fileName}: {ex.Message}");
                }
            }

            Console.WriteLine("✅Pdf Batch conversion complete.");
        }

        public static void ConvertAllPdfToPrnInFolder(string inputFolder, string outputFolder)
        {
            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var docxFiles = Directory.GetFiles(inputFolder, "*.pdf");

            foreach (var file in docxFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                string outputFile = Path.Combine(outputFolder, fileName + ".prn");

                try
                {
                    Console.WriteLine($"Converting: {fileName}.pdf → {fileName}.prn");
                    Convert(file, outputFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error converting {fileName}: {ex.Message}");
                }
            }

            Console.WriteLine("✅ PRN Batch conversion complete.");
        }

        public static void Convert(string sourceFile, string pdfFile)
        {
            // Load the document from disk
            // Note: Ensure you have the Aspose.Words library installed and referenced in your project.
            // You can install it via NuGet: Install-Package Aspose.Words
            // Example: var doc = new Document("input.docx");
            var doc = new Document(sourceFile);

            // Save as DOCX
            doc.Save(sourceFile, SaveFormat.Docx);

            // Save as PDF
            doc.Save(pdfFile, SaveFormat.Pdf);

            // Save as PRN (PlainText to mimic PRN)
            doc.Save(pdfFile + ".prn", SaveFormat.Text);

        }

        public static void CompareDoc()
        {
            var doc1 = new Document("file1.docx");
            var doc2 = new Document("file2.docx");

            doc1.Compare(doc2, "Author", DateTime.Now);
            doc1.Save("diff.docx");

        }
    }
}