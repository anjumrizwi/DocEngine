using Aspose.Cells;
namespace DocEngine.AsposeWords
{
    public class AsposePrnToPdfConverter
    {

        /// <summary>
        /// Convert PRN to PDF Folder
        /// </summary>
        public static void ConvertAllPrnToPdfInFolder(string inputFolder, string outputFolder)
        {
            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var pdfFiles = Directory.GetFiles(inputFolder, "*.prn");

            foreach (var file in pdfFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                string outputFile = Path.Combine(outputFolder, fileName + ".pdf");

                try
                {
                    Console.WriteLine($"Converting: {fileName}.prn → {fileName}.pdf");
                    var workbook = new Workbook(file);
                    workbook.Save("OutputFromPrn.pdf");

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error converting {fileName}: {ex.Message}");
                }
            }

            Console.WriteLine("✅ PRN->Pdf Batch conversion complete.");
        }
    }
}