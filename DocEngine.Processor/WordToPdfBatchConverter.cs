namespace DocEngine.Processor
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Text;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using PdfSharpCore.Pdf;
    using PdfSharpCore.Drawing;


    internal class WordToPdfBatchConverter
    {
        public static void ConvertAllDocxInFolder(string inputFolder, string outputFolder)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Directory.CreateDirectory(outputFolder);

            var docxFiles = Directory.GetFiles(inputFolder, "*.docx");

            foreach (var file in docxFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(file);
                string outputPdf = Path.Combine(outputFolder, fileName + ".pdf");

                try
                {
                    string text = ExtractTextFromDocx(file);
                    SaveTextAsPdf(text, outputPdf);
                    Console.WriteLine($"✅ Converted: {fileName}.pdf");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error converting {fileName}.docx: {ex.Message}");
                }
            }
        }

        private static string ExtractTextFromDocx(string path)
        {
            using var wordDoc = WordprocessingDocument.Open(path, false);
            var body = wordDoc.MainDocumentPart.Document.Body;

            return body.Descendants<Text>()
                       .Select(t => t.Text)
                       .Aggregate((a, b) => a + " " + b);
        }

        private static void SaveTextAsPdf(string text, string outputPath)
        {
            using var doc = new PdfDocument();
            var page = doc.AddPage();
            var gfx = XGraphics.FromPdfPage(page);
            var font = new XFont("Verdana", 12, XFontStyle.Regular);

            gfx.DrawString(text, font, XBrushes.Black,
                new XRect(40, 40, page.Width - 80, page.Height - 80),
                XStringFormats.TopLeft);

            doc.Save(outputPath);
        }
    }
}
