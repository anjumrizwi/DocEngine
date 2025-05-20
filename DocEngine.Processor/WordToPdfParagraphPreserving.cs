
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

    public class WordToPdfParagraphPreserving
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
                    var paragraphs = ExtractParagraphsFromDocx(file);
                    SaveParagraphsAsPdf(paragraphs, outputPdf);
                    Console.WriteLine($"✅ Converted: {fileName}.pdf");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Error: {fileName}.docx - {ex.Message}");
                }
            }
        }

        private static string[] ExtractParagraphsFromDocx(string filePath)
        {
            using var wordDoc = WordprocessingDocument.Open(filePath, false);
            var body = wordDoc.MainDocumentPart.Document.Body;

            return body.Elements<Paragraph>()
                       .Select(p => string.Join("",
                            p.Descendants<Text>()
                             .Select(t => t.Text)))
                       .Where(text => !string.IsNullOrWhiteSpace(text))
                       .ToArray();
        }

        private static void SaveParagraphsAsPdf(string[] paragraphs, string outputPath)
        {
            using var document = new PdfDocument();
            var page = document.AddPage();
            var gfx = XGraphics.FromPdfPage(page);
            var font = new XFont("Verdana", 12, XFontStyle.Regular);

            double margin = 40;
            double lineHeight = font.GetHeight();
            double y = margin;

            foreach (var paragraph in paragraphs)
            {
                if (y + lineHeight > page.Height - margin)
                {
                    page = document.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    y = margin;
                }

                gfx.DrawString(paragraph, font, XBrushes.Black,
                    new XRect(margin, y, page.Width - 2 * margin, lineHeight),
                    XStringFormats.TopLeft);

                y += lineHeight + 5; // Add spacing between paragraphs
            }

            document.Save(outputPath);
        }
    }

}
