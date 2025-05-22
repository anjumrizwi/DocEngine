namespace DocEngine.Processor
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Drawing;
    using PdfSharpCore.Drawing;
    using PdfSharpCore.Pdf;
    using System.Text;

    // Aliases to avoid conflicts
    using Wp = DocumentFormat.OpenXml.Wordprocessing;
    using IOPath = System.IO.Path;

    internal static class DocxToPdfConverter
    {
        public static void ConvertAllInFolder(string inputFolder, string outputFolder)
        {
            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            var docxFiles = Directory.GetFiles(inputFolder, "*.docx");

            foreach (var file in docxFiles)
            {
                string fileName = IOPath.GetFileNameWithoutExtension(file);
                string outputFile = IOPath.Combine(outputFolder, fileName + ".pdf");

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

            Console.WriteLine("✅ Batch conversion complete.");
        }

        public static void Convert(string docxPath, string outputPdfPath)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using var wordDoc = WordprocessingDocument.Open(docxPath, false);
            var body = wordDoc.MainDocumentPart.Document.Body;
            var mainPart = wordDoc.MainDocumentPart;

            using var document = new PdfDocument();
            var page = document.AddPage();
            var gfx = XGraphics.FromPdfPage(page);

            double margin = 40, y = margin, pageWidth = page.Width, pageHeight = page.Height;

            foreach (var element in body.Elements())
            {
                if (y > pageHeight - margin)
                {
                    page = document.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    y = margin;
                }

                switch (element)
                {
                    case Wp.Paragraph para:
                        foreach (var run in para.Elements<Wp.Run>())
                        {
                            if (run.Descendants<Wp.Break>().Any(b => b.Type?.Value == Wp.BreakValues.Page))
                            {
                                page = document.AddPage();
                                gfx = XGraphics.FromPdfPage(page);
                                y = margin;
                                continue;
                            }

                            if (run.Descendants<Blip>().Any())
                            {
                                DrawInlineImage(run, gfx, mainPart, ref y);
                            }
                            else
                            {
                                var text = run.InnerText.Trim();
                                if (string.IsNullOrEmpty(text)) continue;
                                var font = new XFont("Verdana", 12, GetFontStyle(run));
                                gfx.DrawString(text, font, XBrushes.Black, new XPoint(margin, y));
                                y += 20;
                            }
                        }
                        break;

                    case Wp.Table table:
                        DrawAdvancedTable(gfx, table, ref y, pageHeight, pageWidth, margin, document);
                        break;
                }
            }

            document.Save(outputPdfPath);
        }

        private static void DrawInlineImage(Wp.Run run, XGraphics gfx, MainDocumentPart mainPart, ref double y)
        {
            var blip = run.Descendants<Blip>().FirstOrDefault();
            if (blip?.Embed == null) return;

            var part = (ImagePart)mainPart.GetPartById(blip.Embed.Value);
            using var stream = part.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);

            var img = XImage.FromStream(() => new MemoryStream(ms.ToArray()));
            gfx.DrawImage(img, 40, y, 100, 100);
            y += 110;
        }

        private static void DrawAdvancedTable(XGraphics gfx, Wp.Table table, ref double y, double pageHeight, double pageWidth, double margin, PdfDocument document)
        {
            var rows = table.Elements<Wp.TableRow>().ToList();
            if (!rows.Any()) return;

            var columnCount = rows[0].Elements<Wp.TableCell>().Count();
            var columnWidths = CalculateColumnWidths(gfx, rows, columnCount, pageWidth, margin);
            var font = new XFont("Verdana", 10);

            foreach (var row in rows)
            {
                double x = margin, cellHeight = 25;

                if (y + cellHeight > pageHeight - margin)
                {
                    var page = document.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    y = margin;
                }

                var cells = row.Elements<Wp.TableCell>().ToList();
                for (int i = 0; i < columnCount && i < cells.Count; i++)
                {
                    string text = string.Join(" ", cells[i].Descendants<Wp.Text>().Select(t => t.Text));
                    gfx.DrawRectangle(XPens.Black, x, y, columnWidths[i], cellHeight);
                    gfx.DrawString(text, font, XBrushes.Black,
                        new XRect(x + 2, y + 2, columnWidths[i] - 4, cellHeight - 4), XStringFormats.TopLeft);
                    x += columnWidths[i];
                }

                y += cellHeight;
            }

            y += 10;
        }

        private static List<double> CalculateColumnWidths(XGraphics gfx, List<Wp.TableRow> rows, int columnCount, double pageWidth, double margin)
        {
            var font = new XFont("Verdana", 10);
            double availableWidth = pageWidth - margin * 2;
            var maxWidths = new double[columnCount];

            foreach (var row in rows)
            {
                var cells = row.Elements<Wp.TableCell>().ToList();
                for (int i = 0; i < columnCount && i < cells.Count; i++)
                {
                    string text = string.Join(" ", cells[i].Descendants<Wp.Text>().Select(t => t.Text));
                    double width = gfx.MeasureString(text, font).Width + 10;
                    if (width > maxWidths[i])
                        maxWidths[i] = width;
                }
            }

            double total = maxWidths.Sum();
            if (total > availableWidth)
            {
                double scale = availableWidth / total;
                for (int i = 0; i < maxWidths.Length; i++)
                    maxWidths[i] *= scale;
            }

            return maxWidths.ToList();
        }

        private static XFontStyle GetFontStyle(Wp.Run run)
        {
            var props = run.RunProperties;
            if (props == null) return XFontStyle.Regular;

            bool bold = props.Bold != null;
            bool italic = props.Italic != null;

            return (bold, italic) switch
            {
                (true, true) => XFontStyle.BoldItalic,
                (true, false) => XFontStyle.Bold,
                (false, true) => XFontStyle.Italic,
                _ => XFontStyle.Regular
            };
        }
    }


}