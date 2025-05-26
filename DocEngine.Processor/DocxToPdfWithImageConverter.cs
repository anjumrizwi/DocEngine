namespace DocEngine.Processor
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Vml;
    using A = DocumentFormat.OpenXml.Drawing;
    using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
    using PdfSharpCore.Pdf;
    using PdfSharpCore.Drawing;
    using SixLabors.ImageSharp;
    using SixLabors.ImageSharp.Metadata;
    using NLog;
    using System.Diagnostics;

    internal static class DocxToPdfWithImageConverter
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public static void ConvertAllDocxInFolder(string inputFolder, string outputFolder, string archiveFolder)
        {
            #region Directory Check
            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            if (!Directory.Exists(archiveFolder))
                Directory.CreateDirectory(archiveFolder);

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder); 
            #endregion

            logger.Info("Batch DocxToPdf Process started...");
            var stopwatch = Stopwatch.StartNew();

            var docxFiles = Directory.GetFiles(inputFolder, "*.docx");
            var errCnt = 0;
            foreach (var file in docxFiles)
            {
                string fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                string outputFile = System.IO.Path.Combine(outputFolder, fileName + ".pdf");
                string archiveFile = System.IO.Path.Combine(archiveFolder, System.IO.Path.GetFileName(file));
                try
                {
                    ConvertDocxToPdf(file, outputFile);
                    File.Move(file, archiveFile);
                    logger.Info($"Moved {fileName}.docx to archive folder: {archiveFile}");
                }
                catch (Exception ex)
                {
                    errCnt++;
                    logger.Error($"Error converting {fileName}: {ex.Message}");
                }
            }

            stopwatch.Stop();
            if (errCnt > 0)
                logger.Error($"[ERROR]: {errCnt} docx files failed to convert to pdf.");

            logger.Info($"[SUCCESS]: {docxFiles.Length - errCnt} docx files converted to pdf in {stopwatch.Elapsed.TotalSeconds:F2} seconds");
            logger.Info("Docx2PdfWithImage Batch conversion complete.");
        }

        public static void ConvertDocxToPdf(string docxPath, string pdfPath)
        {
            List<(string Id, Stream Stream, string Name, long WidthEmus, long HeightEmus)> images = null;
            List<string> paragraphs = null;
            try
            {
                logger.Info("Starting DOCX to PDF conversion.");
                images = ExtractImages(docxPath);
                paragraphs = ExtractText(docxPath);
                using (var pdfDoc = new PdfDocument())
                {
                    var page = pdfDoc.AddPage();
                    var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Prepend);

                    // Draw background image first (if any)
                    var backgroundImage = images.Find(img => img.Name.Contains("Background"));
                    if (backgroundImage.Stream != null)
                    {
                        DrawImageToPdf(gfx, backgroundImage.Stream, backgroundImage.Name, backgroundImage.WidthEmus, backgroundImage.HeightEmus, 0, 0, page.Width, page.Height);
                        logger.Info("Background image drawn in PDF.");
                    }

                    // Draw other images (logo, graph, etc.)
                    int yOffset = 50;
                    foreach (var image in images)
                    {
                        if (!image.Name.Contains("Background"))
                        {
                            DrawImageToPdf(gfx, image.Stream, image.Name, image.WidthEmus, image.HeightEmus, 50, yOffset);
                            yOffset += (int)(image.HeightEmus / 914400.0 * 72) + 20; // Increment Y position
                        }
                    }

                    // Draw text
                    var font = new XFont("Arial", 12, XFontStyle.Regular); // Fixed: Removed 'using' since XFont is not IDisposable
                    foreach (var paragraph in paragraphs)
                    {
                        gfx.DrawString(paragraph, font, XBrushes.Black, new XPoint(50, yOffset), XStringFormats.TopLeft);
                        yOffset += 20; // Simple line spacing
                        logger.Info($"Drew text: {paragraph.Substring(0, Math.Min(paragraph.Length, 50))}...");
                    }

                    pdfDoc.Save(pdfPath);
                    logger.Info($"PDF saved to {pdfPath}");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to convert DOCX to PDF.");
                throw;
            }
            finally
            {
                // Dispose streams
                if (images != null)
                {
                    foreach (var image in images)
                    {
                        image.Stream?.Dispose();
                    }
                }
            }
        }

        private static List<(string Id, Stream Stream, string Name, long WidthEmus, long HeightEmus)> ExtractImages(string docxPath)
        {
            var images = new List<(string, Stream, string, long, long)>();
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, false))
                {
                    // Extract images from main document
                    foreach (var drawing in doc.MainDocumentPart.Document.Body.Descendants<Drawing>())
                    {
                        var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
                        if (blip?.Embed?.Value != null)
                        {
                            var imagePart = (ImagePart)doc.MainDocumentPart.GetPartById(blip.Embed.Value);
                            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
                            var name = drawing.Descendants<DW.DocProperties>().FirstOrDefault()?.Name?.Value ?? "image";
                            var stream = new MemoryStream();
                            using (var sourceStream = imagePart.GetStream())
                            {
                                sourceStream.CopyTo(stream);
                                stream.Position = 0;
                            }
                            images.Add((blip.Embed.Value, stream, name, extent?.Cx ?? 990000L, extent?.Cy ?? 792000L));
                            logger.Info($"Extracted image {name} from main document with size Cx={extent?.Cx}, Cy={extent?.Cy}");
                        }
                    }

                    // Extract background images from headers/footers
                    foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
                    {
                        foreach (var imageData in headerPart.Header.Descendants<ImageData>())
                        {
                            if (imageData?.RelationshipId?.Value != null)
                            {
                                var imagePart = (ImagePart)headerPart.GetPartById(imageData.RelationshipId.Value);
                                var shape = imageData.Parent as Shape;
                                var name = "BackgroundImage";
                                var style = shape?.Style?.Value ?? "";
                                long widthEmus = 990000L, heightEmus = 792000L; // Default size
                                if (style.Contains("width") && style.Contains("height"))
                                {
                                    try
                                    {
                                        var widthStr = style.Substring(style.IndexOf("width:") + 6).Split(';')[0].Replace("in", "");
                                        var heightStr = style.Substring(style.IndexOf("height:") + 7).Split(';')[0].Replace("in", "");
                                        widthEmus = (long)(double.Parse(widthStr) * 914400);
                                        heightEmus = (long)(double.Parse(heightStr) * 914400);
                                    }
                                    catch
                                    {
                                        logger.Warn("Failed to parse VML shape size for background image.");
                                    }
                                }
                                var stream = new MemoryStream();
                                using (var sourceStream = imagePart.GetStream())
                                {
                                    sourceStream.CopyTo(stream);
                                    stream.Position = 0;
                                }
                                images.Add((imageData.RelationshipId.Value, stream, name, widthEmus, heightEmus));
                                logger.Info($"Extracted background image from header with size Cx={widthEmus}, Cy={heightEmus}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to extract images from DOCX.");
            }
            return images;
        }

        private static List<string> ExtractText(string docxPath)
        {
            var paragraphs = new List<string>();
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, false))
                {
                    foreach (var paragraph in doc.MainDocumentPart.Document.Body.Descendants<Paragraph>())
                    {
                        var text = string.Join("", paragraph.Descendants<Text>().Select(t => t.Text));
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            paragraphs.Add(text);
                            logger.Info($"Extracted paragraph: {text.Substring(0, Math.Min(text.Length, 50))}...");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to extract text from DOCX.");
            }
            return paragraphs;
        }

        private static void DrawImageToPdf(XGraphics gfx, Stream imageStream, string name, long widthEmus, long heightEmus, double x, double y, double? targetWidth = null, double? targetHeight = null)
        {
            try
            {
                // Convert EMUs to points (1 EMU = 1/914400 inches, 1 inch = 72 points)
                double widthPoints = targetWidth ?? (widthEmus / 914400.0 * 72);
                double heightPoints = targetHeight ?? (heightEmus / 914400.0 * 72);

                // Save image to temporary file for PdfSharpCore compatibility
                string tempImagePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"{Guid.NewGuid()}.png");
                using (var img = Image.Load(imageStream))
                {
                    img.Save(tempImagePath, new SixLabors.ImageSharp.Formats.Png.PngEncoder());
                }

                var xImage = XImage.FromFile(tempImagePath);
                gfx.DrawImage(xImage, x, y, widthPoints, heightPoints);
                logger.Info($"Drew image {name} in PDF at position ({x},{y}) with size {widthPoints}x{heightPoints} points.");

                File.Delete(tempImagePath);
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to draw image {name} in PDF.");
            }
        }
    }
}