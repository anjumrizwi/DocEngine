namespace DocEngine.Processor
{
    using System;
    using System.IO;
    using PdfSharpCore.Pdf;
    using PdfSharpCore.Drawing;
    using System.Text;
    //using PdfSharp.Drawing;
    //using PdfSharp.Pdf;

    public class PdfSharpCoreHelper
    {
        public static void SaveTextAsPdf(string text, string outputPath)
        {
            // Enable code pages for Unicode fonts
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using var document = new PdfDocument();
            var page = document.AddPage();
            var gfx = XGraphics.FromPdfPage(page);

            // Use built-in font
            var font = new XFont("Verdana", 12, XFontStyle.Regular);

            gfx.DrawString(text, font, XBrushes.Black,
                           new XRect(40, 40, page.Width - 80, page.Height - 80),
                           XStringFormats.TopLeft);

            document.Save(outputPath);
        }
    }

}
