using DocEngine.MailMerge;
using DocEngine.Data;

namespace DocEngine.Processor
{
    public class DocHelper
    {
        public static void Convert(string basePath)
        {
            //MailMerge-Option1: Required: License to use NGS.Templater;
            //MailMergeUsingTemplater();

            //MailMerge-Option2: DocumentFormat.OpenXml: MIT license
            PerformBatchMailMerge(basePath, 5);

            //Convert: Docx to Pdf
            ConvertDocxToPdf(basePath); //Image size issue
            //ConvertBatchDocxToPdf(basePath); //Single line output pdf file
            //PreserveWordToPdfParagraph(basePath); //Image missing
            //ConvertDocxToPdfWithImage(basePath); //Only Image coming

            //Convert: Pdf to Prn
            ConvertPdfToPrn(basePath);
            Console.WriteLine("All conversions completed.");
        }

        private static void PerformBatchMailMerge(string basePath, int recipientCount)
        {
            var templatePath = Path.Combine(basePath, "template", "Business-Plan-Word-Template.docx");
            var outputFilePath = Path.Combine(basePath, "docx");

            OpenXmlMailMergeHelper.PerformBatchMailMerge(templatePath, outputFilePath, DataHelper.GetRecepientData(recipientCount));
        }

        private static void MailMergeUsingTemplater()
        {
            //var templatePath = Path.Combine(basePath, "template", "Business-Plan-Word-Template.docx");
            //var outputFilePath = Path.Combine(basePath, "docx");

            TemplaterMailMerge.MailMerge();
        }

        private static void ConvertDocxToPdf(string basePath)
        {
            string inputFolder = Path.Combine(basePath, "docx");
            string outputFolder = Path.Combine(basePath, "pdf");
            string archiveFolder = Path.Combine(basePath, "archieve");

            (new DocxToPdfConverter()).ConvertAllInFolder(inputFolder, outputFolder, archiveFolder);
        }

        private static void ConvertBatchDocxToPdf(string basePath)
        {
            string inputFolder = Path.Combine(basePath, "docx");
            string outputFolder = Path.Combine(basePath, "pdf");
            string archiveFolder = Path.Combine(basePath, "archieve");

            WordToPdfBatchConverter.ConvertAllDocxInFolder(inputFolder, outputFolder);
        }

        private static void PreserveWordToPdfParagraph(string basePath)
        {
            string inputFolder = Path.Combine(basePath, "docx");
            string outputFolder = Path.Combine(basePath, "pdf");
            string archiveFolder = Path.Combine(basePath, "archieve");

            WordToPdfParagraphPreserving.ConvertAllDocxInFolder(inputFolder, outputFolder);
        }

        private static void ConvertDocxToPdfWithImage(string basePath)
        {
            string inputFolder = Path.Combine(basePath, "docx");
            string outputFolder = Path.Combine(basePath, "pdf");
            string archiveFolder = Path.Combine(basePath, "archieve");

            DocxToPdfWithImageConverter.ConvertAllDocxInFolder(inputFolder, outputFolder, archiveFolder);
        }

        private static void ConvertPdfToPrn(string basePath)
        {
            string inputFolder = Path.Combine(basePath, "pdf");
            string outputFolder = Path.Combine(basePath, "prn");
            string archiveFolder = Path.Combine(basePath, "archieve");

            PdfToPrnConverter.ConvertAllInFolder(inputFolder, outputFolder, archiveFolder);
        }
    }
}
