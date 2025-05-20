using DocEngine.MailMerge;
using DocEngine.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocEngine.Processor
{
    public class DocHelper
    {
        public static void Convert(string basePath)
        {
            PerformBatchMailMerge (basePath);
            //ConvertDocxToPdf(basePath);
            //ConvertBatchDocxToPdf(basePath);
            PreserveWordToPdfParagraph(basePath);
        }

        private static void PerformBatchMailMerge(string basePath)
        {
            var templatePath = Path.Combine(basePath, "template","Business-Plan-Word-Template.docx");
            var outputFilePath = Path.Combine(basePath, "docx");

            OpenXmlMailMergeHelper.PerformBatchMailMerge(templatePath, outputFilePath, DataHelper.GetRecepientData());
        }

        private static void ConvertDocxToPdf(string basePath)
        {       
            string inputFolder = Path.Combine(basePath, "docx");
            string outputFolder = Path.Combine(basePath, "pdf");

            DocxToPdfConverter.ConvertAllInFolder(inputFolder, outputFolder);

            Console.WriteLine("All conversions completed.");
        }

        private static void ConvertBatchDocxToPdf(string basePath)
        {
            string inputFolder = Path.Combine(basePath, "docx");
            string outputFolder = Path.Combine(basePath, "pdf");

            WordToPdfBatchConverter.ConvertAllDocxInFolder(inputFolder, outputFolder);

            Console.WriteLine("All bach conversions completed.");
        }

        private static void PreserveWordToPdfParagraph(string basePath)
        {
            string inputFolder = Path.Combine(basePath, "docx");
            string outputFolder = Path.Combine(basePath, "pdf");

            WordToPdfParagraphPreserving.ConvertAllDocxInFolder(inputFolder, outputFolder);

            Console.WriteLine("All bach conversions completed.");
        }
    }
}
