using AsposeWords;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailMergeTestApp
{
    internal class AsposeHelper
    {
        const string BASE_PATH = @"D:\\Users\\anjum.rizwi\\source\repos\\anjumrizwi\\MailMerge\\MailMergeTestApp\\";
        const string TEMPLATE_PATH = BASE_PATH + @"template\\";
        const string DOCX_PATH = BASE_PATH + @"docx\\";
        const string REPORT_PATH = BASE_PATH + @"pdf\\";
        const string PRN_PATH = BASE_PATH + @"prn\\";
        public static void Convert()
        {
            var templateFilePath = TEMPLATE_PATH + "Business-Plan-Word-Template.docx";
            AsposeMailMergeHelper.ExecuteDataTable(templateFilePath, DOCX_PATH, DataHelper.GetRecepientData());

            ConvertApposeDocxToPdf(DOCX_PATH, REPORT_PATH);
            ConvertApposePdfToPrn(REPORT_PATH, PRN_PATH);
        }

        private static string GetAsposeLicense()
        {
            // This is a placeholder for the actual license key.
            // In a real application, you would retrieve this from a secure location.
            return "YOUR_LICENSE_KEY_HERE";
        }


        private static void SetAsposeLicense()
        {
            // This is a placeholder for setting the license.
            // In a real application, you would use the Aspose API to set the license.
            string licenseKey = GetAsposeLicense();
            if (string.IsNullOrEmpty(licenseKey))
            {
                throw new InvalidOperationException("License key is not set.");
            }
            // Example: Aspose.Words.License.SetLicense(licenseKey);
        }
        private static void InitializeAspose()
        {
            // This method initializes the Aspose library.
            // In a real application, you would call the appropriate initialization methods here.
            SetAsposeLicense();
            // Example: Aspose.Words.License.SetLicense("path/to/license/file");
        }
        
        private static void ConvertDocxToPdf(string inputFilePath, string outputFilePath)
        {
            // This is a placeholder for the actual conversion logic.
            // In a real application, you would use the Aspose API to convert the file.
            // Example: Aspose.Words.Document doc = new Aspose.Words.Document(inputFilePath);
            // doc.Save(outputFilePath, Aspose.Words.SaveFormat.Pdf);
        }
        private static void ConvertPdfToDocx(string inputFilePath, string outputFilePath)
        {
            // This is a placeholder for the actual conversion logic.
            // In a real application, you would use the Aspose API to convert the file.
            // Example: Aspose.Words.Document doc = new Aspose.Words.Document(inputFilePath);
            // doc.Save(outputFilePath, Aspose.Words.SaveFormat.Docx);
        }

        private static void ConvertDocxToPrn(string inputFilePath, string outputFilePath)
        {
            // This is a placeholder for the actual conversion logic.
            // In a real application, you would use the Aspose API to convert the file.
            // Example: Aspose.Words.Document doc = new Aspose.Words.Document(inputFilePath);
            // doc.Save(outputFilePath, Aspose.Words.SaveFormat.Prn);
        }

        private static void ConvertPrnToDocx(string inputFilePath, string outputFilePath)
        {
            // This is a placeholder for the actual conversion logic.
            // In a real application, you would use the Aspose API to convert the file.
            // Example: Aspose.Words.Document doc = new Aspose.Words.Document(inputFilePath);
            // doc.Save(outputFilePath, Aspose.Words.SaveFormat.Docx);
        }

        private static void ConvertPdfToPrn(string inputFilePath, string outputFilePath)
        {
            // This is a placeholder for the actual conversion logic.
            // In a real application, you would use the Aspose API to convert the file.
            // Example: Aspose.Words.Document doc = new Aspose.Words.Document(inputFilePath);
            // doc.Save(outputFilePath, Aspose.Words.SaveFormat.Prn);
        }

        private static void ConvertApposeDocxToPdf(string docxFilePath, string pdfOutputFilePath)
        {
            string inputFile = docxFilePath;
            string outputFile = pdfOutputFilePath;

            AsposeWords.AsposeDocxToPdfConverter.ConvertAllDocToPdfInFolder(inputFile, outputFile);
        }

        private static void ConvertApposePdfToPrn(string pdfFilePath, string prnOutputFilePath)
        {
            string inputFile = pdfFilePath;
            string outputFile = prnOutputFilePath;
            AsposeWords.AsposePdfToPrnConverter.ConvertAllPdfToPrnInFolder(inputFile, outputFile);
        }

        private static void ConvertApposePrnToPdf(string prnFilePath, string prn_pdfOutputFilePath)
        {
            string inputFile = prnFilePath;
            string outputFile = prn_pdfOutputFilePath;

            AsposePrnToPdfConverter.ConvertAllPrnToPdfInFolder(inputFile, outputFile);
        }

    }
}
