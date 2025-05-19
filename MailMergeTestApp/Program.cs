using System;
using System.Net.Http;
using System.Text;
using DocxToPdf;
using Sagent.XceedUtils;
using OpenXmlMailMerge;

namespace MailMergeTestApp
{
    class Program
    {
        const string BASE_PATH = @"D:\\Users\\anjum.rizwi\\source\repos\\anjumrizwi\\MailMerge\\MailMergeTestApp\\";
        const string TEMPLATE_PATH = BASE_PATH + @"template\\";
        const string DOCX_PATH = BASE_PATH + @"docx\\";
        const string REPORT_PATH = BASE_PATH + @"pdf\\";
        const string PRN_PATH = BASE_PATH + @"prn\\";
        const string PRN_TEST_PATH = BASE_PATH + @"test\\";

        static async Task Main(string[] args)
        {
            //Option1: Licensed: Usning Xceed.Words.NET;
            //await MailMergeLetter();
            // Console.WriteLine("Xceed.Words.NET: Mail merge completed. Press any key to exit.");

            //Option2: Using DocumentFormat.OpenXml SDK
            //PerformBatchMailMerge();
            //XceedHelper.Convert();
            // MailMergeHelper.ExecuteDataTable(DOCX_PATH);

            ///DOCS to PDF
            //await Task.Run(() => ConvertDocxToPdf());
            //await Task.Run(() => WordToPdfBatchConverter());

            //await Task.Run(() => WordToPdfParagraphPreserving());

            //Option3: Using Azure Function

            //await ExecuteMailMergeFunction();
            //Console.WriteLine("Azure Function: Mail merge completed. Press any key to exit.");  

            //Option3: Licensed
            AsposeHelper.Convert();
        }

        private static void PerformBatchMailMerge()
        {
            var docFilePath = TEMPLATE_PATH + "Business-Plan-Word-Template.docx";
            OpenXmlMailMergeHelper.PerformBatchMailMerge(docFilePath, DOCX_PATH, DataHelper.GetRecepientData());
        }

        

        private static async Task ExecuteMailMergeFunction()
        {
            var client = new HttpClient();
            var url = "https://<your-function-app>.azurewebsites.net/api/MailMergeFunction"; // Replace with your actual URL

            var json = @"{
                            ""FirstName"": ""John"",
                            ""LastName"": ""Doe"",
                            ""Address"": ""123 Test Street"",
                            ""Email"": ""john.doe@example.com""
                        }";


            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await client.PostAsync(url, content);

            var responseContent = await response.Content.ReadAsStringAsync();
            Console.WriteLine($"Status: {response.StatusCode}");
            Console.WriteLine($"Response: {responseContent}");
        }

        


        private static void ConvertDocxToPdf()
        {
            string inputFolder = DOCX_PATH;
            string outputFolder = REPORT_PATH;

            DocxToPdfConverter.ConvertAllInFolder(inputFolder, outputFolder);

            Console.WriteLine("All conversions completed.");
        }

        private static void WordToPdfBatchConverter()
        {
            string inputFolder = DOCX_PATH;
            string outputFolder = REPORT_PATH;

            DocxToPdf.WordToPdfBatchConverter.ConvertAllDocxInFolder(inputFolder, outputFolder);

            Console.WriteLine("All bach conversions completed.");
        }

        private static void WordToPdfParagraphPreserving()
        {
            string inputFolder = DOCX_PATH;
            string outputFolder = REPORT_PATH;

            DocxToPdf.WordToPdfParagraphPreserving.ConvertAllDocxInFolder(inputFolder, outputFolder);

            Console.WriteLine("All bach conversions completed.");
        }

        
    }
}
