using System;
using System.Net.Http;
using System.Text;
using DocxToPdf;
using MailMerge.Business;
using OpenXmlMailMerge;

namespace MailMergeTestApp
{
    class Program
    {
        const string FOLDER_PATH = @"D:\\Users\\anjum.rizwi\\Documents\\Valtech\\SAGENT-POC\\MailMergeFun\\MailMergeTestApp\";
        const string TEMPLATE_PATH = FOLDER_PATH + @"template\";
        const string DOCX_PATH = FOLDER_PATH + @"output\";
        const string REPORT_PATH = FOLDER_PATH + @"report\";

        static async Task Main(string[] args)
        {
            //Option1: Usning Xceed.Words.NET;
            //await MailMergeLetter();
            // Console.WriteLine("Xceed.Words.NET: Mail merge completed. Press any key to exit.");

            //Option2: Using DocumentFormat.OpenXml SDK
            PerformBatchMailMerge();

            ///DOCS to PDF
            await Task.Run(() => ConvertDocxToPdf());
            //await Task.Run(() => WordToPdfBatchConverter());

            //await Task.Run(() => WordToPdfParagraphPreserving());

            //Option3: Using Azure Function

            //await ExecuteMailMergeFunction();
            //Console.WriteLine("Azure Function: Mail merge completed. Press any key to exit.");  
        }

        private static void PerformBatchMailMerge()
        {
            OpenXmlMailMergeHelper.PerformBatchMailMerge("file-sample_500kB.docx", "output", new List<Dictionary<string, string>>()
            {
                new Dictionary<string, string>
                {
                    { "Name", "Anjum Rizwi" },
                    { "Date", DateTime.Today.ToShortDateString() },
                    { "Address", "382, 6th Cross, SK Garden, Bengaluru" },
                    { "Company Name", "Valtech India" }
                },
                new Dictionary<string, string>
                {
                    { "Name", "Arun" },
                    { "Date", DateTime.Today.ToShortDateString() },
                    { "Address", "382, 6th Cross, JP Nagar, Bengaluru" },
                    { "Company Name", "Valtech India" }
                },
                new Dictionary<string, string>
                {
                    { "Name", "John Doe" },
                    { "Date", DateTime.Today.ToShortDateString() },
                    { "Address", "123 Test Street, Bengaluru" },
                    { "Company Name", "Valtech US" }
                }
            });
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

        private static async Task MailMergeLetter()
        {
            var templatePath = "MailMerge_Letter_Template.docx";
            var outputPath = "Merged_Letter.docx";

            var mergeFields = new Dictionary<string, string>
            {
                { "Name", "Anjum Rizwi" },
                { "Date", DateTime.Today.ToShortDateString() },
                { "Address", "382, 6th Cross, SK Garden, Bengaluru" }
            };

            await XceedMailMergeHelper.PerformMailMergeAsync(templatePath, outputPath, mergeFields);

            Console.WriteLine("Usage: MailMergeTestApp.exe <function-url>");
            Console.WriteLine("Example: MailMergeTestApp.exe https://<your-function-app>.azurewebsites.net/api/MailMergeFunction");
        }


        private static void ConvertDocxToPdf()
        {
            string inputFolder = DOCX_PATH;
            string outputFolder = REPORT_PATH;

            DocxToPdfConverter.ConvertAllDocxInFolder(inputFolder, outputFolder);

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
