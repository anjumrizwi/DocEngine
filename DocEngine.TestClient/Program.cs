using DocEngine.MailMerge;
using DocEngine.Processor;
using System.Text;
using System.IO;


namespace DocEngine.TestClient
{
    class Program
    {
        static readonly string baseFolderPath = Directory.GetParent(AppContext.BaseDirectory)?.Parent?.Parent?.Parent?.FullName ?? string.Empty;

        public static void Main(string[] args)
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

            DocHelper.Convert(baseFolderPath);

            //Option3: Licensed
            //AsposeHelper.Convert();
            //return Task.CompletedTask;
        }

       
    }
}
