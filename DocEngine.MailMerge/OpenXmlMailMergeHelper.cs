
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using NLog;
using System.Diagnostics;
namespace DocEngine.MailMerge
{

    public static class OpenXmlMailMergeHelper
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        //const string FOLDER_PATH = @"D:\\Users\\anjum.rizwi\\Documents\\Valtech\\SAGENT-POC\\MailMergeFun\\MailMergeTestApp\";
        //const string TEMPLATE_PATH = FOLDER_PATH + @"template\";
        //const string OUTPUT_PATH = FOLDER_PATH + @"output\";

        /// <summary>
        /// Performs mail merge for multiple recipients and creates one output file per recipient.
        /// </summary>
        /// <param name="templatePath">Path to the base template .docx file</param>
        /// <param name="outputDirectory">Folder where merged files will be saved</param>
        /// <param name="recipientData">List of field dictionaries for each recipient</param>
        public static void PerformBatchMailMerge(string templatePath, string outputDirectory, List<Dictionary<string, string>> recipientData)
        {
            logger.Info("Batch MailMerge Process started...");
            var stopwatch = Stopwatch.StartNew();

            for (int i = 0; i < recipientData.Count; i++)
            {
                string fileName = $"Merged_{i + 1}_{Guid.NewGuid().ToString()}.docx"; ;
                string outputPath = Path.Combine(outputDirectory, fileName);

                File.Copy(templatePath, outputPath, true);

                try
                {
                    using (var doc = WordprocessingDocument.Open(outputPath, true))
                    {
                        var body = doc.MainDocumentPart.Document.Body;

                        foreach (var text in body.Descendants<Text>())
                        {
                            //Console.WriteLine(text.Text);

                            foreach (var pair in recipientData[i])
                            {
                                if (text.Text.Contains(pair.Key))
                                {
                                    text.Text = text.Text.Replace(pair.Key, pair.Value);
                                }
                            }
                        }

                        doc.MainDocumentPart.Document.Save();
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex);
                    logger.Error($"template:{templatePath} \n MailMerged: {outputPath}");
                }
            }
            stopwatch.Stop();
            logger.Info($"[SUCCESS]: {recipientData.Count} file mail merged in {stopwatch.Elapsed.TotalSeconds:F2} seconds");
        }

        public static void PerformMailMerge(string templatePath, string outputPath, Dictionary<string, string> placeholders)
        {
            // Copy the template to output to avoid modifying original
            File.Copy(templatePath, outputPath, true);

            using (var doc = WordprocessingDocument.Open(outputPath, true))
            {
                var body = doc.MainDocumentPart.Document.Body;

                foreach (var text in body.Descendants<Text>())
                {
                    foreach (var pair in placeholders)
                    {
                        if (text.Text.Contains("{{" + pair.Key + "}}"))
                        {
                            text.Text = text.Text.Replace("{{" + pair.Key + "}}", pair.Value);
                        }
                    }
                }

                doc.MainDocumentPart.Document.Save();
            }
        }
    }

}
