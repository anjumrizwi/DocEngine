namespace MailMerge.Business
{
    using System.Collections.Generic;
    using System.IO;
    using Xceed.Words.NET;

    public static class XceedMailMergeHelper
    {
        const string FOLDER_PATH = @"D:\\Users\\anjum.rizwi\\Documents\\Valtech\\SAGENT-POC\\MailMergeFun\\MailMergeTestApp\";
        const string TEMPLATE_PATH = FOLDER_PATH + @"template\";
        const string OUTPUT_PATH = FOLDER_PATH + @"output\";
        /// <summary>
        /// Performs simple mail merge by replacing {{placeholders}} in a Word document.
        /// </summary>
        /// <param name="templatePath">Path to the .docx template</param>
        /// <param name="outputPath">Path where the merged file should be saved</param>
        /// <param name="placeholders">Dictionary of placeholders and values</param>
        public static async Task PerformMailMergeAsync(string templatePath, string outputPath, Dictionary<string, string> placeholders)
        {
            // Load file asynchronously into memory
            byte[] templateBytes = await File.ReadAllBytesAsync(TEMPLATE_PATH + templatePath);
            using (var templateStream = new MemoryStream(templateBytes))
            using (var document = DocX.Load(templateStream))
            {
                foreach (var entry in placeholders)
                {
                    document.ReplaceText("{{" + entry.Key + "}}", entry.Value);
                }

                using (var outputStream = new MemoryStream())
                {
                    document.SaveAs(outputStream);
                    outputStream.Seek(0, SeekOrigin.Begin);
                    await File.WriteAllBytesAsync(OUTPUT_PATH + outputPath, outputStream.ToArray());
                }
            }
        }

    }
}
