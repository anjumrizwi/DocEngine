using static System.Net.Mime.MediaTypeNames;
using Aspose.Words.Replacing;
using Aspose.Words;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Data;

namespace DocEngine.AsposeWords
{

    public static class AsposeMailMergeHelper
    {
        public static void ExecuteDataTable(string templatePath, string outputDirectory, List<Dictionary<string, string>> recipientData)
        {
            DataTable table = new DataTable("Test");
            table.Columns.Add("CustomerName");
            table.Columns.Add("Address");
            table.Columns.Add("Company");
            
            table.Rows.Add(new object[] { "Thomas Hardy", "120 Hanover Sq., London", "Valtech" });
            table.Rows.Add(new object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino", "Metaoption" });

            // Below are two ways of using a DataTable as the data source for a mail merge.
            // 1 -  Use the entire table for the mail merge to create one output mail merge document for every row in the table:
            Document doc = CreateSourceDocExecuteDataTable(templatePath);

            doc.MailMerge.Execute(table);

            doc.Save(outputDirectory + "MailMerge.ExecuteDataTable.WholeTable.docx");

            // 2 -  Use one row of the table to create one output mail merge document:
            doc = CreateSourceDocExecuteDataTable(templatePath);

            doc.MailMerge.Execute(table.Rows[1]);

            doc.Save(outputDirectory + "MailMerge.ExecuteDataTable.OneRow.docx");
        }

        public static void ExecuteDataTable(string outputDirectory)
        {
            DataTable table = new DataTable("Test");
            table.Columns.Add("CustomerName");
            table.Columns.Add("Address");
            table.Rows.Add(new object[] { "Thomas Hardy", "120 Hanover Sq., London" });
            table.Rows.Add(new object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });

            // Below are two ways of using a DataTable as the data source for a mail merge.
            // 1 -  Use the entire table for the mail merge to create one output mail merge document for every row in the table:
            Document doc = CreateSourceDocExecuteDataTable();

            doc.MailMerge.Execute(table);
            var fileName = "MailMerge.ExecuteDataTable.WholeTable.docx";
            doc.Save(outputDirectory + fileName);

            // 2 -  Use one row of the table to create one output mail merge document:
            fileName = "MailMerge.ExecuteDataTable.OneRow.docx";
            var stopwatch = Stopwatch.StartNew();
            doc = CreateSourceDocExecuteDataTable();

            doc.MailMerge.Execute(table.Rows[1]);

            doc.Save(outputDirectory + fileName);
            stopwatch.Stop();
            Console.WriteLine($"[SUCCESS] {fileName} converted in {stopwatch.Elapsed.TotalSeconds:F2} seconds");
        }
        /// <summary>
        /// Creates a mail merge source document.
        /// </summary>
        private static Document CreateSourceDocExecuteDataTable()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            //builder.InsertField(" MERGEFIELD CustomerName ");
            //builder.InsertParagraph();
            //builder.InsertField(" MERGEFIELD Address ");

            return doc;
        }

        /// <summary>
        /// Creates a mail merge source document.
        /// </summary>
        private static Document CreateSourceDocExecuteDataTable(string templatePath)
        {
            Document doc = new Document(templatePath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            //builder.InsertField(" MERGEFIELD CustomerName ");
            //builder.InsertParagraph();
            //builder.InsertField(" MERGEFIELD Address ");

            return doc;
        }
        /// <summary>
        /// Performs mail merge for multiple recipients and creates one output file per recipient.
        /// </summary>
        /// <param name="templatePath">Path to the base template .docx file</param>
        /// <param name="outputDirectory">Folder where merged files will be saved</param>
        /// <param name="recipientData">List of field dictionaries for each recipient</param>
        public static void PerformBatchMailMerge(string templatePath, string outputDirectory, List<Dictionary<string, string>> recipientData)
        {
            //Directory.CreateDirectory(outputDirectory);
            //templatePath = templatePath;
            //outputDirectory = OUTPUT_PATH;

            for (int i = 0; i < recipientData.Count; i++)
            {
                string fileName = $"Merged_{i + 1}_{Guid.NewGuid().ToString()}.docx"; ;
                string outputPath = Path.Combine(outputDirectory, fileName);

                var doc = new Document(templatePath);
               
                HeaderFooterCollection headersFooters = doc.FirstSection.HeadersFooters;
                HeaderFooter footer = headersFooters[HeaderFooterType.FooterPrimary];

                FindReplaceOptions options = new FindReplaceOptions { MatchCase = false, FindWholeWordsOnly = false };
                footer.Range.Replace("(C) 2006 Aspose Pty Ltd.", "Copyright (C) 2020 by Aspose Pty Ltd.", options);

                FindReplaceOptions options1 = new FindReplaceOptions();
                //options.ReplacingCallback = new ReplaceWithHtmlEvaluator(options);

                doc.Range.Replace(new Regex(@" <CustomerName>,"), string.Empty, options);

                doc.Save(outputPath, SaveFormat.Docx);

                
            }
        }

        public static void PerformMailMerge(string templatePath, string outputPath, Dictionary<string, string> placeholders)
        {
            // Copy the template to output to avoid modifying original
            File.Copy(templatePath, outputPath, true);

            //using (var doc = WordprocessingDocument.Open(outputPath, true))
            //{
            //    var body = doc.MainDocumentPart.Document.Body;

            //    foreach (var text in body.Descendants<WordText>())
            //    {
            //        foreach (var pair in placeholders)
            //        {
            //            if (text.Text.Contains("{{" + pair.Key + "}}"))
            //            {
            //                text.Text = text.Text.Replace("{{" + pair.Key + "}}", pair.Value);
            //            }
            //        }
            //    }

            //    doc.MainDocumentPart.Document.Save();
            //}
        }
    }

}
