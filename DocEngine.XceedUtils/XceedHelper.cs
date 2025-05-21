namespace DocEngine.XceedUtils
{
    public class XceedHelper
    {
        public static void Convert()
        {
            // Private constructor to prevent instantiation
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
    }
}
