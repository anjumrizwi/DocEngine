namespace DocEngine.Processor
{
    internal class BaseConverter
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        public static void ConvertAllDocx2Pdf(string inputFolder, string outputFolder, string archiveFolder, Action<string, string> DocxToPdf)
        {
            if (!System.IO.Directory.Exists(inputFolder))
                throw new System.IO.DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            if (!System.IO.Directory.Exists(outputFolder))
                System.IO.Directory.CreateDirectory(outputFolder);

            if (!System.IO.Directory.Exists(archiveFolder))
                System.IO.Directory.CreateDirectory(archiveFolder);

            logger.Info("DocxToPdf Batch Process started...");
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start();

            string[] docxFiles = System.IO.Directory.GetFiles(inputFolder, "*.docx");
            int errCnt = 0;
            foreach (string file in docxFiles)
            {
                string fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                string outputFile = System.IO.Path.Combine(outputFolder, fileName + ".pdf");
                string archiveFile = System.IO.Path.Combine(archiveFolder, System.IO.Path.GetFileName(file));

                try
                {
                    DocxToPdf(file, outputFile);
                    System.IO.File.Move(file, archiveFile);
                }
                catch (System.Exception ex)
                {
                    errCnt++;
                    logger.Error($"Error converting {fileName}: {ex.Message}");
                }
            }

            stopwatch.Stop();
            if (errCnt > 0)
                logger.Error($"[ERROR]: {errCnt} docx files failed to convert to pdf.");

            logger.Info($"[SUCCESS]: {docxFiles.Length - errCnt} docx files converted to pdf in {stopwatch.Elapsed.TotalSeconds:F2} seconds");
            logger.Info("Docx2Pdf Batch conversion completed.");
        }
    }
}
