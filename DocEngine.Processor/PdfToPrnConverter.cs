
namespace DocEngine.Processor
{
    using Microsoft.Win32;
    using SixLabors.ImageSharp;
    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Drawing.Printing;
    using System.IO;
    using static System.Runtime.InteropServices.JavaScript.JSType;
    using System.IO.Packaging;
    using System.Reflection;
    using System.Runtime.Intrinsics.X86;
    using NLog;

    public class PdfToPrnConverter
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public static void ConvertAllInFolder(string inputFolder, string outputFolder, string archiveFolder)
        {
            #region Directory Check
            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            if (!Directory.Exists(archiveFolder))
                Directory.CreateDirectory(archiveFolder);

            #endregion

            var pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
            logger.Info("PDF to Prm Batch conversion started...");
            if (pdfFiles.Length == 0)
            {
                logger.Info("No PDF files found in the input folder.");
                return;
            }
            foreach (var file in pdfFiles)
            {
                string fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                string prnFile = System.IO.Path.Combine(outputFolder, fileName + ".prn");
                string archiveFile = System.IO.Path.Combine(archiveFolder, System.IO.Path.GetFileName(file));

                try
                {
                    //Console.WriteLine($"Converting: {fileName}.pdf -> {fileName}.prn");
                    ConvertPdfToPrnUsingPrintCmd(file, prnFile);
                    //Convert(file, prnFile);

                    File.Move(file, archiveFile);
                    //logger.Info($"Moved {fileName}.docx to archive folder: {archiveFile}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting {fileName}: {ex.Message}");
                }
            }

            logger.Info("PDF to PRN Batch conversion completed.");
        }

        public static void Convert(string pdfFilePath, string prnFilePath)
        {
            try
            {
                // Validate input file
                if (!File.Exists(pdfFilePath))
                {
                    logger.Info("PDF file does not exist.");
                    return;
                }

                // Configure print settings
                PrintDocument printDoc = new PrintDocument();
                printDoc.PrinterSettings.PrinterName = "Microsoft Print to PDF"; // Ensure this printer is installed
                printDoc.PrinterSettings.PrintToFile = true;
                printDoc.PrinterSettings.PrintFileName = prnFilePath;

                // Start the default PDF reader to print the PDF
                ProcessStartInfo printProcessInfo = new ProcessStartInfo
                {
                    Verb = "print",
                    FileName = pdfFilePath,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (Process printProcess = new Process())
                {
                    printProcess.StartInfo = printProcessInfo;
                    printProcess.Start();
                    printProcess.WaitForExit(30000); // Wait up to 30 seconds for printing to complete

                    if (!printProcess.HasExited)
                    {
                        printProcess.Kill();
                        Console.WriteLine("Printing timed out.");
                        return;
                    }

                    // Note: Edge's --print-to-pdf creates a PDF, not a PRN. Rename or adjust as needed.
                    string tempPdf = Path.Combine(Path.GetDirectoryName(pdfFilePath), "temp_output.pdf");
                    if (File.Exists(tempPdf))
                    {
                        File.Move(tempPdf, prnFilePath);
                    }
                }

                // Check if the PRN file was created
                if (File.Exists(prnFilePath))
                {
                    Console.WriteLine($"PRN file created successfully at: {prnFilePath}");
                }
                else
                {
                    Console.WriteLine("Failed to create PRN file.");
                }
            }
            catch (System.ComponentModel.Win32Exception ex)
            {
                Console.WriteLine($"Error Code: {ex.ErrorCode}, Native Error: {ex.NativeErrorCode}, Message: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to check PDF file association
        static string GetPdfAssociatedProgram()
        {
            try
            {
                using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(@".pdf\OpenWithProgids"))
                {
                    if (key != null)
                    {
                        string progId = key.GetValueNames()[0];
                        using (RegistryKey progKey = Registry.ClassesRoot.OpenSubKey(progId + @"\shell\print\command"))
                        {
                            if (progKey != null)
                            {
                                return progKey.GetValue("")?.ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking PDF association: {ex.Message}");
            }
            return null;
        }

        // Helper method to check if a printer is available
        static bool IsPrinterAvailable(string printerName)
        {
            try
            {
                foreach (string printer in PrinterSettings.InstalledPrinters)
                {
                    if (printer.Equals(printerName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking printers: {ex.Message}");
            }
            return false;
        }

        /// <summary>
        /// Converts a PDF file to a PRN file using the Windows PRINT command.
        ///  Requires a printer that can output to file configured as default
        /// </summary>
        /// <param name="pdfPath"></param>
        /// <param name="prnPath"></param>
        /// <exception cref="Exception"></exception>
        static void ConvertPdfToPrnUsingPrintCmd(string pdfPath, string prnPath)
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/C PRINT /D:\"{prnPath}\" \"{pdfPath}\"",
                CreateNoWindow = true,
                UseShellExecute = false
            };

            using (Process process = Process.Start(psi))
            {
                process.WaitForExit();
                if (process.ExitCode != 0)
                {
                    throw new Exception("Print command failed");
                }
            }
        }
    }
}
