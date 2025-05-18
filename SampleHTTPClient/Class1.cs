//namespace SampleHTTPClient
//{

//    // Project: MainframeToAzureSharePoint
//    // Azure Functions: OnDemandJob, NightlyBatchJob

//    using System;
//    using System.IO;
//    using System.Linq;
//    using System.Threading.Tasks;
//    using Microsoft.AspNetCore.Mvc;
//    using Microsoft.Azure.WebJobs;
//    using Microsoft.Azure.WebJobs.Extensions.Http;
//    using Microsoft.AspNetCore.Http;
//    using Microsoft.Extensions.Logging;
//    using Renci.SshNet;
//    using Azure.Storage.Blobs;
//    using Azure.Storage.Files.Shares;
//    using Microsoft.Graph;
//    using Microsoft.Identity.Client;

//    public static class FileTransferFunctions
//    {
//        private static string SftpHost = Environment.GetEnvironmentVariable("SFTP_HOST");
//        private static string SftpUser = Environment.GetEnvironmentVariable("SFTP_USER");
//        private static string SftpPass = Environment.GetEnvironmentVariable("SFTP_PASS");
//        private static string AzureBlobConnectionString = Environment.GetEnvironmentVariable("AZURE_BLOB_CONNECTION_STRING");
//        private static string AzureShareName = Environment.GetEnvironmentVariable("AZURE_FILE_SHARE");
//        private static string GraphSiteId = Environment.GetEnvironmentVariable("GRAPH_SITE_ID");

//        [FunctionName("OnDemandJob")]
//        public static async Task<IActionResult> RunOnDemand(
//            [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req,
//            ILogger log)
//        {
//            string jobId = req.Query["jobId"];
//            await RunFileTransfer(log);
//            return new OkObjectResult($"On-Demand Job Completed: {jobId}");
//        }

//        [FunctionName("NightlyBatchJob")]
//        public static async Task RunNightly([
//            TimerTrigger("0 0 2 * * *")] TimerInfo timer,
//            ILogger log)
//        {
//            await RunFileTransfer(log);
//        }

//        private static async Task RunFileTransfer(ILogger log)
//        {
//            var localTemp = Path.Combine(Path.GetTempPath(), "mainframe");
//            Directory.CreateDirectory(localTemp);

//            // Step 1: Download from Mainframe (SFTP)
//            using (var sftp = new SftpClient(SftpHost, SftpUser, SftpPass))
//            {
//                sftp.Connect();
//                var files = sftp.ListDirectory("/export");
//                foreach (var file in files.Where(f => !f.IsDirectory))
//                {
//                    var localFile = Path.Combine(localTemp, file.Name);
//                    using var stream = File.Create(localFile);
//                    sftp.DownloadFile(file.FullName, stream);
//                    log.LogInformation($"Downloaded: {file.Name}");
//                }
//                sftp.Disconnect();
//            }

//            // Step 2: Upload to Azure Blob
//            var blobClient = new BlobServiceClient(AzureBlobConnectionString);
//            var container = blobClient.GetBlobContainerClient("mainframe-files");
//            await container.CreateIfNotExistsAsync();

//            foreach (var file in Directory.GetFiles(localTemp))
//            {
//                using var stream = File.OpenRead(file);
//                await container.UploadBlobAsync(Path.GetFileName(file), stream);
//                log.LogInformation($"Uploaded to Blob: {Path.GetFileName(file)}");
//            }

//            // Step 3: Upload to SharePoint
//            var graphClient = GetGraphClient();
//            foreach (var file in Directory.GetFiles(localTemp))
//            {
//                using var stream = File.OpenRead(file);
//                await graphClient.Sites[GraphSiteId].Drive.Root
//                    .ItemWithPath("Templates/" + Path.GetFileName(file))
//                    .Content.Request().PutAsync<DriveItem>(stream);
//                log.LogInformation($"Uploaded to SharePoint: {Path.GetFileName(file)}");
//            }
//        }

//        private static GraphServiceClient GetGraphClient()
//        {
//            var clientId = Environment.GetEnvironmentVariable("GRAPH_CLIENT_ID");
//            var tenantId = Environment.GetEnvironmentVariable("GRAPH_TENANT_ID");
//            var clientSecret = Environment.GetEnvironmentVariable("GRAPH_CLIENT_SECRET");

//            var confidentialClient = ConfidentialClientApplicationBuilder
//                .Create(clientId)
//                .WithClientSecret(clientSecret)
//                .WithTenantId(tenantId)
//                .Build();

//            var authProvider = new ClientCredentialProvider(confidentialClient);
//            return new GraphServiceClient(authProvider);
//        }
//    }

//}