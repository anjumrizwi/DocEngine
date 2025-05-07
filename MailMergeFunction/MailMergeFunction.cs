using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Azure.Storage.Blobs;
using SendGrid;
using SendGrid.Helpers.Mail;
using DocumentFormat.OpenXml.Packaging;
using System;

namespace MailMergeFunctionApp
{
    public static class MailMergeFunction
    {
        [FunctionName("MailMergeFunction")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("Mail merge function triggered.");

            var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var input = JsonSerializer.Deserialize<MergeRequest>(requestBody);

            if (input == null || string.IsNullOrEmpty(input.Email))
                return new BadRequestObjectResult("Invalid input.");

            var blobService = new BlobServiceClient(Environment.GetEnvironmentVariable("AzureWebJobsStorage"));
            var templateBlob = blobService.GetBlobContainerClient("templates").GetBlobClient("template.docx");

            using var templateStream = new MemoryStream();
            await templateBlob.DownloadToAsync(templateStream);
            templateStream.Position = 0;

            var mergedStream = PerformMailMerge(templateStream, input);

            var fileName = $"Letter_{input.FirstName}_{input.LastName}.docx";
            var outputBlob = blobService.GetBlobContainerClient("output").GetBlobClient(fileName);
            mergedStream.Position = 0;
            await outputBlob.UploadAsync(mergedStream, overwrite: true);

            await SendEmail(input, mergedStream, fileName);

            return new OkObjectResult($"Document generated and sent to {input.Email}");
        }

        private static MemoryStream PerformMailMerge(Stream template, MergeRequest data)
        {
            var stream = new MemoryStream();
            template.CopyTo(stream);
            stream.Position = 0;

            using var wordDoc = WordprocessingDocument.Open(stream, true);
            var docText = new StreamReader(wordDoc.MainDocumentPart.GetStream()).ReadToEnd();

            docText = docText.Replace("{{FirstName}}", data.FirstName)
                             .Replace("{{LastName}}", data.LastName)
                             .Replace("{{Address}}", data.Address ?? "");

            var writer = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create));
            writer.Write(docText);
            writer.Close();

            stream.Position = 0;
            return stream;
        }

        private static async Task SendEmail(MergeRequest data, MemoryStream docStream, string fileName)
        {
            var apiKey = Environment.GetEnvironmentVariable("SendGridApiKey");
            var client = new SendGridClient(apiKey);

            var from = new EmailAddress("no-reply@example.com", "MailMerge App");
            var to = new EmailAddress(data.Email, $"{data.FirstName} {data.LastName}");
            var subject = "Your Mail Merge Letter";
            var content = "Attached is your personalized document.";

            var msg = MailHelper.CreateSingleEmail(from, to, subject, content, content);

            docStream.Position = 0;
            var bytes = docStream.ToArray();
            msg.AddAttachment(fileName, Convert.ToBase64String(bytes), "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

            await client.SendEmailAsync(msg);
        }

        private class MergeRequest
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string Address { get; set; }
            public string Email { get; set; }
        }
    }
}
