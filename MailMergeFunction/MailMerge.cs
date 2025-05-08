using System.Net;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;

namespace MailMergeFunction
{
    public class MailMerge
    {
        private readonly ILogger<MailMerge> _logger;

        public MailMerge(ILogger<MailMerge> logger)
        {
            _logger = logger;
        }

        [Function("MailMerge")]
        public IActionResult Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequest req)
        {

            _logger.LogInformation("C# HTTP trigger function processed a request.");

            return new OkObjectResult("Welcome to Azure Functions!");
        }

        [Function("MailMerge")]
        public static async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req,
        ILogger log)
        {
            log.LogInformation("MailMerge function triggered.");

            string template = await new StreamReader(req.Body).ReadToEndAsync();

            // Example: Simple Replace. You can use RazorEngine, Scriban, or MailMerge libraries for more.
            string result = template
                .Replace("{{Name}}", "Anjum Rizwi")
                .Replace("{{Date}}", System.DateTime.Now.ToShortDateString());

            return new ContentResult
            {
                Content = result,
                ContentType = "text/plain",
                StatusCode = (int)HttpStatusCode.OK
            };
        }
    }
}
