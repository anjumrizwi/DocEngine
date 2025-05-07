using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace MailMergeTestApp
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var client = new HttpClient();
            var url = "https://<your-function-app>.azurewebsites.net/api/MailMergeFunction"; // Replace with your actual URL

            var json = @"{
    ""FirstName"": ""John"",
    ""LastName"": ""Doe"",
    ""Address"": ""123 Test Street"",
    ""Email"": ""john.doe@example.com""
}";


            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await client.PostAsync(url, content);

            var responseContent = await response.Content.ReadAsStringAsync();
            Console.WriteLine($"Status: {response.StatusCode}");
            Console.WriteLine($"Response: {responseContent}");
        }
    }
}
