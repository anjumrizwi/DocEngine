// ----------------------------
// Sample HTTP Client (C# Console)
// ----------------------------

using System.Net.Http;
using System.Threading.Tasks;

class Program
{
    static async Task Main()
    {
        var httpClient = new HttpClient();
        var response = await httpClient.PostAsync(
            "https://<your-function-app>.azurewebsites.net/api/OnDemandJob?jobId=Test123",
            null);
        var result = await response.Content.ReadAsStringAsync();
        Console.WriteLine(result);
    }
}
