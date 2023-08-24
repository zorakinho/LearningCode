using System;
using System.Threading.Tasks;
using Web_Scraping.Service;

class Program
{
    static async Task Main(string[] args)
    {
        await ScrapingService.ScrapingDBSheet();
    }
}
