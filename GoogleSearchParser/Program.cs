using System.Net.NetworkInformation;
using System.Web;
using Czlovek.HttpCommunication;
using GoogleSearchParser;
using HtmlAgilityPack;





ExcelHelper excelHelper = new ExcelHelper();

Console.WriteLine("Güncelleme sırasında dosyada bir kesinti oluşması ihtimaline karşın lütfen dosyanızı yedekleyin.");
Console.WriteLine("Excel dosyasının tam yolunu giriniz (örn: C:\\dosyalar\\firma.xlsx):");
string filePath = Console.ReadLine();

Console.WriteLine("Dosya okunuyor.");

var companyList = new List<CompanyModel>();
var fileReaded = false;

while (!fileReaded) {
    try
    {
        companyList = excelHelper.ReadExcel(filePath);
        fileReaded = true;

    }
    catch (Exception ex)
    {
        Console.WriteLine("Dosya okunamadı lütfen geçerli bir dosya yolu giriniz");

        Console.WriteLine("Excel dosyasının tam yolunu giriniz (örn: C:\\dosyalar\\firma.xlsx):");
        filePath = Console.ReadLine();
    }

}

int totalSteps = companyList.Count();

await UpdateCompanies(companyList.Take(300).ToList(), filePath);

Console.WriteLine(" ");



async Task UpdateCompanies(List<CompanyModel> companyList, string filePath, int currentStep = 0)
{
    foreach (var company in companyList)
    {
        string query = Uri.EscapeDataString(company.Name);

        string googleSearchUrl = $"https://www.google.com/search?q={query}";

        currentStep++;
        int progressPercentage = (int)((currentStep / (float)totalSteps) * 100);

        DrawProgressBar(progressPercentage, 50, "'" + company.Name + "'" + "Firmasının bilgileri aranıyor");

        try
        {
            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(Utils.RandomUserAgent);

            HttpResponseMessage response = await httpClient.GetAsync(googleSearchUrl);
            response.EnsureSuccessStatusCode();

            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(await response.Content.ReadAsStringAsync());

            var phoneNumberNode = htmlDoc.DocumentNode.SelectSingleNode("//span[contains(text(),'+44')]");
            var workingHoursNode = htmlDoc.DocumentNode.SelectSingleNode("//span[contains(text(),'saati:')]");
            var link = htmlDoc.DocumentNode.SelectSingleNode("//a[contains(., 'Web sitesi')]");
            var href = link?.GetAttributeValue("href", string.Empty);

            string phoneNumber = phoneNumberNode?.InnerText.Trim();
            string workingHoursText = workingHoursNode?.InnerText.Trim();
            string websiteText = href;

            if (!string.IsNullOrEmpty(phoneNumber))
            {
                DrawProgressBar(progressPercentage, 50, "Telefon Numarası Bulundu");
                company.Telephone = phoneNumber;
            }
            else
            {
                company.Telephone = "-";
            }

            if (!string.IsNullOrEmpty(workingHoursText))
            {
                DrawProgressBar(progressPercentage, 50, "Kapanış Saati Bulundu");
                company.WorkingHours = workingHoursText.Split("saati:").Last().Trim();
            }
            else
            {
                company.WorkingHours = "-";
            }
            if (!string.IsNullOrEmpty(websiteText))
            {
                DrawProgressBar(progressPercentage, 50, "Website bulundu");
                company.Website = websiteText.Split("q=").Last().Split("&")[0];
            }
            else
            {
                company.Website = "-";
            }
        }
        catch (Exception ex) {
            DrawProgressBar(progressPercentage, 50, "Bir hata oluştu yeniden deneniyor");
            WaitForInternetConnection();
            Thread.Sleep(2000);
        }
    }


    DrawProgressBar(99, 50, "Dosya güncelleniyor");

    excelHelper.UpdateExcel(filePath, companyList);

    DrawProgressBar(100, 50, "Dosya başarıyla güncellendi");

    Thread.Sleep(500);

    Console.WriteLine("");
    Console.WriteLine("Çıkmak için bir tuşa basınız");
    Console.ReadKey();

}

static void DrawProgressBar(int percentage, int barSize, string step)
{
    ClearCurrentConsoleLine();
    Console.Write("[");

    int position = (int)(barSize * (percentage / 100f));
    for (int i = 0; i < barSize; i++)
    {
        if (i < position)
            Console.Write("="); 
        else
            Console.Write(" "); 
    }

    Console.Write("] "); 
    Console.Write($"{percentage}% "); 

    int maxMessageLength = Console.WindowWidth - barSize - 10;
    string message = step.Length <= maxMessageLength ? step : step.Substring(0, maxMessageLength - 3) + "...";
    Console.Write(message);
}

 static void ClearCurrentConsoleLine()
{
    int currentLineCursor = Console.CursorTop;
    Console.SetCursorPosition(0, Console.CursorTop);
    Console.Write(new string(' ', Console.WindowWidth));
    Console.SetCursorPosition(0, currentLineCursor);
}

static void WaitForInternetConnection()
{
    while (!IsInternetAvailable())
    {
        Console.WriteLine("İnternet bağlantısı yok, birkaç saniye sonra yeniden denenecek...");
        Thread.Sleep(5000);
    }
}

static bool IsInternetAvailable()
{
    try
    {
        using (var ping = new Ping())
        {
            PingReply reply = ping.Send("google.com", 3000);
            return reply != null && reply.Status == IPStatus.Success;
        }
    }
    catch
    {
        return false;
    }
}

