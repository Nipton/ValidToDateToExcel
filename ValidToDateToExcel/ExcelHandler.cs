using ClosedXML.Excel;
using System.Net;
using System.Text.Json;


namespace ValidToDateToExcel
{
    public class ExcelHandler
    {
        private static readonly HttpClient httpClient;
        private static readonly int[] yearsToCheck = [2025, 2024, 2026, 2023, 2022];
        private static readonly Dictionary<string, string> modelMap;
        static ExcelHandler()
        {
            httpClient = new HttpClient(new SocketsHttpHandler { PooledConnectionLifetime = TimeSpan.FromMinutes(2) });
            httpClient.DefaultRequestHeaders.Add("User-Agent", "MyExcelParser/1.0");
            modelMap = new()
            {
                { "ПУЛЬС", "\"ПУЛЬС\"" },
                { "ПУЛЬС СТК", "Пульс СТК" },
                { "ФОБОС 1", "ФОБОС 1" }
            };
        }
        public async Task FillValidToDates(string filePath)
        {
            using var wb = new XLWorkbook(filePath);
            var ws = wb.Worksheet(1);
            var rows = ws.RowsUsed().Skip(2);
            int count = 0;
            int total = rows.Count();
            foreach (var row in rows)
            {
                count++;
                WriteColored($"\rОбработка: {count}/{total}  ", ConsoleColor.Green);
                string number = row.Cell("H").GetString();
                string model = modelMap.TryGetValue(row.Cell("I").GetString(), out var mapped) ? mapped : "";
                if (string.IsNullOrEmpty(number) || string.IsNullOrEmpty(model))
                {
                    WriteColored($"\nВ строке {row.RowNumber()} не удалось получить нормер или марку.", ConsoleColor.Red);
                    continue;
                }
                var date = await GetValidToDate(number, model);               
                if (date == null)
                {
                    WriteColored($"\nНе удалось получить дату для строки №{row.RowNumber()}", ConsoleColor.Red);
                    continue;
                }    
                row.Cell("M").SetValue(date.Value);
                row.Cell("M").Style.NumberFormat.Format = "dd.MM.yyyy";             
            }
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string newFilePath = Path.Combine(desktopPath, fileName + "_result.xlsx");
            wb.SaveAs(newFilePath);
        }
        private async Task<DateTime?> GetValidToDate(string number, string model)
        {
            foreach (var year in yearsToCheck)
            {
                string url = $"https://fgis.gost.ru/fundmetrology/eapi/vri?mi_number={Uri.EscapeDataString(number)}&year={year}&mit_notation={Uri.EscapeDataString(model)}&sort=verification_date+desc&rows=1";
                int attempt = 1;
                do
                {
                    try
                    {
                        using var response = await httpClient.GetAsync(url);
                        if (response.StatusCode == HttpStatusCode.TooManyRequests)
                        {
                            Console.WriteLine($"Превышен лимит обращений к API. Ожидание. Попытка №{attempt}.");
                            attempt++;
                            await Task.Delay(2000);
                            continue;
                        }
                        if (!response.IsSuccessStatusCode)
                        {
                            Console.WriteLine($"Не удалось получить ответ на запрос для {number}, {year}");
                            break;
                        }
                        var json = await response.Content.ReadAsStringAsync();
                        using var doc = JsonDocument.Parse(json);
                        int count = doc.RootElement.GetProperty("result").GetProperty("count").GetInt32();
                        if (count > 0)
                        {
                            var item = doc.RootElement.GetProperty("result").GetProperty("items")[0];
                            var dateString = item.GetProperty("valid_date").GetString();
                            if(DateTime.TryParse(dateString, out var date))
                            {
                                return date;
                            }
                            Console.WriteLine($"Ошибка преобразования даты {dateString}");
                            return null;
                        }                        
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка в запросе с {number}, {year}: {ex.Message}. Попытка №{attempt}.");
                        if (attempt <= 3)
                            await Task.Delay(1000);
                        attempt++;
                    }
                } while (attempt <= 3);
                await Task.Delay(600);
            }
            return null;
        }
        private void WriteColored(string message, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            if(color == ConsoleColor.Green)
                Console.Write(message);
            else 
                Console.WriteLine(message);
            Console.ResetColor();
        }
    }
}
