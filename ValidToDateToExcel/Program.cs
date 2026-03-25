using DocumentFormat.OpenXml.Vml;
using System.Threading.Tasks;

namespace ValidToDateToExcel
{
    public class Program
    {
        static async Task Main(string[] args)
        {           
            Console.WriteLine("=== ValidToDateToExcel ===");
            var excelHandler = new ExcelHandler();
            do
            {
                Console.WriteLine();
                Console.WriteLine("Меню:");
                Console.WriteLine("1)Выбрать файл");
                Console.WriteLine("0)Выход");
                Console.Write("Ввод: ");
                string? input = Console.ReadLine();
                switch (input)
                {
                    case "1":
                        string? filePath = ShowOpenFileDialog();
                        if (string.IsNullOrEmpty(filePath))
                        {
                            Console.WriteLine("Путь не указан!");
                            continue;
                        }
                        else if (!File.Exists(filePath))
                        {
                            Console.WriteLine("Файл не найден!");
                            continue;
                        }                        
                        try
                        {
                            Console.WriteLine("Выполнение.");
                            await excelHandler.FillValidToDates(filePath);
                            Console.WriteLine();
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("Обработка завершена.");
                            Console.ResetColor();                           
                        }
                        catch (IOException ex) when (ex.Message.Contains("занят") || ex.Message.Contains("used by another process"))
                        {
                            Console.WriteLine("Файл открыт в другой программе. Закройте его и повторите попытку.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Ошибка {ex.Message}");
                        }                       
                        break;
                    case "0":
                        return;
                    default:
                        Console.WriteLine("Неверный ввод.");
                        break;
                }                          
            } while (true);
        }
        static string? ShowOpenFileDialog()
        {
            string? path = null;
            var t = new Thread(() =>
            {
                using var dialog = new OpenFileDialog();
                dialog.Title = "Выберите файл";
                dialog.Filter = "Excel файлы|*.xlsx;*.xls";
                if (dialog.ShowDialog() == DialogResult.OK)
                    path = dialog.FileName;
            });
            t.SetApartmentState(ApartmentState.STA); 
            t.Start();
            t.Join(); 
            return path;
        }
    }
}
