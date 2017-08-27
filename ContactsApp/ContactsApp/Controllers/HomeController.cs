using ContactsApp.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ContactsApp.Controllers
{
    public class HomeController : Controller
    {

        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private readonly IHostingEnvironment _hostingEnvironment;

        public HomeController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> Index(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return View("Index", "Файл пуст или не выбран");
            }

            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    await file.CopyToAsync(memoryStream).ConfigureAwait(false);

                    using (var package = new ExcelPackage(memoryStream))
                    {
                        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                        var rowCount = worksheet.Dimension.Rows;
                        var colCount = worksheet.Dimension.Columns;
                        var columns = new Dictionary<string, byte>();
                        
                        for (byte col = 1; col <= colCount; col++)
                        {
                            columns.Add(worksheet.Cells[1, col].Value.ToString().ToLower(), col);
                        }

                        if (!columns.ContainsKey("фио") || !columns.ContainsKey("телефон"))
                        {
                            return View("Index", "Ошибка - отсутствуют обязательные для заполнения колонки");
                        }

                        if (!CloudStorageAccount.TryParse("DefaultEndpointsProtocol=http;AccountName=devstoreaccount1;AccountKey=Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw==;TableEndpoint=http://127.0.0.1:10002/devstoreaccount1", out var storageAccount))
                        {
                            return View("Index", "Ошибка CloudStorageAccount.TryParse");
                        }

                        var tableClient = storageAccount.CreateCloudTableClient();
                        var table = tableClient.GetTableReference("Contacts");
                        await table.CreateIfNotExistsAsync();
                        var skipped = 0;
                        var saved = 0;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            var sourcePhone = worksheet.Cells[row, columns.GetValueOrDefault("телефон")].Text;
                            var sourceName = worksheet.Cells[row, columns.GetValueOrDefault("фио")].Text;
                            if (string.IsNullOrEmpty(sourcePhone) || string.IsNullOrEmpty(sourceName))
                            {
                                skipped++;
                                continue;
                            }

                            var phone = ClearPhone(sourcePhone);
                            var name = ClearName(sourceName);
                            if (string.IsNullOrEmpty(phone) || string.IsNullOrEmpty(name))
                            {
                                skipped++;
                                continue;
                            }

                            var zip = worksheet.Cells[row, columns.GetValueOrDefault("почтовый индекс")].Text;
                            var region = worksheet.Cells[row, columns.GetValueOrDefault("регион")].Text;
                            var city = worksheet.Cells[row, columns.GetValueOrDefault("город")].Text;
                            var address = worksheet.Cells[row, columns.GetValueOrDefault("адрес")].Text;
                            var email = worksheet.Cells[row, columns.GetValueOrDefault("email")].Text;
                        
                            await table.ExecuteAsync(TableOperation.InsertOrReplace(new Contact(phone, name, zip, region, city, address, email)));
                            saved++;
                        }

                        if (saved == 0)
                        {
                            return View("Index", $"Записано {saved} контактов, пропущено - {skipped}");
                        }

                        return View("Index", $"Записано {saved} контактов, пропущено - {skipped}");
                    }
                }

            }
            catch (Exception ex)
            {
                return View("Index", $"Ошибка - {ex.Message}");
            }
        }

        private string ClearName(string sourceName)
        {
            if (!sourceName.All(x => Char.IsLetter(x) || Char.IsWhiteSpace(x)))
            {
                return string.Empty;
            }

            var parts = sourceName.Split(' ');
            var result = parts.Aggregate((s1, s2) => $"{s1.Trim()} {s2.Trim()}").TrimEnd();
            
            return result;
        }

        private string ClearPhone(string sourcePhone)
        {
            var correctDelimeters = new char[] { '(', ')', '-', ' ', '+' };
            if (!sourcePhone.All(x => Char.IsDigit(x) || correctDelimeters.Contains(x)) || sourcePhone.Length < 10)
            {
                return string.Empty;
            }
            var result = new string(sourcePhone.Select(x => correctDelimeters.Contains(x) ? ' ' : x).ToArray()).Replace(" ", "");
            var firstDigit = result.First();
            result = firstDigit == '8' ? result : firstDigit == '7' ? $"8{result.Substring(1)}" : firstDigit == '9' ? $"8{result}" : string.Empty;

            return result.Length > 11 ? string.Empty : result;
        }
    }
}