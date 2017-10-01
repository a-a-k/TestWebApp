using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ContactsApp.Utils;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

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

            var excel = new ExcelWorker();
            var azureStorageTable = new AzureCloudStorageWorker();

            try
            {
                await excel.Load(file);
                
                if (!excel.CheckRequiredColumns(new List<string> { "фио", "телефон" }))
                {
                    return View("Index", "Ошибка - отсутствуют обязательные для заполнения колонки");
                }

                var connectionString = Program.Configuration["ConnectionStrings:AzureTableStorage"];
                if (!await azureStorageTable.ConnectAsync(connectionString, "Contacts"))
                {
                    return View("Index", "Ошибка соединения с хранилищем");
                }

                await azureStorageTable.InsertOrReplace(excel.Read(true));
                
                return View("Index", $"Записано {excel.Saved} контактов, пропущено - {excel.Skipped}");
            }
            catch (Exception ex)
            {
                return View("Index", $"Ошибка - {ex.Message}");
            }
        }
    }
}