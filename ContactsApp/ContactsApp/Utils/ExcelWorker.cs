using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ContactsApp.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.WindowsAzure.Storage.Table;
using OfficeOpenXml;

namespace ContactsApp.Utils
{
    public class ExcelWorker
    {
        public int Skipped { get; private set; }
        public int Saved { get; private set; }

        private ExcelPackage _package;
        private Dictionary<string, byte> _columns = new Dictionary<string, byte> { { "", 0 } };

        internal async Task Load(IFormFile file)
        {
            using (var memoryStream = new MemoryStream())
            {
                await file.CopyToAsync(memoryStream).ConfigureAwait(false);
                _package = new ExcelPackage(memoryStream);
            }
        }

        internal bool CheckRequiredColumns(List<string> requiredColumns)
        {
            var worksheet = _package.Workbook.Worksheets.FirstOrDefault();
            if (worksheet != null)
            {
                var colCount = worksheet.Dimension.Columns;
                for (byte col = 1; col <= colCount; col++)
                {
                    _columns.Add(worksheet.Cells[1, col].Value.ToString().ToLower(), col);
                }
            }

            return requiredColumns.All(x => _columns.ContainsKey(x));
        }

        internal List<ITableEntity> Read(bool isNeedToDispose)
        {
            var list = new List<ITableEntity>();
            var worksheet = _package.Workbook.Worksheets.FirstOrDefault();
            if (worksheet != null)
            {
                var rowCount = worksheet.Dimension.Rows;

                Skipped = 0;
                Saved = 0;

                for (int row = 2; row <= rowCount; row++)
                {
                    var sourcePhone = worksheet.Cells[row, _columns.GetValueOrDefault("телефон")].Text;
                    var sourceName = worksheet.Cells[row, _columns.GetValueOrDefault("фио")].Text;
                    if (string.IsNullOrEmpty(sourcePhone) || string.IsNullOrEmpty(sourceName))
                    {
                        Skipped++;
                        continue;
                    }

                    var phone = ClearPhone(sourcePhone);
                    var name = ClearName(sourceName);
                    if (string.IsNullOrEmpty(phone) || string.IsNullOrEmpty(name))
                    {
                        Skipped++;
                        continue;
                    }

                    byte col;
                    var zip = _columns.TryGetValue("почтовый индекс", out col) ? worksheet.Cells[row, col].Text : string.Empty;
                    var region = _columns.TryGetValue("регион", out col) ? worksheet.Cells[row, col].Text : string.Empty;
                    var city = _columns.TryGetValue("город", out col) ? worksheet.Cells[row, col].Text : string.Empty;
                    var address = _columns.TryGetValue("адрес", out col) ? worksheet.Cells[row, col].Text : string.Empty;
                    var email = _columns.TryGetValue("email", out col) ? worksheet.Cells[row, col].Text : string.Empty;

                    list.Add(new Contact(phone, name, zip, region, city, address, email));
                    Saved++;
                }
            }

            if(isNeedToDispose) _package.Dispose();
            return list;
        }

        private string ClearName(string sourceName)
        {
            if (!sourceName.All(x => char.IsLetter(x) || char.IsWhiteSpace(x)))
            {
                return string.Empty;
            }

            var parts = sourceName.Split(' ');
            var result = parts.Aggregate((s1, s2) => $"{s1.Trim()} {s2.Trim()}").TrimEnd();

            return result;
        }

        private string ClearPhone(string sourcePhone)
        {
            var correctDelimeters = new[] { '(', ')', '-', ' ', '+' };
            if (!sourcePhone.All(x => char.IsDigit(x) || correctDelimeters.Contains(x)) || sourcePhone.Length < 10)
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
