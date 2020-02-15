using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using ClosedXML.Excel;

namespace MailAdressParser
{
    class Program
    {
        public static void ReadExcel(string filePath, string fileSave)
        {
            RecipientsList RecipentList = new RecipientsList();
            if(!File.Exists(filePath))return;
            // начало использования библиотеке ClosedXML
            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);
            // получим все строки в файле
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row
            // пример чтения строк файла.
            var count = rows.Count();
            Recipient recipient;
            foreach (var row in rows)
            {
                var id = RecipentList.Recipients.Count + 1;
                recipient = new Recipient
                {
                    Name = row.Cell(2).Value.ToString(),
                    Address = row.Cell(8).Value.ToString(),
                    Company = row.Cell(4).Value.ToString(),
                    Id = id,
                    INN = row.Cell(6).Value.ToString(),
                    Occupation = row.Cell(5).Value.ToString(),
                    Phone = row.Cell(7).Value.ToString(),
                    Position = row.Cell(3).Value.ToString(),
                    WasSent = row.Cell(1).Value.ToString() == "+"
                };
                if(!RecipentList.Recipients.Contains(recipient)) RecipentList.Recipients.Add(recipient);
            }

            worksheet = workbook.Worksheet(2);

            rows = worksheet.RangeUsed().RowsUsed();
            count = rows.Count();
            foreach (var row in rows)
            {
                var id = RecipentList.Recipients.Count + 1;
                recipient = new Recipient
                {
                    Id = id,
                    Name = row.Cell(1).Value.ToString(),
                    Address = row.Cell(2).Value.ToString()
                };

                if (!RecipentList.Recipients.Contains(recipient)) RecipentList.Recipients.Add(recipient);
            }

            try
            {
                RecipentList.Save(fileSave);
            }
            catch (Exception e)
            {
                
            }
            workbook.Dispose();
        }
        static void Main(string[] args)
        {
            var filePath = Path.Combine(Environment.CurrentDirectory,$"Data\\контакты.xlsx");
            var fileSave = Path.Combine(Environment.CurrentDirectory, $"Data\\Recipients.info");
            var directory = Path.Combine(Environment.CurrentDirectory, "Data");
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);

            ReadExcel(filePath, fileSave);
        }
    }
}
