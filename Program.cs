using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace consolereadjsonfile
{
    class Program
    {
        public static List<Employee> data;
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            readjson();
            writexlsx();
        }
        public static void readjson()
        {
            using (StreamReader r = new StreamReader("D:\\baranee\\lockdown\\dotnetcore\\consolereadjsonfile\\punchdata.json"))
            {
                string json = r.ReadToEnd();
                // MyObject company = JsonConvert.DeserializeObject<MyObject>(json);
                data = JsonConvert.DeserializeObject<List<Employee>>(json);

                foreach (var emp in data)
                {
                    Console.WriteLine("{0} {1} {2}:", emp.Id, emp.Name, emp.dept);

                    foreach (var punch in emp.Punchs)
                        Console.WriteLine("\t{0} {1} {2}:", punch.slno, punch.pdate, punch.pday);
                }
            }
        }
        public static void writexlsx()
        {
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(data), (typeof(DataTable)));
            var memoryStream = new MemoryStream();

            using (var fs = new FileStream("Result.xlsx", FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Sheet1");


                List<String> rows = new List<string>();
                Console.WriteLine(table.Rows.Count);

                IRow row = excelSheet.CreateRow(0);
                int rowIndex = 1;

                foreach (DataRow dsrow in table.Rows)
                {
                    row = excelSheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    row.CreateCell(cellIndex).SetCellValue(dsrow[0].ToString());

                    rowIndex++;
                }



                workbook.Write(fs);
            }
        }
    }
}
