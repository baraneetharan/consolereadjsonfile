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

                // foreach (var emp in data)
                // {
                //     Console.WriteLine("{0} {1} {2}:", emp.Id, emp.Name, emp.dept);

                //     foreach (var punch in emp.Punchs)
                //         Console.WriteLine("\t{0} {1} {2}:", punch.slno, punch.pdate, punch.pday);
                // }
            }
        }
        public static void writexlsx()
        {
            foreach (var emp in data)
            {
                // Console.WriteLine("{0} {1} {2}:", emp.Id, emp.Name, emp.dept);
                var memoryStream = new MemoryStream();

                using (var fs = new FileStream(emp.Id + ".xlsx", FileMode.Create, FileAccess.Write))
                {
                    IWorkbook workbook = new XSSFWorkbook();
                    ISheet excelSheet = workbook.CreateSheet("Sheet1");

                    IRow row = excelSheet.CreateRow(0);
                    row.CreateCell(0).SetCellValue("Employee ID : ");
                    row.CreateCell(1).SetCellValue(emp.Id);

                    IRow row1 = excelSheet.CreateRow(1);
                    row1.CreateCell(0).SetCellValue("Employee Name : ");
                    row1.CreateCell(1).SetCellValue(emp.Name);

                    IRow row2 = excelSheet.CreateRow(2);
                    row2.CreateCell(0).SetCellValue("Department : ");
                    row2.CreateCell(1).SetCellValue(emp.dept);

                    IRow row3 = excelSheet.CreateRow(3);
                    IRow row4 = excelSheet.CreateRow(4);

                    // IRow rowpunch = excelSheet.CreateRow(5);
                    // rowpunch.CreateCell(0).SetCellValue("slno");
                    // rowpunch.CreateCell(1).SetCellValue("pdate");
                    // rowpunch.CreateCell(2).SetCellValue("pday");

                    DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(emp.Punchs), (typeof(DataTable)));
                    List<String> columns = new List<string>();
                    IRow punchrow = excelSheet.CreateRow(5);
                    int columnIndex = 0;

                    foreach (System.Data.DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);
                        punchrow.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                        columnIndex++;
                    }

                    int rowIndex = 6;
                    foreach (DataRow dsrow in table.Rows)
                    {
                        punchrow = excelSheet.CreateRow(rowIndex);
                        int cellIndex = 0;
                        foreach (String col in columns)
                        {
                            punchrow.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                            cellIndex++;
                        }

                        rowIndex++;
                    }
                    workbook.Write(fs);
                }
            }
        }
    }
}