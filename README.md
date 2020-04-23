```
using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace consolereadjsonfile
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            using (StreamReader r = new StreamReader("D:\\baranee\\lockdown\\dotnetcore\\consolereadjsonfile\\jsondata.json"))
            {
                string json = r.ReadToEnd();
                // MyObject company = JsonConvert.DeserializeObject<MyObject>(json);
                List<User> data = JsonConvert.DeserializeObject<List<User>>(json);
                foreach (User room in data)
                {
                    long id = room.Id;
                    string name = room.Name;
                    string email = room.email;
                    Console.WriteLine(id + " " + name + " " + email);
                }
            }
        }
    }
}


```

```
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
        public static void readjson(){
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
        public static void writexlsx(){
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(data), (typeof(DataTable)));
            var memoryStream = new MemoryStream();

            using (var fs = new FileStream("Result.xlsx", FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Sheet1");
                
                excelSheet.SetColumnWidth(1, 12* 256);
                excelSheet.SetColumnWidth(3, 25 * 256);

                //font style1: underlined, italic, red color, fontsize=20
                IFont font1 = workbook.CreateFont();
                font1.Color = IndexedColors.Red.Index;
                font1.IsItalic = true;
                font1.Underline = FontUnderlineType.Double;
                font1.FontHeightInPoints = 20;

                //bind font with style 1
                // ICellStyle style1 = workbook.CreateCellStyle();
                // style1.SetFont(font1);

                ICellStyle cellStyleBlue = workbook.CreateCellStyle();
                cellStyleBlue.FillForegroundColor = IndexedColors.LightBlue.Index;
                cellStyleBlue.FillPattern = FillPattern.SolidForeground;
                cellStyleBlue.SetFont(font1);



                List<String> columns = new List<string>();
                IRow row = excelSheet.CreateRow(0);
                int columnIndex = 0;

                foreach (System.Data.DataColumn column in table.Columns)
                {
                    // columns.Add(column.ColumnName);
                    row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    row.Cells[columnIndex].CellStyle = cellStyleBlue;
                    columnIndex++;
                }

                int rowIndex = 1;
                foreach (DataRow dsrow in table.Rows)
                {
                    row = excelSheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    foreach (String col in columns)
                    {
                        row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        cellIndex++;
                    }

                    rowIndex++;
                }
                workbook.Write(fs);
            }
        }
    }
}


```