using ArrayToExcel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ImportDataExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //TestDataTable();
            //ExportSupplier();
            ImportSupplier2(@"D:\KERJA\AXATA\AxataPOS_V3\AxataPOS-V3\Src\AxataPOSV3\TestDllForm\bin\Supplier.xlsx");
            Console.ReadLine();
        }

        static IEnumerable<SomeItem> SomeItems = Enumerable.Range(1, 10).Select(x => new SomeItem
        {
            Prop1 = $"Text #{x}",
            Prop2 = x * 1000,
            Prop3 = DateTime.Now.AddDays(-x),
        });

        static void TestDataTable()
        {
            var table = new DataTable("Table2");

            table.Columns.Add("Column #1", typeof(string));
            table.Columns.Add("Column #2", typeof(int));
            table.Columns.Add("Column #3", typeof(DateTime));

            for (var x = 1; x <= 100; x++)
                table.Rows.Add($"Text #{x}", x * 1000, DateTime.Now.AddDays(-x));

            /*var excel = table.ToExcel(s => s
                .AddSheet(table, ss => ss.SheetName("Table2"))
                .AddSheet(SomeItems));*/
            var excel = table.ToExcel();

            File.WriteAllBytes($@"..\{nameof(TestDataTable)}.xlsx", excel);
        }

        static void ExportSupplier()
        {
            DataTable x = AxataPOS.Models.Supplier.GetData("", "", "");
            var excel = x.ToExcel();
            File.WriteAllBytes($@"..\Supplier.xlsx", excel);
        }
        static void ImportSupplier2(string Path)
        {
            DataTable x = ArrayToExcel.EXcel2DataSet.MyExcelData(Path);
            PrintDataTable(x);
        }

        static void PrintDataTable(DataTable data)
        {
            foreach (DataColumn item in data.Columns)
            {
                System.Console.Write(item.ColumnName.ToString() + " | ");
            }
            println("");
            int a = 0;
            foreach (DataRow item in data.Rows)
            {
                if (a == 2)
                {
                    break;
                }
                foreach (var dt in item.ItemArray)
                {
                    System.Console.Write(dt + " | ");
                }
                System.Console.WriteLine();
                a++;
            }
        }

        static void println(string msg)
        {
            System.Console.WriteLine(msg);
        }
    }
    internal class SomeItem
    {
        public string Prop1 { get; set; }
        public int Prop2 { get; set; }
        public DateTime Prop3 { get; set; }
    }
}
