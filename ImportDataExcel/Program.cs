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
            TestDataTable();
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
    }
    internal class SomeItem
    {
        public string Prop1 { get; set; }
        public int Prop2 { get; set; }
        public DateTime Prop3 { get; set; }
    }
}
