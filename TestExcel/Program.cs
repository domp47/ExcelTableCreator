using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using ExcelTableCreator;


namespace TestExcel
{
    class Program
    {

        static void Main(string[] args)
        {
            var file = "C:\\temp\\Excel\\test.xlsx";
            // var file = "/tmp/Excel/test.xlsx";

            var englishDic = File.ReadAllLines("English.dic");

            var nColumns = 50;
            
            var columns = new List<string>();

            for (var i = 0; i < nColumns; i++) {
                columns.Add($"Column #{i}");
            }

            var tableRow = new List<TableCell>() {
                new TableCell("this"), new TableCell("is"), new TableCell("the first"), new TableCell("row"),
                new TableCell(47)
            };
            
            for (var i = tableRow.Count; i < nColumns; i++) {
                tableRow.Add(new TableCell($"Some Data :) {i}"));
            }

            var excelTable = new Table(columns, new[] {tableRow});

            const int seed = 32;
            var rnd = new Random(seed);

            var nRows = rnd.Next(250, 500);
            for (var i = 0; i < nRows; i++) {

                var row = new List<TableCell>();
                
                for (var c = 0; c < nColumns; c++) {
                    var type = rnd.Next(5);

                    switch (type) {
                        case 0:
                            row.Add(new TableCell(rnd.Next()));
                            break;
                        case 1:
                            row.Add(new TableCell(rnd.NextDouble()));
                            break;
                        default:
                            row.Add(new TableCell(englishDic[rnd.Next(englishDic.Length)]));
                            break;
                    }
                }
                
                excelTable.AddRow(row);
            }
            
            Console.WriteLine("Generating Excel File...");
            excelTable.GenerateExcel(file);
        }
    }
}