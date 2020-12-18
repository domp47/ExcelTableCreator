using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Op = DocumentFormat.OpenXml.CustomProperties;

namespace ExcelTableCreator
{
    // ReSharper disable PossiblyMistakenUseOfParamsMethod
    public class Table
    {
        private readonly List<string> _tableColumns;
        private readonly List<List<TableCell>> _rows;

        /// <summary>
        /// Initialize an excel table object with the specified columns
        /// </summary>
        /// <param name="columns">List of Columns to define for the table</param>
        /// <exception cref="ArgumentNullException">No columns were specified</exception>
        public Table(IEnumerable<string> columns)
        {

            if (columns == null)
                throw new ArgumentNullException(nameof(columns));

            _tableColumns = columns.ToList();
            _rows = new List<List<TableCell>>();
            
            if (!_tableColumns.Any())
                throw new ArgumentNullException(nameof(columns));
        }

        /// <summary>
        /// Initialize an excel table object with specified columns and predefined rows
        /// </summary>
        /// <param name="columns">List of Columns to define for the table</param>
        /// <param name="rows">List of initial rows to set in the table</param>
        public Table(IEnumerable<string> columns, IEnumerable<IEnumerable<TableCell>> rows) : this(columns)
        {
            AddRowRange(rows);
        }

        /// <summary>
        /// Adds a row to the end of the list of rows
        /// </summary>
        /// <param name="row">The row to be added to the end of the list of rows</param>
        /// <exception cref="ArgumentException">Numbers of cells in the row doesn't equal the number of columns in the table</exception>
        public void AddRow(IEnumerable<TableCell> row)
        {
            var rowList = row.ToList();
            if (rowList.Count != _tableColumns.Count)
                throw new ArgumentException($"Row Count ({rowList.Count}) does not equal Columns Count ({_tableColumns.Count})");
            
            _rows.Add(rowList);
        }

        /// <summary>
        /// Adds the rows in the specified collection to the end of the list of rows
        /// </summary>
        /// <param name="rows">the collection of rows to be added to the end of the list of rows</param>
        /// <exception cref="ArgumentException">Numbers of cells in a row doesn't equal the number of columns in the table</exception>
        public void AddRowRange(IEnumerable<IEnumerable<TableCell>> rows)
        {
            foreach (var row in rows) {
                AddRow(row);
            }
        }

        /// <summary>
        /// Creates an excel workbook and saves it to the specified location.
        /// </summary>
        /// <param name="filePath">File name and path to save the file to</param>
        public void GenerateExcel(string filePath)
        {
            using var package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            CreateParts(package);
            package.Save();
            package.Close();
        }

        /// <summary>
        ///  Creates an excel workbook to the supplied stream
        /// </summary>
        /// <param name="stream">Stream to write to</param>
        public void GenerateExcel(Stream stream)
        {
            using var package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            CreateParts(package);
        }
        
        /// <summary>
        /// Creates all the necessary parts for the workbook 
        /// </summary>
        /// <param name="document">The spreadsheet document to create</param>
        private void CreateParts(SpreadsheetDocument document)
        {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            GenerateWorkbookContent(workbookPart);

            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPart1Content(workbookStylesPart);

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetContent(worksheetPart);

            TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>("rId2");
            GenerateTableDefinition(tableDefinitionPart);
        }
        
        #region ExcelFunctions
        
        /// <summary>
        /// Generates content of workbook. 
        /// </summary>
        /// <param name="workbookPart">WorkbookPart to add content to</param>
        private void GenerateWorkbookContent(WorkbookPart workbookPart)
        {
            Workbook workbook = new Workbook { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x15 xr xr6 xr10 xr2" }  };
            workbook.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            workbook.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            workbook.AddNamespaceDeclaration("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6");
            workbook.AddNamespaceDeclaration("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10");
            workbook.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");

            Sheets sheets = new Sheets();
            Sheet sheet1 = new Sheet { Name = "Report", SheetId = 1U, Id = "rId1" };
            
            sheets.Append(sheet1);
            
            workbook.Append(sheets);

            workbookPart.Workbook = workbook;
        }
        
        /// <summary>
        /// Generates Styles for Workbook. 
        /// </summary>
        /// <param name="workbookStylesPart">Workbook Style Part to add styles to</param>
        private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart)
        {
            Stylesheet stylesheet1 = new Stylesheet { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x14ac x16r2 xr" }  };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
            stylesheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

            Fonts fonts1 = new Fonts { Count = 1U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize { Val = 11D };
            Color color1 = new Color { Theme = 1U };
            FontName fontName1 = new FontName { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering { Val = 2 };
            FontScheme fontScheme1 = new FontScheme { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            fonts1.Append(font1);

            Fills fills1 = new Fills { Count = 2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders { Count = 1U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            borders1.Append(border1);


            CellFormats cellFormats1 = new CellFormats { Count = 2U };
            CellFormat cellFormat2 = new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U };

            CellFormat cellFormat3 = new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U, ApplyAlignment = true };
            Alignment alignment1 = new Alignment { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat3.Append(alignment1);

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);

            DifferentialFormats differentialFormats1 = new DifferentialFormats { Count = 4U };

            DifferentialFormat differentialFormat1 = new DifferentialFormat();
            Alignment alignment2 = new Alignment {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Bottom,
                TextRotation = 0U,
                WrapText = false,
                Indent = 0U,
                JustifyLastLine = false,
                ShrinkToFit = false,
                ReadingOrder = 0U
            };

            differentialFormat1.Append(alignment2);

            DifferentialFormat differentialFormat2 = new DifferentialFormat();
            Alignment alignment3 = new Alignment {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Bottom,
                TextRotation = 0U,
                WrapText = false,
                Indent = 0U,
                JustifyLastLine = false,
                ShrinkToFit = false,
                ReadingOrder = 0U
            };

            differentialFormat2.Append(alignment3);

            DifferentialFormat differentialFormat3 = new DifferentialFormat();
            Alignment alignment4 = new Alignment {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Bottom,
                TextRotation = 0U,
                WrapText = false, 
                Indent = 0U,
                JustifyLastLine = false,
                ShrinkToFit = false,
                ReadingOrder = 0U
            };

            differentialFormat3.Append(alignment4);

            DifferentialFormat differentialFormat4 = new DifferentialFormat();
            Alignment alignment5 = new Alignment {
                Horizontal = HorizontalAlignmentValues.Center,
                Vertical = VerticalAlignmentValues.Bottom,
                TextRotation = 0U,
                WrapText = false,
                Indent = 0U,
                JustifyLastLine = false,
                ShrinkToFit = false,
                ReadingOrder = 0U
            };

            differentialFormat4.Append(alignment5);

            differentialFormats1.Append(differentialFormat1);
            differentialFormats1.Append(differentialFormat2);
            differentialFormats1.Append(differentialFormat3);
            differentialFormats1.Append(differentialFormat4);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(differentialFormats1);

            workbookStylesPart.Stylesheet = stylesheet1;
        }
        
        /// <summary>
        /// Generates the Content of the table in the worksheet
        /// </summary>
        /// <param name="worksheetPart">Worksheet to add the content onto</param>
        private void GenerateWorksheetContent(WorksheetPart worksheetPart)
        {
            var worksheet = new Worksheet { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x14ac xr xr2 xr3" }  };
            worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            worksheet.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            worksheet.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            worksheet.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{AB762A7B-1BC4-4BEC-86BE-67D5F0445939}"));

            SheetData sheetData = new SheetData();
            
            for (var y = 0; y < _rows.Count; y++) {
                uint rowIndex = (uint) (y + 1);
                Row row = new Row { RowIndex = rowIndex, Spans = new ListValue<StringValue> { InnerText = "1:2" }, DyDescent = 0.4D };

                for (var x = 0; x < _rows[y].Count; x++) {

                    TableCell tableCell = _rows[y][x];

                    string column = ConvertToAlphabetBasedNumber(x+1);

                    Cell cell = new Cell { CellReference = $"{column}{rowIndex}",
                        StyleIndex = 1U, DataType = tableCell.ValueType, CellValue = tableCell.GetValue()};

                    row.Append(cell);
                }

                sheetData.Append(row);
            }

            TableParts tableParts = new TableParts { Count = 1U };
            TablePart tablePart = new TablePart { Id = "rId2" };

            tableParts.Append(tablePart);
            
            worksheet.Append(sheetData);
            worksheet.Append(tableParts);

            worksheetPart.Worksheet = worksheet;
        }
        
        /// <summary>
        /// Generates Table Definition. 
        /// </summary>
        /// <param name="tableDefinitionPart">Table Definition To add Table to</param>
        private void GenerateTableDefinition(TableDefinitionPart tableDefinitionPart)
        {
            var columnEnd = ConvertToAlphabetBasedNumber(_tableColumns.Count-1); //This function is zero based index
            var rowEnd = _rows.Count + 1; //Need to include the header

            var table = new DocumentFormat.OpenXml.Spreadsheet.Table {
                Id = 1U,
                Name = "Table1",
                DisplayName = "Table1",
                Reference = $"A1:{columnEnd}{rowEnd}",
                TotalsRowShown = false,
                HeaderRowFormatId = 1U,
                DataFormatId = 0U,
                MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "xr xr3" }
            };
            table.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            table.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            table.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{ED5DF9B8-D8E7-4542-9E4F-8E05D7A30983}"));

            AutoFilter autoFilter = new AutoFilter { Reference = $"A1:{columnEnd}{rowEnd}" };
            
            TableColumns tableColumns = new TableColumns { Count = (uint)_tableColumns.Count };

            for (var i = 0; i < tableColumns.Count; i++) {
                tableColumns.Append(new TableColumn
                    {Id = (uint)i, Name = _tableColumns[i], DataFormatId = 2U});
            }

            TableStyleInfo tableStyleInfo1 = new TableStyleInfo { Name = "TableStyleMedium15", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false };

            table.Append(autoFilter);
            table.Append(tableColumns);
            table.Append(tableStyleInfo1);

            tableDefinitionPart.Table = table;
        }
        
        #endregion

        /// <summary>
        /// Convert from decimal number to excel column letter
        /// </summary>
        /// <param name="column">The zero based column number to convert to letter</param>
        /// <returns>The Excel Representation of the column</returns>
        private static string ConvertToAlphabetBasedNumber(int column)
        {
            var dividend = column;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = (char)('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            } 

            return columnName;
        }
    }
}