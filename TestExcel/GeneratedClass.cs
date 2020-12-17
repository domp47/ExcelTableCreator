using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Op = DocumentFormat.OpenXml.CustomProperties;

// ReSharper disable PossiblyMistakenUseOfParamsMethod
namespace TestExcel
{
    public class GeneratedClass
    {
        // Creates a SpreadsheetDocument.
        public void CreatePackage(string filePath)
        {
            using SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            CreateParts(package);
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(SpreadsheetDocument document)
        {
            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPart1Content(workbookStylesPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPart1Content(worksheetPart1);

            TableDefinitionPart tableDefinitionPart1 = worksheetPart1.AddNewPart<TableDefinitionPart>("rId2");
            GenerateTableDefinitionPart1Content(tableDefinitionPart1);
        }

        // Generates content of workbookPart1.
        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook { MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "x15 xr xr6 xr10 xr2" }  };
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relation ships");
            workbook1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            workbook1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            workbook1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            workbook1.AddNamespaceDeclaration("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6");
            workbook1.AddNamespaceDeclaration("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10");
            workbook1.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet { Name = "Report", SheetId = 1U, Id = "rId1" };
            
            sheets1.Append(sheet1);
            
            workbook1.Append(sheets1);

            workbookPart1.Workbook = workbook1;
        }

        // Generates content of workbookStylesPart1.
        private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet(){ MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "x14ac x16r2 xr" }  };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
            stylesheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

            Fonts fonts1 = new Fonts(){ Count = (UInt32Value)1U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize(){ Val = 11D };
            Color color1 = new Color(){ Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName(){ Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering(){ Val = 2 };
            FontScheme fontScheme1 = new FontScheme(){ Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            fonts1.Append(font1);

            Fills fills1 = new Fills(){ Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill(){ PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill(){ PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders(){ Count = (UInt32Value)1U };

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


            CellFormats cellFormats1 = new CellFormats(){ Count = (UInt32Value)2U };
            CellFormat cellFormat2 = new CellFormat(){ NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };

            CellFormat cellFormat3 = new CellFormat(){ NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment1 = new Alignment(){ Horizontal = HorizontalAlignmentValues.Center };

            cellFormat3.Append(alignment1);

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);

            DifferentialFormats differentialFormats1 = new DifferentialFormats(){ Count = (UInt32Value)4U };

            DifferentialFormat differentialFormat1 = new DifferentialFormat();
            Alignment alignment2 = new Alignment(){ Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, TextRotation = (UInt32Value)0U, WrapText = false, Indent = (UInt32Value)0U, JustifyLastLine = false, ShrinkToFit = false, ReadingOrder = (UInt32Value)0U };

            differentialFormat1.Append(alignment2);

            DifferentialFormat differentialFormat2 = new DifferentialFormat();
            Alignment alignment3 = new Alignment(){ Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, TextRotation = (UInt32Value)0U, WrapText = false, Indent = (UInt32Value)0U, JustifyLastLine = false, ShrinkToFit = false, ReadingOrder = (UInt32Value)0U };

            differentialFormat2.Append(alignment3);

            DifferentialFormat differentialFormat3 = new DifferentialFormat();
            Alignment alignment4 = new Alignment(){ Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, TextRotation = (UInt32Value)0U, WrapText = false, Indent = (UInt32Value)0U, JustifyLastLine = false, ShrinkToFit = false, ReadingOrder = (UInt32Value)0U };

            differentialFormat3.Append(alignment4);

            DifferentialFormat differentialFormat4 = new DifferentialFormat();
            Alignment alignment5 = new Alignment(){ Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, TextRotation = (UInt32Value)0U, WrapText = false, Indent = (UInt32Value)0U, JustifyLastLine = false, ShrinkToFit = false, ReadingOrder = (UInt32Value)0U };

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

            workbookStylesPart1.Stylesheet = stylesheet1;
        }
        
        // Generates content of worksheetPart1.
        private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = new Worksheet(){ MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "x14ac xr xr2 xr3" }  };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            worksheet1.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            worksheet1.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            worksheet1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{AB762A7B-1BC4-4BEC-86BE-67D5F0445939}"));
            
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties(){ DefaultRowHeight = 14.6D, DyDescent = 0.4D };

            SheetData sheetData1 = new SheetData();

            Row row1 = new Row(){ RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.4D };

            Cell cell1 = new Cell(){ CellReference = "A1", StyleIndex = (UInt32Value)1U, DataType = CellValues.String, CellValue = new CellValue("Name")};

            Cell cell2 = new Cell(){ CellReference = "B1", StyleIndex = (UInt32Value)1U, DataType = CellValues.String, CellValue = new CellValue("Val")};

            row1.Append(cell1);
            row1.Append(cell2);

            Row row2 = new Row(){ RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.4D };

            Cell cell3 = new Cell(){ CellReference = "A2", StyleIndex = (UInt32Value)1U, DataType = CellValues.String, CellValue = new CellValue("Yeeetus")};

            Cell cell4 = new Cell(){ CellReference = "B2", StyleIndex = (UInt32Value)1U, DataType = CellValues.Number, CellValue = new CellValue(69)};

            row2.Append(cell3);
            row2.Append(cell4);

            sheetData1.Append(row1);
            sheetData1.Append(row2);
            
            PageSetup pageSetup1 = new PageSetup(){ Orientation = OrientationValues.Portrait, Id = "rId1" };

            TableParts tableParts1 = new TableParts(){ Count = (UInt32Value)1U };
            TablePart tablePart1 = new TablePart(){ Id = "rId2" };

            tableParts1.Append(tablePart1);
            
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(pageSetup1);
            worksheet1.Append(tableParts1);

            worksheetPart1.Worksheet = worksheet1;
        }

        // Generates content of tableDefinitionPart1.
        private void GenerateTableDefinitionPart1Content(TableDefinitionPart tableDefinitionPart1)
        {
            Table table1 = new Table(){ Id = (UInt32Value)1U, Name = "Table1", DisplayName = "Table1", Reference = "A1:B7", TotalsRowShown = false, HeaderRowFormatId = (UInt32Value)1U, DataFormatId = (UInt32Value)0U, MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "xr xr3" }  };
            table1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            table1.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            table1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{ED5DF9B8-D8E7-4542-9E4F-8E05D7A30983}"));

            AutoFilter autoFilter1 = new AutoFilter(){ Reference = "A1:B7" };

            TableColumns tableColumns1 = new TableColumns(){ Count = (UInt32Value)2U };

            TableColumn tableColumn1 = new TableColumn(){ Id = (UInt32Value)1U, Name = "Name", DataFormatId = (UInt32Value)3U };

            TableColumn tableColumn2 = new TableColumn(){ Id = (UInt32Value)2U, Name = "Val", DataFormatId = (UInt32Value)2U };

            tableColumns1.Append(tableColumn1);
            tableColumns1.Append(tableColumn2);
            TableStyleInfo tableStyleInfo1 = new TableStyleInfo(){ Name = "TableStyleLight8", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false };

            table1.Append(autoFilter1);
            table1.Append(tableColumns1);
            table1.Append(tableStyleInfo1);

            tableDefinitionPart1.Table = table1;
        }
    }
}
