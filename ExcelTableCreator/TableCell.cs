using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTableCreator
{
    internal enum NumberTypes
    {
        Integer, Double
    }
    
    public class TableCell
    {
        public CellValues ValueType { get; }

        private readonly NumberTypes _numberType;
        
        private readonly int _iVal;
        private readonly double _doubleVal;
        private readonly string _sVal;

        public TableCell(int val)
        {
            ValueType = CellValues.Number;
            _numberType = NumberTypes.Integer;
            _iVal = val;
        }
        
        public TableCell(double val)
        {
            ValueType = CellValues.Number;
            _numberType = NumberTypes.Double;
            _doubleVal = val;
        }

        public TableCell(string val)
        {
            ValueType = CellValues.String;
            _sVal = val;
        }

        public CellValue GetValue()
        {
            return ValueType switch {
                CellValues.Number when _numberType == NumberTypes.Integer => new CellValue(_iVal),
                CellValues.Number when _numberType == NumberTypes.Double => new CellValue(_doubleVal),
                CellValues.String => new CellValue(_sVal),
                _ => new CellValue()
            };
        }
    }
}