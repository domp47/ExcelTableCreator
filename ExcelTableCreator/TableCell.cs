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
        
        private readonly bool _bVal;
        private readonly int _iVal;
        private readonly double _doubleVal;
        private readonly string _sVal;
        private readonly DateTime _dVal;

        public TableCell(bool val)
        {
            ValueType = CellValues.Boolean;
            _bVal = val;
        }

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

        public TableCell(DateTime val)
        {
            ValueType = CellValues.Date;
            _dVal = val;
        }

        public CellValue GetValue()
        {
            return ValueType switch {
                CellValues.Boolean => new CellValue(_bVal),
                CellValues.Number when _numberType == NumberTypes.Integer => new CellValue(_iVal),
                CellValues.Number when _numberType == NumberTypes.Double => new CellValue(_doubleVal),
                CellValues.String => new CellValue(_sVal),
                CellValues.Date => new CellValue(_dVal),
                _ => new CellValue()
            };
        }
    }
}