using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelManager {
    public class ExcelReader<T> where T : class, new() {

        private Stream _stream;

        private IWorkbook _workbook;

        private ISheet _currentSheet;

        private IEnumerable<PropertyInfo> _properties;

        //public IWorkbook WorkBook
        //{
        //    get
        //    {
        //        InitializeWorkBook();
        //        return _workbook;
        //    }
        //}

        private void InitializeProperties() {

            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                    .Where(p => p.GetCustomAttributes(typeof(ExcelReaderCellAttribute), true).Length > 0)
                                    .Select(p => p);

            if (!properties.Any()) {
                throw new ExcelCellReaderAttributeException($"Public instance property with {nameof(ExcelReaderCellAttribute)} not found on {typeof(T).Name}");
            }
            else
                _properties = properties;
        }

        public ExcelReader(Stream stream) {

            ArgumentNotNull(stream, nameof(stream));

            _stream = stream;
        }

        public ExcelReader(string filePath) {

            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentNullException(nameof(filePath));
            }

            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)) {

                _stream = new MemoryStream();

                file.CopyTo(_stream);
                _stream.Position = 0;
            }
        }

        private void InitializeWorkBook() {

            if (_workbook == null)
                _workbook = WorkbookFactory.Create(_stream);
        }

        public IEnumerable<T> Read(bool includeHeaders = false) {
            InitializeProperties();
            InitializeWorkBook();

            var sheets = new ISheet[_workbook.NumberOfSheets];

            for (int i = 0; i < _workbook.NumberOfSheets; i++)
                sheets[i] = _workbook.GetSheetAt(i);

            return ReadSheets(sheets, includeHeaders);
        }

        private IEnumerable<T> ReadSheets(ISheet[] sheets, bool includeHeaders) {

            var result = new List<T>();

            Array.ForEach(sheets, sheet => {

                _currentSheet = sheet;

                var rowsInSheet = sheet.GetRowEnumerator();

                while (rowsInSheet.MoveNext()) {

                    var row = (IRow)rowsInSheet.Current;

                    if ((row == null) || (row.RowNum == 0 && !includeHeaders))
                        continue;

                    var item = ReadRow(row);

                    if (item != null)
                        result.Add(item);
                }
            });
            _currentSheet = null;
            return result;
        }

        private T ReadRow(IRow row) {

            var instance = new T();

            var currentCells = row.Cells;

            foreach (var property in _properties) {

                var propertyAttribute = property.GetCustomAttributes<ExcelReaderCellAttribute>().FirstOrDefault();

                ICell cell;

                if (propertyAttribute.Index != null) {
                    if (propertyAttribute.Index > currentCells.Count)//catches outbound index...
                        continue;

                    cell = currentCells[propertyAttribute.Index.Value];
                }

                else {

                    cell = currentCells.FirstOrDefault((x) => {

                        if (x == null)
                            return false;

                        else {
                            var header = _currentSheet.GetRow(_currentSheet.FirstRowNum).Cells[x.ColumnIndex];

                            return header.StringCellValue
                                    .Equals(propertyAttribute.Name ?? property.Name, StringComparison.InvariantCultureIgnoreCase);
                        }
                    });
                }

                if (cell != null) {
                    try {
                        var cellType = cell.CellType;

                        switch (cellType) {

                            case CellType.Numeric:
                                var dateOrNumeric = DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue.ToString() : cell.NumericCellValue.ToString();
                                var value = ConvertToNumeric(dateOrNumeric, property.PropertyType);
                                property.SetValue(instance, value);
                                break;

                            case CellType.String:
                                property.SetValue(instance, cell.StringCellValue.Trim());
                                break;

                            case CellType.Formula:
                                property.SetValue(instance, cell.CellFormula.Trim());
                                break;

                            case CellType.Boolean:
                                property.SetValue(instance, cell.BooleanCellValue);
                                break;

                            case CellType.Blank:
                            case CellType.Unknown:
                            default:
                                property.SetValue(instance, null);
                                break;
                        }
                    }

                    catch (Exception e) when (e is ArgumentException) {//usually type conversion

                        var msg = $"Invalid conversion in Cell [{cell.RowIndex}, {cell.ColumnIndex}]";
                        throw new CellValueConvertionException(msg);
                    }

                    catch (Exception e) {
                        throw e;
                    }
                }
            }

            return instance;
        }

        private object ConvertToNumeric(string DateOrNumeric, Type type) {

            var dtype = IsNullable(type) ? Nullable.GetUnderlyingType(type) : type;
            return Convert.ChangeType(DateOrNumeric, dtype);
        }

        private static bool IsNullable(Type type) {
            return type.IsGenericType && type.GetGenericTypeDefinition().Equals(typeof(Nullable<>));
        }

        private void ArgumentNotNull(object value, string name) {

            if (value == null) {

                throw new ArgumentNullException(name);
            }
        }
    }
}