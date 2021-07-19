using CommonLibrary.ExcelHelper.Enum;
using CommonLibrary.ExcelHelper.Model;
using CommonLibrary.ExcelHelper.Model.Exception;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;

namespace CommonLibrary.ExcelHelper.Import
{
    /// <summary>
    /// Excel文件流导入处理类
    /// </summary>
    public class ExcelStreamImporter : IExcelImporter, IDisposable
    {
        /// <summary>
        /// 根据要导入的数据源创建的工作簿
        /// </summary>
        protected IWorkbook WorkBook;

        /// <summary>
        /// 创建工作簿(依据文件流)
        /// </summary>
        /// <returns></returns>
        protected virtual void CreateWorkbook()
        {
            if (SourceData == null)
                throw new NoDataException();
            if (ExcelVersion == ExcelVersion.XLS)
            {
                WorkBook = new HSSFWorkbook(SourceData);
            }
            else
            {
                WorkBook = new XSSFWorkbook(SourceData);
            }
        }

        /// <summary>
        /// 从工作表中生成List
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="headerRowIndex">表头行号</param>
        /// <returns></returns>
        protected List<T> GetDataListFromSheet<T>(ISheet sheet, int headerRowIndex) where T : class
        {
            Type type = typeof(T);
            List<T> table = new List<T>();
            IRow headerRow = sheet.GetRow(headerRowIndex);
            if (headerRow == null || string.IsNullOrEmpty(headerRow.Cells[0].StringCellValue))
                return table;
            int cellCount = headerRow.LastCellNum;
            List<string> FieldDic = new List<string>();
            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                if (headerRow.GetCell(i) == null || headerRow.GetCell(i).StringCellValue.Trim() == "")
                {
                    // 如果遇到第一个空列，则不再继续向后读取
                    cellCount = i + 1;
                    break;
                }
                FieldDic.Add(headerRow.GetCell(i).StringCellValue);
            }
            for (int i = (headerRowIndex + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    T dataRow = type.Assembly.CreateInstance(type.FullName, false) as T;
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        var Property = type.GetProperty(FieldDic[j], BindingFlags.Public | BindingFlags.Instance);

                        var cell = row.GetCell(j);
                        if (cell != null)
                        {
                            string cellvaluestr;
                            if (cell.CellType == CellType.Numeric)
                            {
                                if (DateUtil.IsCellDateFormatted(cell))
                                {
                                    cellvaluestr = DateTime.FromOADate(cell.NumericCellValue).ToString();
                                }
                                else
                                {
                                    cellvaluestr = cell.NumericCellValue.ToString();
                                }
                            }
                            else
                            {
                                cellvaluestr = cell.StringCellValue.ToString();
                            }
                            if (string.IsNullOrWhiteSpace(cellvaluestr))
                                continue;
                            Type PropertyType;

                            if (Property.PropertyType.IsGenericType && Property.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                            {
                                PropertyType = Property.PropertyType.GetGenericArguments()[0];
                            }
                            else
                            {
                                PropertyType = Property.PropertyType;
                            }
                            if (PropertyType == typeof(DateTime))
                            {
                                Property.SetValue(dataRow, cell.DateCellValue);
                            }
                            else
                            {
                                Property.SetValue(dataRow, Convert.ChangeType(cellvaluestr, PropertyType));
                            }
                        }
                    }

                    table.Add(dataRow);
                }
            }

            return table;
        }

        /// <summary>
        /// 从工作表中生成DataTable
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="headerRowIndex">表头行号</param>
        /// <returns></returns>
        protected DataTable GetDataTableFromSheet(ISheet sheet, int headerRowIndex)
        {
            DataTable table = new DataTable
            {
                TableName = sheet.SheetName
            };
            IRow headerRow = sheet.GetRow(headerRowIndex);
            if (headerRow == null || string.IsNullOrEmpty(headerRow.Cells[0].StringCellValue))
                return table;
            int cellCount = headerRow.LastCellNum;
            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                if (headerRow.GetCell(i) == null || headerRow.GetCell(i).StringCellValue.Trim() == "")
                {
                    // 如果遇到第一个空列，则不再继续向后读取
                    cellCount = i + 1;
                    break;
                }
                DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(column);
            }
            for (int i = (headerRowIndex + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    if (row.Cells.Count > 0)
                    {
                        DataRow dataRow = table.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            var cell = row.GetCell(j);
                            if (cell != null)
                            {
                                if (cell.CellType == CellType.Numeric)
                                {
                                    if (DateUtil.IsCellDateFormatted(cell))
                                    {
                                        dataRow[j] = DateTime.FromOADate(cell.NumericCellValue);
                                    }
                                    else
                                    {
                                        dataRow[j] = cell.NumericCellValue;
                                    }
                                }
                                else
                                {
                                    dataRow[j] = cell.StringCellValue;
                                }
                            }
                        }
                        table.Rows.Add(dataRow);
                    }
                }
            }

            return table;
        }

        /// <summary>
        /// 初始化导入配置
        /// </summary>
        protected void InitImportSheets()
        {
            if (ImportSheets == null)
                ImportSheets = new List<ImportSheetSetting>();
            if (ImportSheets.Count == 0)
            {
                for (int i = 0; i < WorkBook.NumberOfSheets; i++)
                {
                    ImportSheets.Add(new ImportSheetSetting(i));
                }
            }
            else
            {
                foreach (var item in ImportSheets)
                {
                    if (!item.SheetIndex.HasValue)
                    {
                        item.SheetIndex = WorkBook.GetSheetIndex(item.SheetName);
                    }
                }
            }
        }

        /// <summary>
        /// Excel版本，默认为XLSX
        /// </summary>
        public ExcelVersion ExcelVersion { get; set; } = ExcelVersion.XLSX;

        /// <summary>
        /// 导入的数据
        /// </summary>
        public DataSet ImportData { get; protected set; }

        /// <summary>
        /// 要导入的数据源导入配置列表
        /// </summary>
        public List<ImportSheetSetting> ImportSheets { get; set; }

        /// <summary>
        /// 要导入的数据源
        /// </summary>
        public Stream SourceData { get; set; }

        /// <summary>
        /// 释放
        /// </summary>
        public virtual void Dispose()
        {
            WorkBook.Close();
            ImportData.Dispose();
        }

        /// <summary>
        /// 导入数据
        /// </summary>
        /// <returns></returns>
        public DataSet Import()
        {
            CreateWorkbook();
            InitImportSheets();
            ImportData = new DataSet();
            foreach (var item in ImportSheets)
            {
                ISheet sheet = WorkBook.GetSheetAt(item.SheetIndex.Value);
                DataTable table = GetDataTableFromSheet(sheet, item.HeaderRowIndex);
                ImportData.Tables.Add(table);
            }

            return ImportData;
        }

        /// <summary>
        /// 导入数据
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <returns></returns>
        public List<T> Import<T>() where T : class
        {
            CreateWorkbook();
            InitImportSheets();
            ISheet sheet = WorkBook.GetSheetAt(ImportSheets[0].SheetIndex.Value);
            return GetDataListFromSheet<T>(sheet, ImportSheets[0].HeaderRowIndex);
        }
    }
}