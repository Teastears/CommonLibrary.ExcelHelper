using CommonLibrary.ExcelHelper.Enum;
using CommonLibrary.ExcelHelper.Export;
using CommonLibrary.ExcelHelper.Import;
using CommonLibrary.ExcelHelper.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace CommonLibrary.ExcelHelper
{
    /// <summary>
    /// Excel导入操出类创建工厂
    /// </summary>
    public static class ExcelHelperFactory
    {
        /// <summary>
        /// 创建导出处理对象
        /// </summary>
        /// <param name="SourceData">数据源</param>
        /// <param name="ExcelVersion">导出数据的Excel版本</param>
        /// <param name="SheetName">工作表名称列表</param>
        /// <returns></returns>
        public static DataSetExporter CreateExporter(DataSet SourceData, ExcelVersion ExcelVersion = ExcelVersion.XLSX, List<string> SheetName = null)
        {
            var Helper = new DataSetExporter
            {
                SourceData = SourceData,
                ExcelVersion = ExcelVersion,
                SheetName = SheetName
            };
            return Helper;
        }

        /// <summary>
        /// 创建导出处理对象
        /// </summary>
        /// <param name="SourceData">数据源</param>
        /// <param name="SheetName">工作表名称列表</param>
        /// <returns></returns>
        public static DataSetExporter CreateExporter(DataSet SourceData, List<string> SheetName)
        {
            return CreateExporter(SourceData, ExcelVersion.XLSX, SheetName);
        }

        /// <summary>
        /// 创建导出处理对象
        /// </summary>
        /// <param name="SourceData">数据源</param>
        /// <param name="ExcelVersion">导出数据的Excel版本</param>
        /// <param name="SheetName">工作表名称</param>
        /// <returns></returns>
        public static DataTableExporter CreateExporter(DataTable SourceData, ExcelVersion ExcelVersion = ExcelVersion.XLSX, string SheetName = "")
        {
            var Helper = new DataTableExporter
            {
                SourceData = SourceData,
                ExcelVersion = ExcelVersion,
                SheetName = SheetName
            };
            return Helper;
        }

        /// <summary>
        /// 创建导出处理对象
        /// </summary>
        /// <param name="SourceData">数据源</param>
        /// <param name="SheetName">工作表名称列表</param>
        /// <returns></returns>
        public static DataTableExporter CreateExporter(DataTable SourceData, string SheetName)
        {
            return CreateExporter(SourceData, ExcelVersion.XLSX, SheetName);
        }

        /// <summary>
        /// 创建导出处理对象
        /// </summary>
        /// <param name="SourceData">数据源</param>
        /// <param name="ExcelVersion">导出数据的Excel版本</param>
        /// <param name="SheetName">工作表名称</param>
        /// <param name="ValueProvidor">导出数据时的数据提供者</param>
        /// <returns></returns>
        public static IEnumerableExporter<T> CreateExporter<T>(IEnumerable<T> SourceData, ExcelVersion ExcelVersion = ExcelVersion.XLSX, string SheetName = "", Func<string, T, object> ValueProvidor = null)
        {
            var Helper = new IEnumerableExporter<T>
            {
                SourceData = SourceData,
                ExcelVersion = ExcelVersion,
                SheetName = SheetName,
                ValueProvidor = ValueProvidor
            };
            return Helper;
        }

        /// <summary>
        /// 创建导出处理对象
        /// </summary>
        /// <param name="SourceData">数据源</param>
        /// <param name="SheetName">工作表名称列表</param>
        /// <returns></returns>
        public static IEnumerableExporter<T> CreateExporter<T>(IEnumerable<T> SourceData, string SheetName)
        {
            return CreateExporter(SourceData, ExcelVersion.XLSX, SheetName);
        }

        /// <summary>
        /// 创建导入处理对象
        /// </summary>
        /// <param name="DataSourceFilePath">要导入的Excel文件路径</param>
        /// <param name="ImportSheets">导入配置列表</param>
        /// <returns></returns>
        public static ExcelFileImporter CreateImporter(string DataSourceFilePath, List<ImportSheetSetting> ImportSheets = null)
        {
            var Helper = new ExcelFileImporter
            {
                DataSourceFilePath = DataSourceFilePath,
                ImportSheets = ImportSheets
            };
            return Helper;
        }

        /// <summary>
        /// 创建导入处理对象
        /// </summary>
        /// <param name="DataSource">要导入的Excel文件流</param>
        /// <param name="ImportSheets">导入配置列表</param>
        /// <returns></returns>
        public static ExcelStreamImporter CreateImporter(Stream DataSource, List<ImportSheetSetting> ImportSheets = null)
        {
            var Helper = new ExcelStreamImporter
            {
                SourceData = DataSource,
                ImportSheets = ImportSheets
            };
            return Helper;
        }
    }
}