using CommonLibrary.ExcelHelper.Base;
using CommonLibrary.ExcelHelper.ExportStyle;
using CommonLibrary.ExcelHelper.Model;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace CommonLibrary.ExcelHelper.Export
{
    /// <summary>
    /// IEnumerable&lt;T&gt;导出处理类
    /// </summary>
    public class IEnumerableExporter<T> : AbstractExcelExporter
    {
        /// <summary>
        /// 数据验证
        /// </summary>
        /// <returns></returns>
        protected override bool DataCheck()
        {
            return SourceData.Count() > 0;
        }

        /// <summary>
        /// 初始化表头显示名称
        /// </summary>
        protected void InitHeaderNames()
        {
            if (HeaderNames != null)
                return;
            var Type = typeof(T);
            var Properties = Type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            HeaderNames = new List<KeyValuePair<string, string>>();

            foreach (var item in Properties)
            {
                HeaderNames.Add(new KeyValuePair<string, string>(item.Name, item.Name));
            }
        }

        /// <summary>
        /// 初始化工作表名称列表
        /// </summary>
        protected void InitSheetName()
        {
            if (!string.IsNullOrWhiteSpace(SheetName))
                return;
            SheetName = "Sheet1";
        }

        /// <summary>
        /// 导出后的表头显示配置
        /// </summary>
        public List<KeyValuePair<string, string>> HeaderNames { get; set; }
        /// <summary>
        /// 导出数据时的数据提供者
        /// </summary>
        public Func<string, T, object> ValueProvidor { get; set; }

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 要导出的数据源
        /// </summary>
        public IEnumerable<T> SourceData { get; set; }

        /// <summary>
        /// 导出到内存流
        /// </summary>
        /// <param name="ExportStyle">导出时应用的样式</param>
        /// <returns></returns>
        public override NPOIMemoryStream ExportToStream(IExportStyle ExportStyle = null)
        {
            DataCheck();
            InitSheetName();
            InitHeaderNames();
            if (ValueProvidor == null)
                ValueProvidor = DefaultValueProvidor;
            CreateWorkbook();
            if (ExportStyle != null)
            {
                ExportStyle.CreateNewStyle = CreateNewStyle;
                ExportStyle.CreateNewDataFormat = CreateNewDataFormat;
            }
            ISheet sheet = Workbook.CreateSheet(SheetName);
            IRow headerRow = sheet.CreateRow(0);

            for (int i = 0; i < HeaderNames.Count; i++)
            {
                if (ExportStyle != null)
                    sheet.SetColumnWidth(i, ExcelBase.ColumnWidth(ExportStyle.GetColumnWidth(0, i)));
                ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(HeaderNames[i].Value);
                if (ExportStyle != null)
                    cell.CellStyle = ExportStyle.GetHeaderStyle(0, i);
            }

            int rowIndex = 1;
            foreach (T item in SourceData)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);
                for (int n = 0; n < HeaderNames.Count; n++)
                {
                    object pValue = ValueProvidor(HeaderNames[n].Key, item);
                    var cell = dataRow.CreateCell(n);
                    SetCellValue(cell, pValue);
                    if (ExportStyle != null)
                        cell.CellStyle = ExportStyle.GetCellStyle(0, n, rowIndex - 1);
                }
                rowIndex++;
            }
            NPOIMemoryStream Stream = new NPOIMemoryStream
            {
                AllowClose = false
            };
            Workbook.Write(Stream);
            Workbook.Close();
            Stream.AllowClose = true;
            return Stream;
        }

        private object DefaultValueProvidor(string PropertyKey, T obj)
        {
            Type t = typeof(T);
            object pValue = t.GetProperty(PropertyKey).GetValue(obj, null);
            return pValue;
        }
    }
}