using CommonLibrary.ExcelHelper.Base;
using CommonLibrary.ExcelHelper.ExportStyle;
using CommonLibrary.ExcelHelper.Model;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Data;

namespace CommonLibrary.ExcelHelper.Export
{
    /// <summary>
    /// DataTable导出处理类
    /// </summary>
    public class DataTableExporter : AbstractExcelExporter
    {
        /// <summary>
        /// 数据验证
        /// </summary>
        /// <returns></returns>
        protected override bool DataCheck()
        {
            return SourceData.Rows.Count > 0;
        }

        /// <summary>
        /// 初始化表头显示名称
        /// </summary>
        protected void InitHeaderNames()
        {
            if (HeaderNames != null)
                return;
            HeaderNames = new List<KeyValuePair<string, string>>();

            foreach (DataColumn item in SourceData.Columns)
            {
                HeaderNames.Add(new KeyValuePair<string, string>(item.ColumnName, item.ColumnName));
            }
        }

        /// <summary>
        /// 初始化工作表名称列表
        /// </summary>
        protected void InitSheetName()
        {
            if (!string.IsNullOrWhiteSpace(SheetName))
                return;
            SheetName = string.IsNullOrWhiteSpace(SourceData.TableName) ? "Sheet1" : SourceData.TableName;
        }

        /// <summary>
        /// 导出后的表头显示配置
        /// </summary>
        public List<KeyValuePair<string, string>> HeaderNames { get; set; }

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 要导出的数据源
        /// </summary>
        public DataTable SourceData { get; set; }

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
            CreateWorkbook();
            ISheet sheet = Workbook.CreateSheet(SheetName);
            IRow headerRow = sheet.CreateRow(0);
            if (ExportStyle != null)
            {
                ExportStyle.CreateNewStyle = CreateNewStyle;
                ExportStyle.CreateNewDataFormat = CreateNewDataFormat;
            }
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
            foreach (DataRow row in SourceData.Rows)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);
                for (int n = 0; n < HeaderNames.Count; n++)
                {
                    var cell = dataRow.CreateCell(n);
                    SetCellValue(cell, row[HeaderNames[n].Key]);
                    if (ExportStyle != null)
                        cell.CellStyle = ExportStyle.GetCellStyle(0, n, rowIndex - 1);
                }
                rowIndex++;
            }
            NPOIMemoryStream Stream = new NPOIMemoryStream()
            {
                AllowClose = false
            };
            Workbook.Write(Stream);
            Workbook.Close();
            Stream.AllowClose = true;
            return Stream;
        }
    }
}