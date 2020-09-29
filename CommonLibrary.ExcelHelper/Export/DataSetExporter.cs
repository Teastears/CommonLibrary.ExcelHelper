using CommonLibrary.ExcelHelper.Base;
using CommonLibrary.ExcelHelper.ExportStyle;
using CommonLibrary.ExcelHelper.Model;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Data;

namespace CommonLibrary.ExcelHelper.Export
{
    /// <summary>
    /// DataSet导出处理类
    /// </summary>
    public class DataSetExporter : AbstractExcelExporter
    {
        /// <summary>
        /// 工作表名称列表
        /// </summary>
        protected List<string> _SheetName;

        /// <summary>
        /// 数据验证
        /// </summary>
        /// <returns></returns>
        protected override bool DataCheck()
        {
            int count = 0;
            for (int i = 0; i < SourceData.Tables.Count; i++)
            {
                count += SourceData.Tables[i].Rows.Count;
            }
            return count > 0;
        }

        /// <summary>
        /// 初始化表头显示名称
        /// </summary>
        protected void InitHeaderNames()
        {
            if (HeaderNames != null)
                return;
            HeaderNames = new List<KeyValuePair<string, List<KeyValuePair<string, string>>>>();

            foreach (var Sheet in SheetName)
            {
                var SubHeaderNames = new List<KeyValuePair<string, string>>();
                foreach (DataColumn Column in SourceData.Tables[Sheet].Columns)
                {
                    SubHeaderNames.Add(new KeyValuePair<string, string>(Column.ColumnName, Column.ColumnName));
                }
                HeaderNames.Add(new KeyValuePair<string, List<KeyValuePair<string, string>>>(Sheet, SubHeaderNames));
            }
        }

        /// <summary>
        /// 初始化工作表名称列表
        /// </summary>
        protected void InitSheetName()
        {
            int i = 0;
            var NewSheetName = new List<string>();
            if (_SheetName != null)
            {
                for (; i < _SheetName.Count; i++)
                {
                    NewSheetName.Add(_SheetName[i]);
                }
            }
            for (; i < SourceData.Tables.Count; i++)
            {
                NewSheetName.Add(SourceData.Tables[i].TableName);
            }
            _SheetName = NewSheetName;
        }

        /// <summary>
        /// 导出后的表头显示配置
        /// </summary>
        public List<KeyValuePair<string, List<KeyValuePair<string, string>>>> HeaderNames { get; set; }

        /// <summary>
        /// 工作表名称列表
        /// </summary>
        public List<string> SheetName
        {
            get
            {
                return _SheetName;
            }
            set
            {
                _SheetName = new List<string>();
                if (value != null)
                {
                    foreach (var item in value)
                    {
                        _SheetName.Add(item);
                    }
                }
            }
        }

        /// <summary>
        /// 要导出的数据源
        /// </summary>
        public DataSet SourceData { get; set; }

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
            for (int j = 0; j < HeaderNames.Count; j++)
            {
                DataTable table = SourceData.Tables[j];
                ISheet sheet = Workbook.CreateSheet(SheetName[j]);
                IRow headerRow = sheet.CreateRow(0);
                var SubHeaderNames = HeaderNames[j].Value;
                for (int i = 0; i < SubHeaderNames.Count; i++)
                {
                    if (ExportStyle != null)
                        sheet.SetColumnWidth(i, ExcelBase.ColumnWidth(ExportStyle.GetColumnWidth(j, i)));
                    ICell cell = headerRow.CreateCell(i);
                    cell.SetCellValue(SubHeaderNames[i].Value);
                    if (ExportStyle != null)
                        cell.CellStyle = ExportStyle.GetHeaderStyle(j, i);
                }
                int rowIndex = 1;
                foreach (DataRow row in table.Rows)
                {
                    IRow dataRow = sheet.CreateRow(rowIndex);
                    for (int n = 0; n < SubHeaderNames.Count; n++)
                    {
                        var cell = dataRow.CreateCell(n);
                        SetCellValue(cell, row[SubHeaderNames[n].Key]);
                        if (ExportStyle != null)
                            cell.CellStyle = ExportStyle.GetCellStyle(j, n, rowIndex - 1);
                    }
                    rowIndex++;
                }
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
    }
}