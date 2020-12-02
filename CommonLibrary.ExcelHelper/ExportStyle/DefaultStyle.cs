using NPOI.SS.UserModel;
using System;

namespace CommonLibrary.ExcelHelper.ExportStyle
{
    /// <summary>
    /// 默认样式
    /// </summary>
    /// <remarks>
    /// <para>此默认样式</para>
    /// <para>所有列宽均为8字符宽度</para>
    /// <para>数据表表头单元格设置为灰色(#C0C0C0)背景,细实线边框</para>
    /// <para>数据表内容单元格样式无背景颜色,细实线边框</para>
    /// </remarks>
    public class DefaultStyle : IExportStyle
    {
        /// <summary>
        /// 数据表内容单元格样式
        /// </summary>
        protected ICellStyle BodyStyle;

        /// <summary>
        /// 数据表表头单元格样式
        /// </summary>
        protected ICellStyle HeaderStyle;

        /// <summary>
        /// 创建新样式对象代理
        /// </summary>
        public Func<ICellStyle> CreateNewStyle { set; protected get; }

        /// <summary>
        /// 创建新格式的代理
        /// </summary>
        public Func<string, short> CreateNewDataFormat { set; protected get; }

        /// <summary>
        /// 获取数据表内容单元格样式
        /// </summary>
        /// <param name="SheetIndex">工作表索引号，从0开始</param>
        /// <param name="ColumnIndex">列索引号，从0开始</param>
        /// <param name="RowIndex">行号，从0开始</param>
        /// <returns></returns>
        /// <remarks>不包括表头，计算行号时也不计算表头</remarks>
        public ICellStyle GetCellStyle(int SheetIndex, int ColumnIndex, int RowIndex)
        {
            if (BodyStyle == null)
            {
                BodyStyle = CreateNewStyle();
                BodyStyle.BorderBottom = BorderStyle.Thin;
                BodyStyle.BorderLeft = BorderStyle.Thin;
                BodyStyle.BorderRight = BorderStyle.Thin;
                BodyStyle.BorderTop = BorderStyle.Thin;
            }
            return BodyStyle;
        }

        /// <summary>
        /// 获取列宽
        /// </summary>
        /// <param name="SheetIndex">工作表索引号，从0开始</param>
        /// <param name="ColumnIndex">列索引号，从0开始</param>
        /// <returns>字符数量</returns>
        public int GetColumnWidth(int SheetIndex, int ColumnIndex)
        {
            return 8;
        }

        /// <summary>
        /// 获取数据表表头单元格样式
        /// </summary>
        /// <param name="SheetIndex">工作表索引号，从0开始</param>
        /// <param name="ColumnIndex">列索引号，从0开始</param>
        /// <returns></returns>
        public ICellStyle GetHeaderStyle(int SheetIndex, int ColumnIndex)
        {
            if (HeaderStyle == null)
            {
                HeaderStyle = CreateNewStyle();
                HeaderStyle.FillPattern = FillPattern.SolidForeground;
                HeaderStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
                HeaderStyle.BorderBottom = BorderStyle.Thin;
                HeaderStyle.BorderLeft = BorderStyle.Thin;
                HeaderStyle.BorderRight = BorderStyle.Thin;
                HeaderStyle.BorderTop = BorderStyle.Thin;
            }
            return HeaderStyle;
        }
    }
}