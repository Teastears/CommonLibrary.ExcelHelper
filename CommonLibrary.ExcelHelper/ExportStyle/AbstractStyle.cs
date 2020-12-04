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
    /// <para>所有行高为14.75</para> 
    /// </remarks>
    public abstract class AbstractStyle : IExportStyle
    {
 

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
        public virtual ICellStyle GetCellStyle(int SheetIndex, int ColumnIndex, int RowIndex)
        {
            return null;
        }

        /// <summary>
        /// 获取列宽
        /// </summary>
        /// <param name="SheetIndex">工作表索引号，从0开始</param>
        /// <param name="ColumnIndex">列索引号，从0开始</param>
        /// <returns>字符数量</returns>
        public virtual int GetColumnWidth(int SheetIndex, int ColumnIndex)
        {
            return 8;
        }

        /// <summary>
        /// 获取数据表表头单元格样式
        /// </summary>
        /// <param name="SheetIndex">工作表索引号，从0开始</param>
        /// <param name="ColumnIndex">列索引号，从0开始</param>
        /// <returns></returns>
        public virtual ICellStyle GetHeaderStyle(int SheetIndex, int ColumnIndex)
        {
            return null;
        }
        /// <summary>
        /// 获取行高
        /// </summary>
        /// <param name="SheetIndex">工作表索引号，从0开始</param>
        /// <param name="RowIndex">行号，从0开始</param>
        /// <returns></returns>
        public virtual float GetRowHeigth(int SheetIndex, int RowIndex)
        {
            return 14.25F;
        }
    }
}