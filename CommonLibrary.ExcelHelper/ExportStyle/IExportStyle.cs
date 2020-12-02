using NPOI.SS.UserModel;
using System;

namespace CommonLibrary.ExcelHelper.ExportStyle
{
    /// <summary>
    /// 导出样式
    /// </summary>
    public interface IExportStyle
    {
        /// <summary>
        /// 创建新样式对象代理
        /// </summary>
        Func<ICellStyle> CreateNewStyle { set; }

        /// <summary>
        /// 创建新格式的代理
        /// </summary>
        Func<string, short> CreateNewDataFormat { set; }

        /// <summary>
        /// 获取数据表内容单元格样式
        /// </summary>
        /// <param name="SheetIndex">工作表索引号，从0开始</param>
        /// <param name="ColumnIndex">列索引号，从0开始</param>
        /// <param name="RowIndex">行号，从0开始</param>
        /// <returns></returns>
        /// <remarks>不包括表头，计算行号时也不计算表头</remarks>
        ICellStyle GetCellStyle(int SheetIndex, int ColumnIndex, int RowIndex);

        /// <summary>
        /// 获取列宽
        /// </summary>
        /// <param name="SheetIndex">工作表索引号，从0开始</param>
        /// <param name="ColumnIndex">列索引号，从0开始</param>
        /// <returns>字符数量</returns>
        int GetColumnWidth(int SheetIndex, int ColumnIndex);

        /// <summary>
        /// 获取数据表表头单元格样式
        /// </summary>
        /// <param name="SheetIndex">工作表索引号，从0开始</param>
        /// <param name="ColumnIndex">列索引号，从0开始</param>
        /// <returns></returns>
        ICellStyle GetHeaderStyle(int SheetIndex, int ColumnIndex);
    }
}