using CommonLibrary.ExcelHelper.Enum;
using CommonLibrary.ExcelHelper.Model.Exception;
using System;

namespace CommonLibrary.ExcelHelper.Base
{
    internal static class ExcelBase
    {
        /// <summary>
        /// 计算列宽
        /// </summary>
        /// <param name="Width">要设置的列宽，单位字符</param>
        /// <returns></returns>
        public static int ColumnWidth(int Width)
        {
            return Width * 256 + 200;
        } 

        /// <summary>
        /// 判断是否为2007以前格式
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static ExcelVersion GetVersion(string filePath)
        {
            if (filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            {
                return ExcelVersion.XLS;
            }
            else if (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return ExcelVersion.XLSX;
            }
            else
            {
                throw new UnSupportedTypeException();
            }
        }
    }
}