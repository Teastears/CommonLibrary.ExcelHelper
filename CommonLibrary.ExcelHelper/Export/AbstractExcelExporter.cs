using CommonLibrary.ExcelHelper.Base;
using CommonLibrary.ExcelHelper.Enum;
using CommonLibrary.ExcelHelper.ExportStyle;
using CommonLibrary.ExcelHelper.Model;
using CommonLibrary.ExcelHelper.Model.Exception;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace CommonLibrary.ExcelHelper.Export
{
    /// <summary>
    /// 导出操作抽象类
    /// </summary>
    public abstract class AbstractExcelExporter : IExcelExporter
    {
        /// <summary>
        /// 工作簿
        /// </summary>
        protected IWorkbook Workbook;

        /// <summary>
        /// Excel版本
        /// </summary>
        public ExcelVersion ExcelVersion { set; get; } = ExcelVersion.XLSX;

        /// <summary>
        /// 创建工作薄
        /// </summary> 
        /// <returns></returns>
        protected void CreateWorkbook()
        {
            if (ExcelVersion == ExcelVersion.XLS)
            {
                Workbook = new HSSFWorkbook();
            }
            else
            {
                Workbook = new XSSFWorkbook();
            }
        }

        /// <summary>
        /// 数据验证
        /// </summary>
        /// <returns></returns>
        protected abstract bool DataCheck();

        /// <summary>
        /// 设置单元格的值
        /// </summary>
        /// <param name="cell">要设置的单元格</param>
        /// <param name="obj">值对象</param>
        protected void SetCellValue(ICell cell, object obj)
        {
            if (obj == null)
                return;

            switch (obj)
            {
                case string val:
                    cell.SetCellValue(val);
                    break;

                case byte val:
                    cell.SetCellValue(val);
                    break;

                case short val:
                    cell.SetCellValue(val);
                    break;

                case int val:
                    cell.SetCellValue(val);
                    break;

                case long val:
                    cell.SetCellValue(val);
                    break;

                case sbyte val:
                    cell.SetCellValue(val);
                    break;

                case ushort val:
                    cell.SetCellValue(val);
                    break;

                case uint val:
                    cell.SetCellValue(val);
                    break;

                case ulong val:
                    cell.SetCellValue(val);
                    break;

                case float val:
                    cell.SetCellValue(val);
                    break;

                case double val:
                    cell.SetCellValue(val);
                    break;

                case decimal val:
                    cell.SetCellValue((double)val);
                    break;

                case bool val:
                    cell.SetCellValue(val);
                    break;

                case char val:
                    cell.SetCellValue(val.ToString());
                    break;

                case DateTime val:
                    cell.SetCellValue(val);
                    break;

                default:
                    cell.SetCellValue(obj.ToString());
                    break;
            }
        }

        /// <summary>
        /// 创建新的单元格样式对象
        /// </summary>
        /// <returns></returns>
        public ICellStyle CreateNewStyle()
        {
            return Workbook.CreateCellStyle();
        }

        /// <summary>
        /// 导出到文件
        /// </summary>
        /// <param name="ExportFilePath">导出文件路径</param>
        /// <param name="ExportStyle">导出时应用的样式</param>
        public virtual void ExportToFile(string ExportFilePath = "ExportFile.xlsx", IExportStyle ExportStyle = null)
        {
            if (string.IsNullOrEmpty(ExportFilePath)) throw new EmptyPathException();
            var data = ExportToStream(ExportStyle);
            FileStream fs = new FileStream(ExportFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            data.Flush();
            data.Seek(0, SeekOrigin.Begin);
            data.WriteTo(fs);

            fs.Dispose();
            data.Close();
            data.Dispose();
        }

        /// <summary>
        /// 导出到内存流
        /// </summary>
        /// <param name="ExportStyle">导出时应用的样式</param>
        /// <returns></returns>
        public abstract NPOIMemoryStream ExportToStream(IExportStyle ExportStyle = null);
    }
}