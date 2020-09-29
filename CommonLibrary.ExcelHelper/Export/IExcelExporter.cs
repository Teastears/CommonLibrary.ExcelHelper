using CommonLibrary.ExcelHelper.ExportStyle;
using CommonLibrary.ExcelHelper.Model;

namespace CommonLibrary.ExcelHelper.Export
{
    /// <summary>
    /// 导出Excel接口
    /// </summary>
    public interface IExcelExporter
    {
        /// <summary>
        /// 导出到文件
        /// </summary>
        /// <param name="ExportFilePath">导出文件路径</param>
        /// <param name="ExportStyle">导出时应用的样式</param>
        void ExportToFile(string ExportFilePath,IExportStyle ExportStyle = null);

        /// <summary>
        /// 导出到内存流
        /// </summary>
        /// <param name="ExportStyle">导出时应用的样式</param>
        /// <returns></returns>
        NPOIMemoryStream ExportToStream(IExportStyle ExportStyle = null);
    }
}