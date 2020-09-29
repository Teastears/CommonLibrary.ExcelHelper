using CommonLibrary.ExcelHelper.Base;
using CommonLibrary.ExcelHelper.Model.Exception;
using System.IO;

namespace CommonLibrary.ExcelHelper.Import
{
    /// <summary>
    /// 文件导入处理类
    /// </summary>
    public class ExcelFileImporter : ExcelStreamImporter
    {
        /// <summary>
        /// 创建工作簿(依据文件流)
        /// </summary>
        /// <returns></returns>
        protected override void CreateWorkbook()
        {
            ReadFile();
            ExcelVersion = ExcelBase.GetVersion(DataSourceFilePath);
            base.CreateWorkbook();
        }

        /// <summary>
        /// 读取文件到文件流
        /// </summary>
        protected void ReadFile()
        {
            if (string.IsNullOrWhiteSpace(DataSourceFilePath))
                throw new EmptyPathException();
            if (!File.Exists(DataSourceFilePath))
                throw new FileNotFoundException();
            SourceData = File.OpenRead(DataSourceFilePath);
        }

        /// <summary>
        /// 要导入的数据源文件路径
        /// </summary>
        public string DataSourceFilePath { get; set; }

        /// <summary>
        /// 释放
        /// </summary>
        public override void Dispose()
        {
            SourceData.Dispose();
            base.Dispose();
        }
    }
}