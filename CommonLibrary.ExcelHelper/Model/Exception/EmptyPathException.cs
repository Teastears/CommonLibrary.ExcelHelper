namespace CommonLibrary.ExcelHelper.Model.Exception
{
    /// <summary>
    /// 导出目标文件路径为空
    /// </summary>
    public class EmptyPathException : System.Exception
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public EmptyPathException() : base("导出目标文件路径为空")
        {
        }
    }
}