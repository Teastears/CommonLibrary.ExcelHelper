namespace CommonLibrary.ExcelHelper.Model.Exception
{
    /// <summary>
    /// 不支持的文件类型
    /// </summary>
    public class UnSupportedTypeException : System.Exception
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public UnSupportedTypeException() : base("不支持的文件类型")
        {
        }
    }
}