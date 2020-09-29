namespace CommonLibrary.ExcelHelper.Model.Exception
{
    /// <summary>
    /// 数据为空
    /// </summary>
    public class NoDataException : System.Exception
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public NoDataException() : base("数据为空")
        {
        }
    }
}