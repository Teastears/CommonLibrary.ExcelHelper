using System.Collections.Generic;
using System.Data;

namespace CommonLibrary.ExcelHelper.Import
{
    /// <summary>
    /// 导入Excel接口
    /// </summary>
    public interface IExcelImporter
    {
        /// <summary>
        /// 导入数据
        /// </summary>
        /// <returns></returns>
        DataSet Import();

        /// <summary>
        /// 导入数据
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <returns></returns>
        List<T> Import<T>() where T : class;
    }
}