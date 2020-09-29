using System.Collections.Generic;

namespace CommonLibrary.ExcelHelper.Model
{
    /// <summary>
    /// 导入配置类
    /// </summary>
    public class ImportSheetSetting
    {
        /// <summary>
        /// 导入配置类构造函数
        /// </summary>
        /// <param name="SheetIndex">工作表索引序号，从0开始</param>
        /// <param name="HeaderRowIndex">导入数据表表头行号，从0开始</param>
        public ImportSheetSetting(int SheetIndex, int HeaderRowIndex = 0)
        {
            this.SheetIndex = SheetIndex;
            this.HeaderRowIndex = HeaderRowIndex;
        }

        /// <summary>
        /// 导入配置类构造函数
        /// </summary>
        /// <param name="SheetName">工作表名称</param>
        /// <param name="HeaderRowIndex">导入数据表表头行号，从0开始</param>
        public ImportSheetSetting(string SheetName, int HeaderRowIndex = 0)
        {
            this.SheetName = SheetName;
            this.HeaderRowIndex = HeaderRowIndex;
        }

        /// <summary>
        /// 导入数据表表头行号
        /// </summary>
        public int HeaderRowIndex { get; set; }

        /// <summary>
        /// 工作表索引序号
        /// </summary>
        public int? SheetIndex { get; set; }

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { get; set; }         
    }
}