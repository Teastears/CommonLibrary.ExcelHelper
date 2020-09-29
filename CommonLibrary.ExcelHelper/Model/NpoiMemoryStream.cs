using System.IO;

namespace CommonLibrary.ExcelHelper.Model
{
    /// <summary>
    /// NPOI专用内存流
    /// </summary>
    public class NPOIMemoryStream : MemoryStream
    {
        /// <summary>
        /// 是否允许关闭
        /// </summary>
        public bool AllowClose { get; set; } = true;

        /// <summary>
        /// 重写关闭操作
        /// </summary>
        public override void Close()
        {
            if (AllowClose)
                base.Close();
        }
    }
}

//NPOI在写入流后，会关闭流，导致无法后续操作，所以需要重写一个新的内存流类，用于写入