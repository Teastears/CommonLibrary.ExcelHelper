using CommonLibrary.ExcelHelper.ExportStyle;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace CommonLibrary.ExcelHelper.Demo
{
    internal class MyStyle : IExportStyle
    {
        protected bool IsInit = false;
        protected ICellStyle DefaultBodyStyle;

        protected ICellStyle BodyStyle_Date;

        protected ICellStyle HeaderStyle;
        public Func<ICellStyle> CreateNewStyle { set; protected get; }
        public Func<string, short> CreateNewDataFormat { set; protected get; }

        protected void Init()
        {
            if (IsInit)
                return;

            DefaultBodyStyle = CreateNewStyle();
            DefaultBodyStyle.BorderBottom = BorderStyle.Thin;
            DefaultBodyStyle.BorderLeft = BorderStyle.Thin;
            DefaultBodyStyle.BorderRight = BorderStyle.Thin;
            DefaultBodyStyle.BorderTop = BorderStyle.Thin;

            BodyStyle_Date = CreateNewStyle();
            BodyStyle_Date.BorderBottom = BorderStyle.Thin;
            BodyStyle_Date.BorderLeft = BorderStyle.Thin;
            BodyStyle_Date.BorderRight = BorderStyle.Thin;
            BodyStyle_Date.BorderTop = BorderStyle.Thin;
            BodyStyle_Date.DataFormat = CreateNewDataFormat("yyyy年m月d日"); 
        }

        public ICellStyle GetCellStyle(int SheetIndex, int ColumnIndex, int RowIndex)
        {
            Init();
            switch (ColumnIndex)
            {
                case 3: return BodyStyle_Date;
                default:
                    return DefaultBodyStyle;
            } 
        }

        public int GetColumnWidth(int SheetIndex, int ColumnIndex)
        {
            switch (ColumnIndex)
            {
                case 3: return 15;
                default:
                    return 8;
            }
        }

        public ICellStyle GetHeaderStyle(int SheetIndex, int ColumnIndex)
        {
            if (HeaderStyle == null)
            {
                HeaderStyle = CreateNewStyle();
                HeaderStyle.FillPattern = FillPattern.SolidForeground;
                HeaderStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
                HeaderStyle.BorderBottom = BorderStyle.Thin;
                HeaderStyle.BorderLeft = BorderStyle.Thin;
                HeaderStyle.BorderRight = BorderStyle.Thin;
                HeaderStyle.BorderTop = BorderStyle.Thin;
            }
            return HeaderStyle;
        }
    }
}