using CommonLibrary.ExcelHelper.ExportStyle;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace CommonLibrary.ExcelHelper.Demo
{
    internal class MyStyle : AbstractStyle
    {
        protected bool IsInit = false;
        protected ICellStyle DefaultBodyStyle;

        protected ICellStyle BodyStyle_Date;

        protected ICellStyle HeaderStyle; 

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

        public override ICellStyle GetCellStyle(int SheetIndex, int ColumnIndex, int RowIndex)
        {
            Init();
            switch (ColumnIndex)
            {
                case 3: return BodyStyle_Date;
                default:
                    return DefaultBodyStyle;
            } 
        }

        public override int GetColumnWidth(int SheetIndex, int ColumnIndex)
        {
            switch (ColumnIndex)
            {
                case 3: return 15;
                default:
                    return 8;
            }
        }

        public override ICellStyle GetHeaderStyle(int SheetIndex, int ColumnIndex)
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