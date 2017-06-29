using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using NPOI.HSSF.UserModel;//2007office
using NPOI.XSSF.UserModel;//xlsx
using NPOI.SS.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Util;

namespace NPOI
{
    /// <summary>
    /// 表示单元格的维度，通常用于表达合并单元格的维度
    /// </summary>
    public struct Dimension
    {
        /// <summary>
        /// 含有数据的单元格(通常表示合并单元格的第一个跨度行第一个跨度列)，该字段可能为null
        /// </summary>
        public ICell DataCell;

        /// <summary>
        /// 行跨度(跨越了多少行)
        /// </summary>
        public int RowSpan;

        /// <summary>
        /// 列跨度(跨越了多少列)
        /// </summary>
        public int ColumnSpan;

        /// <summary>
        /// 合并单元格的起始行索引
        /// </summary>
        public int FirstRowIndex;

        /// <summary>
        /// 合并单元格的结束行索引
        /// </summary>
        public int LastRowIndex;

        /// <summary>
        /// 合并单元格的起始列索引
        /// </summary>
        public int FirstColumnIndex;

        /// <summary>
        /// 合并单元格的结束列索引
        /// </summary>
        public int LastColumnIndex;
    }

    public static class ExcelExtension
    {
        /// <summary>
        /// 判断指定行列所在的单元格是否为合并单元格，并且输出该单元格的维度
        /// </summary>
        /// <param name="sheet">Excel工作表</param>
        /// <param name="rowIndex">行索引，从0开始</param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <param name="dimension">单元格维度</param>
        /// <returns>返回是否为合并单元格的布尔(Boolean)值</returns>
        public static bool IsMergeCell(this ISheet sheet, int rowIndex, int columnIndex, out Dimension dimension)
        {
            dimension = new Dimension
            {
                DataCell = null,
                RowSpan = 1,
                ColumnSpan = 1,
                FirstRowIndex = rowIndex,
                LastRowIndex = rowIndex,
                FirstColumnIndex = columnIndex,
                LastColumnIndex = columnIndex
            };

            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                sheet.IsMergedRegion(range);

                //这种算法只有当指定行列索引刚好是合并单元格的第一个跨度行第一个跨度列时才能取得合并单元格的跨度
                //if (range.FirstRow == rowIndex && range.FirstColumn == columnIndex)
                //{
                //    dimension.DataCell = sheet.GetRow(range.FirstRow).GetCell(range.FirstColumn);
                //    dimension.RowSpan = range.LastRow - range.FirstRow + 1;
                //    dimension.ColumnSpan = range.LastColumn - range.FirstColumn + 1;
                //    dimension.FirstRowIndex = range.FirstRow;
                //    dimension.LastRowIndex = range.LastRow;
                //    dimension.FirstColumnIndex = range.FirstColumn;
                //    dimension.LastColumnIndex = range.LastColumn;
                //    break;
                //}

                if ((rowIndex >= range.FirstRow && range.LastRow >= rowIndex) && (columnIndex >= range.FirstColumn && range.LastColumn >= columnIndex))
                {
                    dimension.DataCell = sheet.GetRow(range.FirstRow).GetCell(range.FirstColumn);
                    dimension.RowSpan = range.LastRow - range.FirstRow + 1;
                    dimension.ColumnSpan = range.LastColumn - range.FirstColumn + 1;
                    dimension.FirstRowIndex = range.FirstRow;
                    dimension.LastRowIndex = range.LastRow;
                    dimension.FirstColumnIndex = range.FirstColumn;
                    dimension.LastColumnIndex = range.LastColumn;
                    break;
                }
            }

            bool result;
            if (rowIndex >= 0 && sheet.LastRowNum > rowIndex)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (columnIndex >= 0 && row.LastCellNum > columnIndex)
                {
                    ICell cell = row.GetCell(columnIndex);
                    result = cell.IsMergedCell;

                    if (dimension.DataCell == null)
                    {
                        dimension.DataCell = cell;
                    }
                }
                else
                {
                    result = false;
                }
            }
            else
            {
                result = false;
            }

            return result;
        }

        /// <summary>
        /// 判断指定行列所在的单元格是否为合并单元格，并且输出该单元格的行列跨度
        /// </summary>
        /// <param name="sheet">Excel工作表</param>
        /// <param name="rowIndex">行索引，从0开始</param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <param name="rowSpan">行跨度，返回值最小为1，同时表示没有行合并</param>
        /// <param name="columnSpan">列跨度，返回值最小为1，同时表示没有列合并</param>
        /// <returns>返回是否为合并单元格的布尔(Boolean)值</returns>
        public static bool IsMergeCell(this ISheet sheet, int rowIndex, int columnIndex, out int rowSpan, out int columnSpan)
        {
            Dimension dimension;
            bool result = sheet.IsMergeCell(rowIndex, columnIndex, out dimension);

            rowSpan = dimension.RowSpan;
            columnSpan = dimension.ColumnSpan;

            return result;
        }

        /// <summary>
        /// 判断指定单元格是否为合并单元格，并且输出该单元格的维度
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="dimension">单元格维度</param>
        /// <returns>返回是否为合并单元格的布尔(Boolean)值</returns>
        public static bool IsMergeCell(this ICell cell, out Dimension dimension)
        {
            return cell.Sheet.IsMergeCell(cell.RowIndex, cell.ColumnIndex, out dimension);
        }

        /// <summary>
        /// 判断指定单元格是否为合并单元格，并且输出该单元格的行列跨度
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="rowSpan">行跨度，返回值最小为1，同时表示没有行合并</param>
        /// <param name="columnSpan">列跨度，返回值最小为1，同时表示没有列合并</param>
        /// <returns>返回是否为合并单元格的布尔(Boolean)值</returns>
        public static bool IsMergeCell(this ICell cell, out int rowSpan, out int columnSpan)
        {
            return cell.Sheet.IsMergeCell(cell.RowIndex, cell.ColumnIndex, out rowSpan, out columnSpan);
        }

        /// <summary>
        /// 返回上一个跨度行，如果rowIndex为第一行，则返回null
        /// </summary>
        /// <param name="sheet">Excel工作表</param>
        /// <param name="rowIndex">行索引，从0开始</param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <returns>返回上一个跨度行</returns>
        public static IRow PrevSpanRow(this ISheet sheet, int rowIndex, int columnIndex)
        {
            return sheet.FuncSheet(rowIndex, columnIndex, (currentDimension, isMerge) =>
            {
                //上一个单元格维度
                Dimension prevDimension;
                sheet.IsMergeCell(currentDimension.FirstRowIndex - 1, columnIndex, out prevDimension);
                return prevDimension.DataCell.Row;
            });
        }

        /// <summary>
        /// 返回下一个跨度行，如果rowIndex为最后一行，则返回null
        /// </summary>
        /// <param name="sheet">Excel工作表</param>
        /// <param name="rowIndex">行索引，从0开始</param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <returns>返回下一个跨度行</returns>
        public static IRow NextSpanRow(this ISheet sheet, int rowIndex, int columnIndex)
        {
            return sheet.FuncSheet(rowIndex, columnIndex, (currentDimension, isMerge) =>
                isMerge ? sheet.GetRow(currentDimension.FirstRowIndex + currentDimension.RowSpan) : sheet.GetRow(rowIndex));
        }

        /// <summary>
        /// 返回上一个跨度行，如果row为第一行，则返回null
        /// </summary>
        /// <param name="row">行</param>
        /// <returns>返回上一个跨度行</returns>
        public static IRow PrevSpanRow(this IRow row)
        {
            return row.Sheet.PrevSpanRow(row.RowNum, row.FirstCellNum);
        }

        /// <summary>
        /// 返回下一个跨度行，如果row为最后一行，则返回null
        /// </summary>
        /// <param name="row">行</param>
        /// <returns>返回下一个跨度行</returns>
        public static IRow NextSpanRow(this IRow row)
        {
            return row.Sheet.NextSpanRow(row.RowNum, row.FirstCellNum);
        }

        /// <summary>
        /// 返回上一个跨度列，如果columnIndex为第一列，则返回null
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <returns>返回上一个跨度列</returns>
        public static ICell PrevSpanCell(this IRow row, int columnIndex)
        {
            return row.Sheet.FuncSheet(row.RowNum, columnIndex, (currentDimension, isMerge) =>
            {
                //上一个单元格维度
                Dimension prevDimension;
                row.Sheet.IsMergeCell(row.RowNum, currentDimension.FirstColumnIndex - 1, out prevDimension);
                return prevDimension.DataCell;
            });
        }

        /// <summary>
        /// 返回下一个跨度列，如果columnIndex为最后一列，则返回null
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <returns>返回下一个跨度列</returns>
        public static ICell NextSpanCell(this IRow row, int columnIndex)
        {
            return row.Sheet.FuncSheet(row.RowNum, columnIndex, (currentDimension, isMerge) =>
                row.GetCell(currentDimension.FirstColumnIndex + currentDimension.ColumnSpan));
        }

        /// <summary>
        /// 返回上一个跨度列，如果cell为第一列，则返回null
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>返回上一个跨度列</returns>
        public static ICell PrevSpanCell(this ICell cell)
        {
            return cell.Row.PrevSpanCell(cell.ColumnIndex);
        }

        /// <summary>
        /// 返回下一个跨度列，如果columnIndex为最后一列，则返回null
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>返回下一个跨度列</returns>
        public static ICell NextSpanCell(this ICell cell)
        {
            return cell.Row.NextSpanCell(cell.ColumnIndex);
        }

        /// <summary>
        /// 返回指定行索引所在的合并单元格(区域)中的第一行(通常是含有数据的行)
        /// </summary>
        /// <param name="sheet">Excel工作表</param>
        /// <param name="rowIndex">行索引，从0开始</param>
        /// <returns>返回指定列索引所在的合并单元格(区域)中的第一行</returns>
        public static IRow GetDataRow(this ISheet sheet, int rowIndex)
        {
            return sheet.FuncSheet(rowIndex, 0, (currentDimension, isMerge) => sheet.GetRow(currentDimension.FirstRowIndex));
        }

        /// <summary>
        /// 返回指定列索引所在的合并单元格(区域)中的第一行第一列(通常是含有数据的单元格)
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="columnIndex">列索引</param>
        /// <returns>返回指定列索引所在的合并单元格(区域)中的第一行第一列</returns>
        public static ICell GetDataCell(this IRow row, int columnIndex)
        {
            return row.Sheet.FuncSheet(row.RowNum, columnIndex, (currentDimension, isMerge) => currentDimension.DataCell);
        }

        /// <summary>
        /// 判断指定单元格是否为合并单元格，并且输出该单元格的高度和宽度
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="height">单元格高度</param>
        /// <param name="width">单元格宽度</param>
        /// <returns>返回是否为合并单元格的布尔(Boolean)值</returns>
        public static bool IsMergeCellShape(this ICell cell, out float height, out float width)
        {
            //宽度计算说明 96dpi下 8个字符的宽度实际上是7.29
            //Using the Calibri font as an example, the maximum digit width of 11 point font size is 7 pixels (at 96 dpi). 
            //If you set a column width to be eight characters wide, e.g. setColumnWidth(columnIndex, 8*256), 
            //then the actual value of visible characters (the value shown in Excel) is derived from the following equation: 
            //Truncate([numChars*7+5]/7*256)/256 = 8; which gives 7.29.

            Dimension dimension = new Dimension();
            height = 0;
            width = 0;
            bool isMerge = cell.Sheet.IsMergeCell(cell.RowIndex, cell.ColumnIndex, out dimension);
            //
            if (!isMerge)
            {
                //方法一
                //getColumnWidth
                //int getColumnWidth(int columnIndex)
                //get the width(in units of 1 / 256th of a character width)
                //Character width is defined as the maximum digit width of the numbers 0, 1, 2, ... 9 as rendered using the default font (first font in the workbook)
                //Parameters:
                //                columnIndex - -the column to get(0 - based)
                //Returns:
                //                width - the width in units of 1 / 256th of a character width
                //PS:尝试32.04左右的倍率，转换成像素 pix=width/32.04
                width = cell.Sheet.GetColumnWidth(dimension.FirstColumnIndex);
                width = width / 32;//test 20170724
                //

                //方法二 说明 不同字体有可能拉伸
                //getColumnWidthInPixels
                //public float getColumnWidthInPixels(int column)
                //Description copied from interface: Sheet
                //get the width in pixel
                //Please note, that this method works correctly only for workbooks with the default font size 
                //(Arial 10pt for .xls and Calibri 11pt for .xlsx). If the default font is changed the column width can be streched
                //width= cell.Sheet.GetColumnWidthInPixels(dimension.FirstColumnIndex);

                //short getHeight()
                //Get the row's height measured in twips (1/20th of a point). If the height is not set, the default worksheet value is returned, See Sheet.getDefaultRowHeightInPoints()
                //Returns:
                //row height measured in twips(1 / 20th of a point)
                height = cell.Sheet.GetRow(dimension.FirstRowIndex).Height;
            }
            else//单元格高度宽度特殊处理
            {
                for (int i = dimension.FirstColumnIndex; i <= dimension.LastColumnIndex; i++)
                {
                    //方法一 字符
                    //width = width + cell.Sheet.GetColumnWidth(i);
                    //方法二 像素
                    width = width + cell.Sheet.GetColumnWidthInPixels(i);
                }
                for (int i= dimension.FirstRowIndex; i <= dimension.LastRowIndex; i++)
                {
                    height = height + cell.Sheet.GetRow(i).Height;
                }
            }
            return isMerge;
        }


        private static T FuncSheet<T>(this ISheet sheet, int rowIndex, int columnIndex, Func<Dimension, bool, T> func)
        {
            //当前单元格维度
            Dimension currentDimension;
            //是否为合并单元格
            bool isMerge = sheet.IsMergeCell(rowIndex, columnIndex, out currentDimension);

            return func(currentDimension, isMerge);
        }
    }
}
