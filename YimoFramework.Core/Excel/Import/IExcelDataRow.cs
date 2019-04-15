using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace YimoFramework.ExcelImport
{
    /// <summary>
    /// DataRow的包装类
    /// </summary>
    public interface IExcelDataRow
    {
        String this[Int32 columnIndex] { get; }
        String this[String columnName] { get; }

        /// <summary>
        /// 原始数据行
        /// </summary>
        DataRow OriginalDataRow { get; }

        /// <summary>
        /// 获取一个值，该值指示位于指定索引处的列是否包含空值。
        /// </summary>
        /// <param name="columnIndex">列的从零开始的索引。</param>
        /// <returns> 如果列包含空值，则为 true；否则为 false</returns>
        Boolean IsNull(Int32 columnIndex);

        /// <summary>
        /// 获取一个值，该值指示位于指定索引处的列是否包含空值。
        /// </summary>
        /// <param name="columnName">列的名称</param>
        /// <returns> 如果列包含空值，则为 true；否则为 false</returns>
        Boolean IsNull(String columnName);

        /// <summary>
        /// 信息项大小
        /// </summary>
        Int32 ItemCount { get; }

        /// <summary>
        /// 获取第几行信息
        /// </summary>
        Int32 RowIndex { get; }
    }
}
