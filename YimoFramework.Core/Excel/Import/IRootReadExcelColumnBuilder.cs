using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;

namespace YimoFramework.ExcelImport
{
    public interface IRootReadExcelColumnBuilder<T> where T : class
    {
        /// <summary>
        /// 生成工作表的数据列读取器
        /// </summary>
        /// <param name="expression">读取单元格数据的表达式</param>
        /// <param name="index">按数据列索引读取</param>
        /// <returns></returns>
        IReadExcelColumnBuilder<T> For(Action<T, String> expression, Int32 index);

        /// <summary>
        /// 生成工作表的数据列读取器
        /// </summary>
        /// <param name="expression">读取单元格数据的表达式</param>
        /// <param name="name">列标题名称</param>
        /// <returns></returns>
        IReadExcelColumnBuilder<T> For(Action<T, String> expression, String name);

        /// <summary>
        /// 生成工作表的数据列读取器
        /// </summary>
        /// <param name="expression">读取工作表一行数据的表达式</param>
        /// <returns></returns>
        IReadExcelColumnBuilder<T> For(Action<T, IExcelDataRow> expression);
    }
}
