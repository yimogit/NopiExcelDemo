using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;

namespace YimoFramework.ExcelExport
{
    /// <summary>
    /// 生成工作表中的数据列
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface IRootExcelColumnBuilder<T>
        where T : class
    {
        /// <summary>
        /// 生成工作表的数据列
        /// </summary>
        /// <param name="name">获取数据源中的属性作为单元格数据，列标题</param>
        /// <returns></returns>
        INestedExcelColumnBuilder<T> For(String name);

        /// <summary>
        /// 从表达式中生成工作表的 数据列
        /// </summary>
        /// <param name="expression">表达式，表达式的属性名称作为列标题，表达式的值作为单元格数据</param>
        /// <returns></returns>
        INestedExcelColumnBuilder<T> For(Expression<Func<T, Object>> expression);

        /// <summary>
        /// 从表达式中生成工作表的 数据列
        /// </summary>
        /// <param name="func">表达式的值作为单元格的数据</param>
        /// <param name="name">列标题</param>
        /// <returns></returns>
        INestedExcelColumnBuilder<T> For(Func<T, dynamic> func, String name);
    }
}
