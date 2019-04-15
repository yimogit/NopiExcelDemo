using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace YimoFramework.ExcelExport
{
    /// <summary>
    /// 格式化工作表中的单元格、样式、数据
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface INestedExcelColumnBuilder<T> where T : class
    {
        /// <summary>
        /// 格式化单元格数据
        /// </summary>
        /// <param name="format">格式字符串</param>
        /// <returns></returns>
        INestedExcelColumnBuilder<T> Formatted(String format);

        /// <summary>
        /// 设置单元格的超链接
        /// </summary>
        /// <param name="href">超链接地址</param>
        /// <returns></returns>
        INestedExcelColumnBuilder<T> Href(String href);

        /// <summary>
        /// 设置单元格的链接地址
        /// </summary>
        /// <param name="expression">链接地址表达式</param>
        /// <returns></returns>
        INestedExcelColumnBuilder<T> Href(Func<T, Object> expression);

        /// <summary>
        /// 自定义工作表的单元格列标题
        /// </summary>
        /// <param name="block">列标题</param>
        /// <returns></returns>
        INestedExcelColumnBuilder<T> Header(Action<ICell> block);

        /// <summary>
        /// 自定义设置数据列头样式
        /// </summary>
        /// <param name="block"></param>
        /// <returns></returns>
        INestedExcelColumnBuilder<T> HeaderStyle(Action<ICellStyle> block);

        /// <summary>
        /// 自定义设置单元格样式
        /// </summary>
        /// <param name="block"></param>
        /// <returns></returns>
        INestedExcelColumnBuilder<T> BodyStyle(Action<ICellStyle> block);

        /// <summary>
        /// 自定义生成单元格
        /// </summary>
        /// <param name="block"></param>
        /// <returns></returns>
        INestedExcelColumnBuilder<T> Do(Action<T, ICell> block);

        /// <summary>
        /// 此单元格是否是集合
        /// </summary>
        /// <param name="isCollection"></param>
        /// <returns></returns>
        INestedExcelColumnBuilder<T> IsCollection(bool isCollection);
    }
}
