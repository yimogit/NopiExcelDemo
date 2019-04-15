using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Linq.Expressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace YimoFramework.ExcelExport
{
    /// <summary>
    /// 工作表中的单元格数据列构建器
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelColumnBuilder<T> : INestedExcelColumnBuilder<T>, IRootExcelColumnBuilder<T>, IEnumerable<ExcelColumn<T>>
        where T : class
    {
        /// <summary>
        /// 工作表的数据列
        /// </summary>
        private readonly List<ExcelColumn<T>> columns = new List<ExcelColumn<T>>();

        /// <summary>
        /// 工作表的当前数据列
        /// </summary>
        private ExcelColumn<T> currentColumn;

        /// <summary>
        ///
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ExcelColumn<T> this[int index]
        {
            get
            {
                return columns[index];
            }
        }

        /// <summary>
        /// 获取列数
        /// </summary>
        public Int32 ColumnCount
        {
            get { return columns.Count; }
        }

        #region 构造器
        public ExcelColumnBuilder()
        {
        }
        #endregion

        /// <summary>
        ///  从表达式中获取属性的名称
        /// </summary>
        /// <param name="expression">表达式</param>
        /// <returns>得到属性名称</returns>
        public static String ExpressionToName(Expression<Func<T, Object>> expression)
        {
            var memberExpression = RemoveUnary(expression.Body) as MemberExpression;

            return memberExpression == null ? String.Empty : memberExpression.Member.Name;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="body"></param>
        /// <returns></returns>
        private static Expression RemoveUnary(Expression body)
        {
            var unary = body as UnaryExpression;
            if (unary != null)
            {
                return unary.Operand;
            }
            return body;
        }

        #region IRootExcelColumnBuilder<T> 成员

        /// <summary>
        /// 生成列
        /// </summary>
        /// <param name="name">属性名称，列标题</param>
        /// <returns></returns>
        public INestedExcelColumnBuilder<T> For(String name)
        {
            currentColumn = new ExcelColumn<T> { Name = name };
            columns.Add(currentColumn);
            return this;
        }

        /// <summary>
        /// 从表达式中生成工作表的 数据列
        /// </summary>
        /// <param name="expression">表达式，表达式的属性名称作为列标题，表达式的值作为单元格数据</param>
        /// <returns></returns>
        public INestedExcelColumnBuilder<T> For(Expression<Func<T, Object>> expression)
        {
            currentColumn = new ExcelColumn<T>
            {
                Name = ExpressionToName(expression),
                ColumnDelegate = expression.Compile(),
            };

            columns.Add(currentColumn);
            return this;
        }

        /// <summary>
        /// 从表达式中生成工作表的 数据列
        /// </summary>
        /// <param name="func">表达式的值作为单元格的数据</param>
        /// <param name="name">列标题</param>
        /// <returns></returns>
        public INestedExcelColumnBuilder<T> For(Func<T, dynamic> func, String name)
        {
            currentColumn = new ExcelColumn<T>
            {
                Name = name,
                ColumnDelegate = func
            };

            columns.Add(currentColumn);
            return this;
        }
        #endregion

        #region INestedExcelColumnBuilder<T> 成员

        /// <summary>
        /// 格式化单元格数据
        /// </summary>
        /// <param name="format">格式字符串</param>
        /// <returns></returns>
        public INestedExcelColumnBuilder<T> Formatted(String format)
        {
            this.currentColumn.Format = format;
            return this;
        }

        /// <summary>
        /// 为单元格设置超链接
        /// </summary>
        /// <param name="href">超链接地址</param>
        /// <returns></returns>
        public INestedExcelColumnBuilder<T> Href(string href)
        {
            this.currentColumn.HrefDelegate = (e => { return href; });
            return this;
        }

        /// <summary>
        /// 设置单元格的链接地址
        /// </summary>
        /// <param name="expression">链接地址表达式</param>
        /// <returns></returns>
        public INestedExcelColumnBuilder<T> Href(Func<T, Object> expression)
        {
            this.currentColumn.HrefDelegate = expression;
            return this;
        }

        /// <summary>
        /// 自定义生成单元格
        /// </summary>
        /// <param name="block"></param>
        /// <returns></returns>
        public INestedExcelColumnBuilder<T> Do(Action<T, ICell> block)
        {
            currentColumn.CustomRenderer = block;
            return this;
        }

        public INestedExcelColumnBuilder<T> IsCollection(bool block)
        {
            currentColumn.IsCollection = block;
            return this;
        }

        /// <summary>
        /// 自定义生成工作表的数据列头
        /// </summary>
        /// <param name="block"></param>
        /// <returns></returns>
        public INestedExcelColumnBuilder<T> Header(Action<ICell> block)
        {
            currentColumn.CustomHeader = block;
            return this;
        }

        /// <summary>
        /// 自定义设置数据列头样式
        /// </summary>
        /// <param name="block"></param>
        /// <returns></returns>
        public INestedExcelColumnBuilder<T> HeaderStyle(Action<ICellStyle> block)
        {
            currentColumn.HeaderStyle = block;
            return this;
        }

        /// <summary>
        /// 自定义设置单元格样式
        /// </summary>
        /// <param name="block"></param>
        /// <returns></returns>
        public INestedExcelColumnBuilder<T> BodyStyle(Action<ICellStyle> block)
        {
            currentColumn.BodyStyle = block;
            return this;
        }

        #endregion

        #region IEnumerable<ExcelColumn<T>> 成员

        public IEnumerator<ExcelColumn<T>> GetEnumerator()
        {
            return this.columns.GetEnumerator();
        }

        #endregion

        #region IEnumerable 成员

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.columns.GetEnumerator();
        }

        #endregion
    }
}
