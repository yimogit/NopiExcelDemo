using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace YimoFramework.ExcelImport
{
    /// <summary>
    /// Excel读取器
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelReader<T> : IDisposable
        where T : class
    {
        #region 验证Excel
        /// <summary>
        /// 验证单元格是否有效
        /// </summary>
        /// <param name="columns"></param>
        /// <param name="table">信息源</param>
        /// <param name="message">验证失败的信息</param>
        /// <returns></returns>
        private void ValidateColumns(ReadExcelColumnBuilder<T> columns, DataTable table, String sheetName)
        {
            String message = String.Empty;
            if (null == columns)
            {
                throw new ArgumentException("columns");
            }
            foreach (var column in columns)
            {
                if (column.CustomDelegate != null)
                {//自定义绑定信息，忽略验证
                    continue;
                }
                if (!String.IsNullOrEmpty(column.ColumnName))
                {//信息列名称 优先级大于 索引
                    if (!table.Columns.Contains(column.ColumnName))
                    {//校验信息列是否有效
                        message += String.Format("工作表中不存在数据列 [{0}] 。{1}", column.ColumnName, "<br>");
                    }
                    continue;
                }
                if (column.ColumnIndex >= 0)
                {//校验信息列是否有效
                    if (table.Columns.Count <= column.ColumnIndex)
                    {
                        message += String.Format("工作表中不存在第 [{0}] 列错误。{1}", column.ColumnIndex, "<br>");
                    }
                    continue;
                }
            }
            if (!String.IsNullOrEmpty(message))
            {//验证失败，抛出异常
                throw new ValidateColumnException(message);
            }
        }

        /// <summary>
        /// 创建对象实例
        /// </summary>
        /// <param name="createInstance"></param>
        /// <returns></returns>
        private T CreateInstance(Func<T> createInstance)
        {
            T item = null;
            if (createInstance != null)
            {
                item = createInstance();
            }
            if (item == null)
            {
                item = System.Activator.CreateInstance<T>();
            }
            if (item == null)
            {
                throw new InvalidOperationException("创建对象实例失败。");
            }
            return item;
        }

        public List<T> ReadDataTable(DataTable sheetTable, ReadExcelColumnBuilder<T> columns, Func<T> createInstance)
        {
            if (sheetTable == null)
                throw new SheetTableNullException("工作表中不存在数据");

            //校验信息列
            this.ValidateColumns(columns, sheetTable, "");
            List<T> items = new List<T>(sheetTable.Rows.Count);
            //循环读取每行Excel信息
            for (Int32 rowIndex = 0; rowIndex < sheetTable.Rows.Count; rowIndex++)
            {
                DataRow row = sheetTable.Rows[rowIndex];
                T item = this.CreateInstance(createInstance);

                //该行数据为空
                if (row.ItemArray.Where(e => null == e).Count() == row.ItemArray.Length)
                {
                    continue;
                }

                using (DataRowWrapper dataWrapper = new DataRowWrapper(row, rowIndex))
                {
                    foreach (var column in columns)
                    {
                        ProcessCellDataException error = this.ProcessCellData(item, dataWrapper, column);
                        if (null != error)
                        {//保存错误信息
                            throw error;
                        }
                    }
                }

                items.Add(item);
            }
            return items;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        private ProcessCellDataException WrapperException(String message, Action action)
        {
            try
            {
                action();
            }
            catch (ProcessCellDataException processExp)
            {
                return processExp;
            }
            catch (Exception exp)
            {
                return new ProcessCellDataException(message, exp);
            }
            return null;
        }

        /// <summary>
        /// 处理一列信息
        /// </summary>
        private ProcessCellDataException ProcessCellData(T item, DataRowWrapper dataWrapper, ReadExcelColumn<T> column)
        {
            ProcessCellDataException error = null;
            if (column.CustomDelegate != null)
            {//自定义获取信息，优先级最高
                error = WrapperException(String.Format("读取第 [{0}] 行信息错误。", dataWrapper.RowIndex + 1), () => column.CustomDelegate(item, dataWrapper));
            }
            else
            {
                String data = null;
                if (!String.IsNullOrEmpty(column.ColumnName))
                {//信息列名称获取信息优先级大于信息索引
                    error = WrapperException(String.Format("读取第 [{0}] 行，[{1}] 列单元格信息错误。", dataWrapper.RowIndex + 1, column.ColumnName),
                        () => data = dataWrapper[column.ColumnName]);
                    if (null == error)
                    {//转换信息
                        error = WrapperException(String.Format("读取第 [{0}] 行，[{1}] 列单元格信息错误。", dataWrapper.RowIndex + 1, column.ColumnName),
                            () => column.CustomEvaluater(item, data));
                    }
                }
                else
                {
                    error = WrapperException(String.Format("读取第 [{0}] 行，第 [{1}] 列单元格信息错误。", dataWrapper.RowIndex + 1, column.ColumnIndex + 1),
                        () => data = dataWrapper[column.ColumnIndex]);
                    if (null == error)
                    {//转换信息
                        error = WrapperException(String.Format("读取第 [{0}] 行，第 [{1}] 列单元格信息错误。", dataWrapper.RowIndex + 1, column.ColumnIndex + 1),
                            () => column.CustomEvaluater(item, data));
                    }
                }
            }

            return error;
        }

        #endregion

        #region Static Method ReadExcel
        /// <summary>
        /// 将DataTable转成集合
        /// </summary>
        /// <param name="sheetTable"></param>
        /// <param name="columns"></param>
        /// <returns></returns>
        public static List<T> ReadDataTable(DataTable sheetTable, Action<IRootReadExcelColumnBuilder<T>> columns)
        {
            //读取Excel
            return ReadDataTable(sheetTable, columns, null);
        }

        /// <summary>
        /// 读取DataTable到集合中
        /// </summary>
        /// <param name="sheetTable"></param>
        /// <param name="columns"></param>
        /// <returns></returns>
        public static List<T> ReadDataTable(DataTable sheetTable, Action<IRootReadExcelColumnBuilder<T>> columns, Func<T> createInstance)
        {
            ExcelReader<T> reader = new ExcelReader<T>();

            ReadExcelColumnBuilder<T> columnBuilder = CreateColumnBuilder(columns);
            //读取Excel
            List<T> result = reader.ReadDataTable(sheetTable, columnBuilder, createInstance);

            return result;
        }

        /// <summary>
        /// 创建信息列构造器
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="columns"></param>
        /// <returns></returns>
        private static ReadExcelColumnBuilder<T> CreateColumnBuilder(Action<IRootReadExcelColumnBuilder<T>> columns)
        {
            var builder = new ReadExcelColumnBuilder<T>();

            if (columns != null)
            {
                columns(builder);
            }

            return builder;
        }
        #endregion

        #region IDisposable 成员

        public void Dispose() { }

        #endregion
    }
}