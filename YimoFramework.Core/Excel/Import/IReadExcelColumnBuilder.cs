using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YimoFramework.ExcelImport
{
    /// <summary>
    /// 格式化单元格读取数据
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface IReadExcelColumnBuilder<T> where T : class
    {
        ///// <summary>
        ///// 格式化，
        ///// </summary>
        ///// <param name="format"></param>
        ///// <returns></returns>
        //IReadExcelColumnBuilder<T> Formatted(String format);

        ///// <summary>
        ///// 自定义单元格
        ///// </summary>
        ///// <param name="block"></param>
        ///// <returns></returns>
        //IReadExcelColumnBuilder<T> Do(Func<T, String> block);
    }
}
