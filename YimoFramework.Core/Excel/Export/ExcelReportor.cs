using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.SS.UserModel;

namespace YimoFramework.ExcelExport
{
    /// <summary>
    /// 生成Excel
    /// </summary>
    public class ExcelExporter : IRootExcelBuilder
    {
        #region Private Properties
        /// <summary>
        /// Excel保存数据流
        /// </summary>
        private Stream Writer { get; set; }
        #endregion

        #region Public Properties
        /// <summary>
        /// 获取或文档属性
        /// </summary>
        public Action<DocumentSummaryInformation> DocumentProperty { get; set; }
        /// <summary>
        /// 获取或文档属性
        /// </summary>
        public Action<SummaryInformation> SummaryProperty { get; set; }
        /// <summary>
        /// 获取或设置默认的列头样式
        /// </summary>
        public Action<ICellStyle> DefaultHeaderStyle { get; set; }
        /// <summary>
        /// 获取或设置默认的数据列样式
        /// </summary>
        public Action<ICellStyle> DefaultBodyStyle { get; set; }

        /// <summary>
        /// 获取或设置默认的标题行样式
        /// </summary>
        public ICellStyle HeaderDefaultStyle { get; set; }

        /// <summary>
        /// 获取或设置默认的数据行样式
        /// </summary>
        public ICellStyle BodyDefaultStyle { get; set; }

        /// <summary>
        /// 获取或设置工作薄
        /// </summary>
        public HSSFWorkbook Workbook { get; set; }
        #endregion

        #region 构造器
        /// <summary>
        /// 
        /// </summary>
        /// <param name="writer">Excel数据流</param>
        public ExcelExporter(Stream writer)
        {
            this.Workbook = new HSSFWorkbook();
            this.Writer = writer;

            this.Initialize();
        }

        /// <summary>
        /// 初始化Excel工作薄
        /// </summary>
        private void Initialize()
        {
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();

            // -----------------------------------------------
            //  SummaryInformation
            // -----------------------------------------------
            si.Author = "yimo";
            si.LastAuthor = "yimo.link";
            si.CreateDateTime = DateTime.Now;
            si.LastSaveDateTime = DateTime.Now;
            si.ApplicationName = "yimo";
            si.Keywords = "yimo";
            si.Subject = "";
            si.Title = "";
            dsi.Company = "";

            if (this.DocumentProperty != null)
            {
                this.DocumentProperty(dsi);
            }
            if (this.SummaryProperty != null)
            {
                this.SummaryProperty(si);
            }
            this.Workbook.SummaryInformation = si;
            this.Workbook.DocumentSummaryInformation = dsi;

            var headerStyle = this.Workbook.CreateCellStyle();
            this.SetDefaultHeaderStyle(headerStyle);
            if (this.DefaultHeaderStyle != null)
            {
                this.DefaultHeaderStyle(headerStyle);
            }
            this.HeaderDefaultStyle = headerStyle;

            var bodyStyle = this.Workbook.CreateCellStyle();
            this.SetDefaultBodyStyle(bodyStyle);
            if (this.DefaultBodyStyle != null)
            {
                this.DefaultBodyStyle(bodyStyle);
            }
            this.BodyDefaultStyle = bodyStyle;
        }

        #endregion

        #region 公开的方法

        /// <summary>
        /// 生成工作表
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="items">数据源</param>
        /// <param name="columns">数据列构建器</param>
        public void GenerateSheet<T>(String sheetName, IEnumerable<T> items, ExcelColumnBuilder<T> columns)
            where T : class
        {
            //设置Table样式
            var sheet = String.IsNullOrEmpty(sheetName) ? this.Workbook.CreateSheet() : this.Workbook.CreateSheet(sheetName);
            this.CurrentSheet = sheet;
            //生成样式
            this.GenerateHeader(columns, sheet);
            //生成Body
            this.GenerateItems(items, columns, sheet);
        }

       

        /// <summary>
        /// 保存Excel
        /// </summary>
        public void Save()
        {
            ///保存Excel
            this.Workbook.Write(this.Writer);
        }

        #endregion

        #region IRootExcelSheetBuilder 成员
        /// <summary>
        /// 获取或设置当前工作薄
        /// </summary>
        public ISheet CurrentSheet
        {
            get;
            private set;
        }

        /// <summary>
        /// 生成工作表
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="dataSource">数据源</param>
        /// <param name="columns">数据列构建器</param>
        /// <returns></returns>
        public IRootExcelBuilder Sheet<T>(string sheetName, IEnumerable<T> dataSource, Action<IRootExcelColumnBuilder<T>> columns)
            where T : class
        {
            var columnBuilder = CreateColumnBuilder(columns);

            this.GenerateSheet(sheetName, dataSource, columnBuilder);

            return this;
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// 生成数据行
        /// </summary>
        protected virtual void GenerateItems<T>(IEnumerable<T> dataSource, ExcelColumnBuilder<T> columns, ISheet sheet)
            where T : class
        {
            if (!dataSource.Any())
            {
                return;
            }

            //保存单元格样式信息
            List<ICellStyle> bodyStyles = new List<ICellStyle>(columns.ColumnCount);
            //生成样式信息
            foreach (var column in columns)
            {
                #region 设置样式
                var bodyStyle = this.Workbook.CreateCellStyle();
                this.SetDefaultBodyStyle(bodyStyle);
                if (column.BodyStyle != null)
                {
                    column.BodyStyle(bodyStyle);
                }
                bodyStyle.WrapText = false;//设置为自动换行
                bodyStyle.Alignment = HorizontalAlignment.CenterSelection;//设置为水平居中
                bodyStyle.VerticalAlignment = VerticalAlignment.Center;//设置为垂直居中
                bodyStyles.Add(bodyStyle);
                #endregion
            }

            //第一行用户数据列头
            Int32 rowIndex = 1;
            foreach (var item in dataSource)
            {
                var row = sheet.CreateRow(rowIndex);
                Int32 columnIndex = 0;
                rowIndex++;

                foreach (var column in columns)
                {
                    var cell = row.CreateCell(columnIndex);
                    cell.CellStyle = bodyStyles[columnIndex];
                    /************如果是集合就将集合的数据显示到一个Cell并换行***********/
                    if (column.IsCollection)
                    {
                        if (column.CustomRenderer == null)
                        {
                            #region 获取数据 并 格式化单元格数据
                            Object value = null;
                            if (column.ColumnDelegate != null)
                            {
                                value = column.ColumnDelegate(item);
                            }
                            else
                            {
                                var property = item.GetType().GetProperty(column.Name);
                                if (property != null)
                                {
                                    value = property.GetValue(item, null);
                                }
                            }

                            String formattedValue = null;
                            if (value != null)
                            {
                                if (column.Format != null)
                                {
                                    formattedValue = String.Format(column.Format, value);
                                }
                                else
                                {
                                    if (value is IEnumerable)
                                    {
                                        var IValue = value as IEnumerable;
                                        var enumerable = IValue as object[] ?? IValue.Cast<object>().ToArray();
                                        var list = Enumerable.OfType<object>(enumerable);
                                        if (list.Count() > 1)
                                        {
                                            foreach (var obj in list)
                                            {
                                                formattedValue += "\r\n" + obj.ToString();
                                            }
                                        }
                                        else
                                        {
                                            var firstOrDefault = list.FirstOrDefault();
                                            if (firstOrDefault != null)
                                                formattedValue = firstOrDefault.ToString();
                                        }
                                    }
                                    //formattedValue = value.ToString();
                                }
                            }
                            #endregion

                            cell.SetCellValue(formattedValue);
                        }
                        else
                        {
                            //自定义呈现单元格 
                            column.CustomRenderer(item, cell);
                        }
                    }
                    /***********************/
                    else
                    {
                        if (column.CustomRenderer == null)
                        {
                            #region 获取数据 并 格式化单元格数据
                            dynamic value = null;
                            if (column.ColumnDelegate != null)
                            {
                                value = column.ColumnDelegate(item);
                            }
                            else
                            {
                                var property = item.GetType().GetProperty(column.Name);
                                if (property != null)
                                {
                                    value = property.GetValue(item, null);
                                }
                            }

                            if (value != null)
                            {
                                if (column.Format != null)
                                {
                                    var formattedValue = String.Format(column.Format, value);
                                    cell.SetCellValue(formattedValue);
                                }
                                else
                                {
                                    if (value is decimal)
                                    {
                                        cell.SetCellValue((double)value);

                                    }
                                    else if (value is int)
                                    {
                                        cell.SetCellValue((double)value);
                                    }
                                    else
                                    {
                                        cell.SetCellValue(value.ToString());
                                    }
                                }
                            }
                            #endregion

                        }
                        else
                        {
                            //自定义呈现单元格 
                            column.CustomRenderer(item, cell);
                        }
                    }

                    #region 设置链接
                    if (null != column.HrefDelegate)
                    {
                        var href = column.HrefDelegate(item);
                        if (null != href)
                        {
                            var link = new HSSFHyperlink(new HyperlinkType())
                            {
                                Address = href.ToString()
                            };
                            cell.Hyperlink = link;
                        }
                    }
                    #endregion

                    columnIndex++;
                }
            }
        }

        /// <summary>
        /// 生成数据列头
        /// </summary>
        /// <returns></returns>
        protected virtual void GenerateHeader<T>(ExcelColumnBuilder<T> columns, ISheet sheet)
            where T : class
        {
            //设置默认的行高、列宽
            sheet.DefaultColumnWidth = 15;
            sheet.DefaultRowHeight = 14;

            var headRow = sheet.CreateRow(0);
            Int32 columnIndex = 0;
            foreach (var column in columns)
            {
                //设置为自动调整数据列宽度
                sheet.AutoSizeColumn(columnIndex, true);

                var headerStyle = this.Workbook.CreateCellStyle();
                var columnCell = headRow.CreateCell(columnIndex);
                this.SetDefaultHeaderStyle(headerStyle);
                if (column.HeaderStyle != null)
                {
                    //自定义设置数据列标题样式
                    column.HeaderStyle(headerStyle);
                }
                if (column.CustomHeader != null)
                {
                    //自定义生成数据列标题单元格                 
                    column.CustomHeader(columnCell);
                }
                else
                {
                    columnCell.SetCellValue(column.Name);
                }
                //设置样式
                columnCell.CellStyle = headerStyle;

                columnIndex++;
            }
            //创建最后一空列，防止读取时出现问题。
            headRow.CreateCell(columnIndex);
        }

        /// <summary>
        /// 设置默认的列标题样式
        /// </summary>
        private void SetDefaultHeaderStyle(ICellStyle style)
        {
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            var font = this.Workbook.CreateFont();
            font.FontName = "宋体";
            font.Boldweight = 800;
            font.FontHeightInPoints = 12;
            style.SetFont(font);
        }

        /// <summary>
        /// 设置默认的数据行的样式
        /// </summary>
        private void SetDefaultBodyStyle(ICellStyle style)
        {
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            var font = this.Workbook.CreateFont();
            font.FontName = "宋体";
            font.FontHeightInPoints = 11;
            style.SetFont(font);
        }

        #endregion

        #region Static Method ReprotExcel
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="writer"></param>
        /// <param name="sheets"></param>
        public static void ReportExcel(Stream writer, Action<IRootExcelBuilder> sheets)
        {
            ReportExcel(writer, null, null, null, null, sheets);
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="writer"></param>
        /// <param name="documentProperty"></param>
        /// <param name="summaryProperty"></param>
        /// <param name="defaultHeaderStyle"></param>
        /// <param name="defaultBodyStyle"></param>
        /// <param name="sheets"></param>
        public static void ReportExcel(Stream writer,
            Action<DocumentSummaryInformation> documentProperty, Action<SummaryInformation> summaryProperty,
            Action<ICellStyle> defaultHeaderStyle, Action<ICellStyle> defaultBodyStyle,
            Action<IRootExcelBuilder> sheets)
        {
            var exporter = new ExcelExporter(writer);
            exporter.DocumentProperty = documentProperty;
            exporter.SummaryProperty = summaryProperty;
            exporter.DefaultHeaderStyle = defaultHeaderStyle;
            exporter.DefaultBodyStyle = defaultBodyStyle;

            //生成多个Sheet
            if (null != sheets)
            {
                sheets(exporter);
            }

            exporter.Save();
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataSource"></param>
        /// <param name="sheetName"></param>
        /// <param name="writer"></param>
        /// <param name="columns"></param>
        public static void ReportExcel<T>(IEnumerable<T> dataSource, String sheetName, Stream writer, Action<IRootExcelColumnBuilder<T>> columns)
            where T : class
        {
            ReportExcel(dataSource, sheetName, writer, columns, null, null, null, null);
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataSource"></param>
        /// <param name="sheetName"></param>
        /// <param name="writer"></param>
        /// <param name="columns"></param>
        /// <param name="documentProperty"></param>
        /// <param name="workbookProperty"></param>
        public static void ReportExcel<T>(IEnumerable<T> dataSource, String sheetName, Stream writer, Action<IRootExcelColumnBuilder<T>> columns,
            Action<DocumentSummaryInformation> documentProperty, Action<SummaryInformation> summaryProperty,
            Action<ICellStyle> defaultHeaderStyle, Action<ICellStyle> defaultBodyStyle)
            where T : class
        {
            var exporter = new ExcelExporter(writer);
            exporter.DocumentProperty = documentProperty;
            exporter.SummaryProperty = summaryProperty;
            exporter.DefaultHeaderStyle = defaultHeaderStyle;
            exporter.DefaultBodyStyle = defaultBodyStyle;

            var columnBuilder = CreateColumnBuilder(columns);
            //生成Excel
            exporter.GenerateSheet(sheetName, dataSource, columnBuilder);

            exporter.Save();
        }

        /// <summary>
        /// 创建数据列构造器
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="columns"></param>
        /// <returns></returns>
        private static ExcelColumnBuilder<T> CreateColumnBuilder<T>(Action<IRootExcelColumnBuilder<T>> columns)
            where T : class
        {
            var builder = new ExcelColumnBuilder<T>();

            if (columns != null)
            {
                columns(builder);
            }

            return builder;
        }



        #endregion
    }
}
