using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace YimoFramework.ExcelImport
{
    public class ExcelHelper
    {
        /// <summary>
        /// 根据文件路径导入Excel
        /// </summary>
        /// <param name="filePath">文件完整路径</param>
        /// <param name="sheetName">表名，默认取第一张</param>
        /// <returns>可能为null的DataTable</returns>
        /// <remarks>无需返回错误信息</remarks>
        public static DataTable Import(string filePath, string sheetName = "")
        {
            string msg = string.Empty;
            return Import(filePath, ref msg, sheetName);
        }
        /// <summary>
        /// 根据文件路径导入Excel
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="errorMsg">错误信息</param>
        /// <param name="sheetName">表名，默认取第一张</param>
        /// <returns>可能为null的DataTable</returns>
        public static DataTable Import(string filePath, ref string errorMsg, string sheetName = "")
        {
            var excelType = GetExcelFileType(filePath);
            if (GetExcelFileType(filePath) == null)
            {
                errorMsg = "请选择正确的Excel文件";
                return null;
            }
            if (!File.Exists(filePath))
            {
                errorMsg = "没有找到要导入的Excel文件";
                return null;
            }
            DataTable dt;
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                dt = ImportExcel(stream, excelType.Value, sheetName);
            }
            if (dt == null)
                errorMsg = "导入失败,请选择正确的Excel文件";
            return dt;
        }

        /// <summary>
        /// 上传Excel导入
        /// </summary>
        /// <param name="file">上载文件对象</param>
        /// <param name="sheetName">表名，默认取第一张</param>
        /// <returns>可能为null的DataTable</returns>
        /// <remarks>无需返回错误信息</remarks>
        public static DataTable Import(System.Web.HttpPostedFileBase file, string sheetName = "")
        {
            string msg = string.Empty;
            return Import(file, ref msg, sheetName);
        }
        /// <summary>
        /// 上传Excel导入
        /// </summary>
        /// <param name="file">上载文件对象</param>
        /// <param name="errorMsg">错误信息</param>
        /// <param name="sheetName">表名，默认取第一张</param>
        /// <returns></returns>
        public static DataTable Import(System.Web.HttpPostedFileBase file, ref string errorMsg, string sheetName = "")
        {
            if (file == null || file.InputStream == null || file.InputStream.Length == 0)
            {
                errorMsg = "请选择要导入的Excel文件";
                return null;
            }
            var excelType = GetExcelFileType(file.FileName);
            if (excelType == null)
            {
                errorMsg = "请选择正确的Excel文件";
                return null;
            }
            using (var stream = new MemoryStream())
            {
                file.InputStream.Position = 0;
                file.InputStream.CopyTo(stream);
                var dt = ImportExcel(stream, excelType.Value, sheetName);
                if (dt == null)
                    errorMsg = "导入失败,请选择正确的Excel文件";
                return dt;
            }
        }
        /// <summary>
        /// 根据Excel格式读取Excel
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="type">Excel类型，xls/xlsx</param>
        /// <param name="sheetName">表名，默认取第一张</param>
        /// <returns>DataTable</returns>
        private static DataTable ImportExcel(Stream stream, ExcelExtType type, string sheetName)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            try
            {
                //xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现
                switch (type)
                {
                    case ExcelExtType.xlsx:
                        workbook = new XSSFWorkbook(stream);
                        break;
                    default:
                        workbook = new HSSFWorkbook(stream);
                        break;
                }
                ISheet sheet = null;
                //获取工作表 默认取第一张
                if (string.IsNullOrWhiteSpace(sheetName))
                    sheet = workbook.GetSheetAt(0);
                else
                    sheet = workbook.GetSheet(sheetName);

                if (sheet == null)
                    return null;
                IEnumerator rows = sheet.GetRowEnumerator();
                #region 获取表头
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                for (int j = 0; j < cellCount; j++)
                {
                    ICell cell = headerRow.GetCell(j);
                    if (cell != null)
                    {
                        dt.Columns.Add(cell.ToString());
                    }
                    else
                    {
                        dt.Columns.Add("");
                    }
                }
                #endregion
                #region 获取内容
                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    DataRow dataRow = dt.NewRow();

                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            //判断单元格是否为日期格式
                            if (row.GetCell(j).CellType == NPOI.SS.UserModel.CellType.Numeric && HSSFDateUtil.IsCellDateFormatted(row.GetCell(j)))
                            {
                                if (row.GetCell(j).DateCellValue.Year > 1000)
                                {
                                    dataRow[j] = row.GetCell(j).DateCellValue.ToString();
                                }
                                else
                                {
                                    dataRow[j] = row.GetCell(j).ToString();

                                }
                            }
                            else
                            {
                                dataRow[j] = row.GetCell(j).ToString();
                            }
                        }
                    }
                    dt.Rows.Add(dataRow);
                }
                #endregion

            }
            catch (Exception ex)
            {
                dt=null;
            }
            finally
            {
                //if (stream != null)
                //{
                //    stream.Close();
                //    stream.Dispose();
                //}
            }

            return dt;
        }

        private enum ExcelExtType
        {
            xls,
            xlsx,
        }
        private static Nullable<ExcelExtType> GetExcelFileType(string fileName)
        {
            var ext = Path.GetExtension(fileName);
            if (!string.IsNullOrWhiteSpace(ext) && (ext.ToLower() == ".xls" || ext.ToLower() == ".xlsx"))
                return ext.ToLower() == ".xls" ? ExcelExtType.xls : ExcelExtType.xlsx;
            else
                return null;
        }

    }
}
