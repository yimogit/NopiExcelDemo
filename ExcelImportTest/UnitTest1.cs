using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using YimoFramework.ExcelImport;

namespace ExcelImportTest
{
    /*
     手动添加了以下引用
    //System.Data
    //System.Xml
    //System.Web
     表格格式若不规范则特殊处理
     */
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void 导入测试()
        {
            var xlsxPath = @"F:\test\ExcelDateCalculation\ExcelImportTest\TestExcel\测试xlsx导入.xlsx";
            var dt = ExcelHelper.Import(xlsxPath);
        }
        [TestMethod]
        public void xls导入并读取测试()
        {
            var xlsPath = @"F:\test\ExcelDateCalculation\ExcelImportTest\TestExcel\测试xls导入.xls";
            var dt = ExcelHelper.Import(xlsPath);
            var result = ExcelReader<TestModel>.ReadDataTable(dt, c =>
            {
                c.For((k, v) => { k.Title = v; }, "标题");
                c.For((k, v) =>
                {
                    k.Content = v;
                }, "内容");
                c.For((k, v) =>
                {
                    k.CreateTime = DateTime.Parse(v);
                }, "创建时间");
            });
            Assert.IsTrue(result.Count > 0);
        }
        [TestMethod]
        public void xlsx导入并读取测试()
        {
            var xlsxPath = @"F:\test\ExcelDateCalculation\ExcelImportTest\TestExcel\测试xlsx导入.xlsx";
            var dt2 = ExcelHelper.Import(xlsxPath);
            var result2 = ExcelReader<TestModel>.ReadDataTable(dt2, c =>
            {
                //使用索引转换
                c.For((k, v) => { k.Title = v; }, 1);
                c.For((k, v) =>
                {
                    k.Content = v;
                }, 2);
                c.For((k, v) =>
                {
                    DateTime dtime;
                    DateTime.TryParse(v, out dtime);
                    k.CreateTime = dtime;
                }, "创建时间");
            });
            Assert.IsTrue(result2.Count > 0);
        }
        class TestModel
        {
            public string Title { get; set; }
            public string Content { get; set; }
            public DateTime CreateTime { get; set; }
        }
    }
}
