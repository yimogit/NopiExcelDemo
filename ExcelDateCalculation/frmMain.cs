using ExcelDateCalculation.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using YimoFramework.ExcelImport;

namespace ExcelDateCalculation
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "表格文件 (*.xls,*.xlsx)|*.xls;*.xlsx";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                DataTable dt = ExcelHelper.Import(openFileDialog.FileName);
                if (dt == null)
                {
                    MessageBox.Show("读取失败");
                    return;
                }
                var result = ExcelReader<TestModel>.ReadDataTable(dt, c =>
                {
                    #region dev
                    //c.For((k, v) => { k.No = v; }, "案件號");
                    //c.For((k, v) =>
                    //{
                    //    k.BeginTime = DateTime.Parse(v);
                    //}, "派工時間");
                    //c.For((k, v) => { k.EndTime = DateTime.Parse(v); ; }, "結案時間");
                    //c.For((k, v) =>
                    //{
                    //    k.Result = v;
                    //}, "處理時效");
                    #endregion
                    for (var i = 1; i < dt.Rows.Count; i++)
                    {
                        c.For((k, v) => { k.Result += v + Environment.NewLine; }, i);
                    }
                });
                Thread t = new Thread(() =>
                {
                    txtResult.BeginInvoke(new Action(() =>
                    {
                        foreach (var item in result)
                        {
                            //txtResult.AppendText(item.No + "----------" + new TestModelExt().GetJishuanResult(item) + Environment.NewLine);
                            txtResult.AppendText(item.Result + Environment.NewLine);
                        }
                    }));

                });

                t.Start();
            }
        }
        private void frmMain_Load(object sender, EventArgs e)
        {

        }

    }
}
