using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel数据分类整理工具
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public List<CItem> Items { get; set; } = new List<CItem>();

        private void button1_Click(object sender, EventArgs e)
        {
            using (var stream = File.OpenRead(@"C:\Data\1.xlsx"))
            {
                var workbook = new XSSFWorkbook(stream);

                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    if (Regex.IsMatch(sheet.SheetName, @"[\d|\.]+"))
                    {
                        for (int j = 5; j < sheet.LastRowNum; j++)
                        {
                            var row = sheet.GetRow(j);
                            var val = row.GetCell(8).StringCellValue;
                            if (!string.IsNullOrWhiteSpace(val))
                            {
                                var designation = row.GetCell(3);
                                var qty = row.GetCell(9);
                                var cs = row.GetCell(10);
                                var rmb = row.GetCell(13);
                                var power = row.GetCell(16);
                                Items.Add(new CItem()
                                {
                                    SheetName = sheet.SheetName,
                                    Designation = designation,
                                    Qty = qty,
                                    CS = cs,
                                    RMB = rmb,
                                    Power = power
                                });
                            }
                        }
                    }
                }

                var categories = Items.GroupBy(i => i.Designation.StringCellValue.Trim().ToLower())
                    .OrderBy(g => g.Key);
                foreach (var cate in categories)
                {
                    var node = treeView1.Nodes.Add(cate.Key);
                    node.Tag = cate;
                    node.Nodes.AddRange(cate.Select(c => new TreeNode()
                    {
                        Tag = c,
                        Text = string.Format("{0}({1}:R{2})", c.Designation.StringCellValue, c.SheetName, c.Designation.RowIndex)
                    }).ToArray());
                }
                //using (var s = File.Create("11.xlsx"))
                //{
                //    Items[0].RMB.SetCellFormula("10000 * 0.5");
                //    workbook.Write(s);
                //}
            }

            MessageBox.Show(Items.Count.ToString());

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Parent == null)
            {
                var cate = e.Node.Tag as IGrouping<string, CItem>;
                gv.DataSource = cate.ToList();
            }
        }

        private void dataGridView1_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {

        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            //var cell = gv.Rows[e.RowIndex].Cells[1];
            //var cellValue = cell.Value as ICell;
            //cell.Value = cellValue.StringCellValue;
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            var item = gv.Rows[e.RowIndex].DataBoundItem;
            var colName = gv.Columns[e.ColumnIndex].Name;
            if (e.ColumnIndex > 0)
            {
                var cellValue = typeof(CItem).GetProperty(colName).GetValue(item) as ICell;
                try
                {
                    switch (cellValue.CellType)
                    {
                        case CellType.Numeric:
                            e.Value = cellValue.NumericCellValue;
                            e.FormattingApplied = true;
                            break;
                        case CellType.String:
                            e.Value = cellValue.StringCellValue;
                            e.FormattingApplied = true;
                            break;
                        case CellType.Formula:
                            e.Value = cellValue.CellFormula;
                            e.FormattingApplied = true;
                            break;
                        default:
                            break;
                    }

                }
                catch (Exception ex)
                {

                }
            }

        }
    }
}
