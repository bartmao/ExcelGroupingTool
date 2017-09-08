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

        XSSFWorkbook workbook;

        public List<CItem> Items { get; set; } = new List<CItem>();

        private void button1_Click(object sender, EventArgs e)
        {
            gv.AutoGenerateColumns = false;
            gv.Columns.Add("Sheet", "Sheet");
            gv.Columns.Add("Designation", "Designation");
            gv.Columns.Add("Qty", "Qty");
            gv.Columns.Add("CS", "CS");
            gv.Columns.Add("RMB", "RMB");
            gv.Columns.Add("Power", "Power");

            using (var stream = File.OpenRead(@"C:\Data\1.xlsx"))
            {
                workbook = new XSSFWorkbook(stream);
                workbook.SetForceFormulaRecalculation(true);
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
            }

            MessageBox.Show(Items.Count.ToString());

        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (var s = File.Create("11.xlsx"))
            {
                workbook.Write(s);
            }
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Parent == null)
            {
                var cate = e.Node.Tag as IGrouping<string, CItem>;

                gv.Rows.Clear();
                foreach (var citem in cate.ToList())
                {
                    var row = new DataGridViewRow();
                    row.Tag = citem;
                    row.CreateCells(gv);
                    row.Cells[0].Value = citem.SheetName;
                    row.Cells[1].Value = citem.Designation.StringCellValue;
                    row.Cells[2].Value = GetCellValue(citem.Qty);
                    row.Cells[3].Value = GetCellValue(citem.CS);
                    row.Cells[4].Value = GetCellValue(citem.RMB);
                    row.Cells[5].Value = GetCellValue(citem.Power);
                    gv.Rows.Add(row);
                }
            }
        }

        private void gv_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            var citem = gv.Rows[e.RowIndex].Tag as CItem;
            var cell = typeof(CItem).GetProperty(gv.Columns[e.ColumnIndex].Name).GetValue(citem) as ICell;
            if (cell.CellType == CellType.Formula)
            {
                cell.SetCellFormula(e.Value.ToString());
                var eval = workbook.GetCreationHelper().CreateFormulaEvaluator();
                eval.EvaluateFormulaCell(cell);
            }
            else if (cell.CellType == CellType.Numeric)
            {
                cell.SetCellValue(double.Parse(e.Value.ToString()));
            }
            else
            {
                cell.SetCellValue(e.Value.ToString());
            }
        }

        private string GetCellValue(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Formula:
                    return cell.CellFormula;
                case CellType.Blank:
                    return null;
                default:
                    return null;
            }
        }
    }
}
