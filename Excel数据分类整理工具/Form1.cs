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
using System.Reflection;
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
        string FilePath { get; set; }

        XSSFWorkbook workbook;

        List<ICell> ChangedCells = new List<ICell>();

        public List<CItem> Items { get; set; } = new List<CItem>();

        private void button1_Click(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Filter = "表格文件|*.xlsx";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                FilePath = dlg.FileName;
                ReadExcel();
            }
        }

        private void ReadExcel()
        {
            var tbHeaderNum = int.Parse(txtTbHeader.Text.Trim());
            var groupName = txtGroupName.Text.Trim();

            //gv.AutoGenerateColumns = false;
            //gv.Columns.Add("Sheet", "Sheet");
            //gv.Columns.Add("Designation", "Designation");
            //gv.Columns.Add("Qty", "Qty");
            //gv.Columns.Add("CS", "CS");
            //gv.Columns.Add("RMB", "RMB");
            //gv.Columns.Add("Power", "Power");

            var templateFile = AppDomain.CurrentDomain.BaseDirectory + "//Template1.txt";
            var items = new ExcelTableParser(FilePath).Parse(templateFile);

            return;
            using (var stream = File.OpenRead(FilePath))
            {
                workbook = new XSSFWorkbook(stream);
                //workbook.SetForceFormulaRecalculation(true);

                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    if (Regex.IsMatch(sheet.SheetName, @"[\d|\.]+"))
                    {
                        var sheetName = sheet.SheetName;
                        var parser = new ExcelTableParser(FilePath);
                        parser.Parse(sheetName, 5);
                    }
                }

                return;

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



            //MessageBox.Show(Items.Count.ToString());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            using (var s = File.Create("11.xlsx"))
            {
                workbook.Write(s);
            }
        }

        private void gv_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            var gvRow = gv.Rows[e.RowIndex];
            var gvCell = gv.Rows[e.RowIndex].Cells[e.ColumnIndex];
            var value = e.Value.ToString();
            var citem = gv.Rows[e.RowIndex].Tag as CItem;
            var cell = typeof(CItem).GetProperty(gv.Columns[e.ColumnIndex].Name).GetValue(citem) as ICell;
            var oldVal = GetCellValue(cell);

            // if using a formula
            if (value.StartsWith("="))
            {
                cell.SetCellType(CellType.Formula);
                cell.SetCellFormula(value.Substring(1, value.Length - 1));
                var eval = workbook.GetCreationHelper().CreateFormulaEvaluator();
                eval.EvaluateFormulaCell(cell);
            }
            else
            {
                // only consider number/string here
                double v;
                if (double.TryParse(value, out v))
                {
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(v);
                }
                else
                {
                    cell.SetCellValue(value.ToString());
                }
            }

            if (string.IsNullOrWhiteSpace(oldVal))
                oldVal = "空值";
            var node = treeView1.SelectedNode;
            while (node.Parent != null)
            {
                node = node.Parent;
            }
            listBox1.Items.Add(new MyListBoxItem(string.Format("{0}!{1}从【{2}】改变到【{3}】\r\n", citem.SheetName, GetCellPosition(cell), oldVal, GetCellValue(cell)), node));
            gvCell.Style.BackColor = Color.SkyBlue;
            //RefreshRow(gvRow);
            ChangedCells.Add(cell);
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Parent == null)
            {
                var items = (e.Node.Tag as IGrouping<string, CItem>).ToList();
                Bind(items);
            }
        }

        private void Bind(IEnumerable<CItem> items)
        {
            gv.Rows.Clear();
            gv.Tag = items;
            foreach (var citem in items)
            {
                var row = BindRow(citem);
                gv.Rows.Add(row);
            }

            var sumRow = new DataGridViewRow();
            sumRow.DefaultCellStyle.Font = new Font("宋体", 10, FontStyle.Bold);
            sumRow.ReadOnly = true;
            sumRow.Tag = null;
            sumRow.CreateCells(gv);
            sumRow.Cells[0].Value = "汇总";
            sumRow.Cells[1].Value = "";
            sumRow.Cells[2].Value = CalculateSum(items.Select(i => i.Qty));
            sumRow.Cells[3].Value = CalculateSum(items.Select(i => i.CS));
            sumRow.Cells[4].Value = CalculateSum(items.Select(i => i.RMB));
            sumRow.Cells[5].Value = CalculateSum(items.Select(i => i.Power));
            gv.Rows.Add(sumRow);
        }

        private DataGridViewRow BindRow(CItem citem, DataGridViewRow row = null)
        {
            if (row == null)
            {
                row = new DataGridViewRow();
                row.Tag = citem;
                row.CreateCells(gv);
            }

            row.Cells[0].Value = citem.SheetName;
            row.Cells[1].Value = citem.Designation.StringCellValue;
            row.Cells[2].Value = GetCellValue(citem.Qty);
            row.Cells[3].Value = GetCellValue(citem.CS);
            row.Cells[4].Value = GetCellValue(citem.RMB);
            row.Cells[5].Value = GetCellValue(citem.Power);

            if (ChangedCells.Contains(citem.Qty)) row.Cells[2].Style.BackColor = Color.SkyBlue;
            if (ChangedCells.Contains(citem.CS)) row.Cells[3].Style.BackColor = Color.SkyBlue;
            if (ChangedCells.Contains(citem.RMB)) row.Cells[4].Style.BackColor = Color.SkyBlue;
            if (ChangedCells.Contains(citem.Power)) row.Cells[5].Style.BackColor = Color.SkyBlue;

            return row;
        }

        private void RefreshSum()
        {
            var items = gv.Tag as IEnumerable<CItem>;

            var sumRow = gv.Rows[gv.Rows.Count - 1];
            sumRow.Cells[2].Value = CalculateSum(items.Select(i => i.Qty));
            sumRow.Cells[3].Value = CalculateSum(items.Select(i => i.CS));
            sumRow.Cells[4].Value = CalculateSum(items.Select(i => i.RMB));
            sumRow.Cells[5].Value = CalculateSum(items.Select(i => i.Power));
        }

        private string GetCellValue(ICell cell)
        {
            Func<ICell, CellType, string> getCellValueByType = (icell, tp) =>
            {
                switch (tp)
                {
                    case CellType.Numeric:
                        return icell.NumericCellValue.ToString();
                    case CellType.String:
                        return icell.StringCellValue;
                    case CellType.Boolean:
                        return icell.BooleanCellValue.ToString();
                    case CellType.Blank:
                    case CellType.Unknown:
                    case CellType.Error:
                    default:
                        return "";
                }
            };

            if (cell.CellType == CellType.Formula)
            {
                return getCellValueByType(cell, cell.CachedFormulaResultType);
            }
            else
            {
                return getCellValueByType(cell, cell.CellType);
            }
        }

        private string GetCellFormula(ICell cell)
        {
            if (cell.CellType == CellType.Formula)
            {
                return cell.CellFormula;
            }
            else
            {
                switch (cell.CellType)
                {
                    case CellType.Numeric:
                        return cell.NumericCellValue.ToString();
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Blank:
                    case CellType.Unknown:
                    case CellType.Error:
                    default:
                        return "";
                }
            }
        }

        private string CalculateSum(IEnumerable<ICell> cells)
        {
            double sum = 0;
            foreach (var cell in cells)
            {
                if (cell.CellType == CellType.Numeric || cell.CellType == CellType.Formula)
                {
                    try
                    {
                        sum += cell.NumericCellValue;
                    }
                    catch (Exception)
                    {
                        return "";
                    }

                }
            }
            return sum.ToString();
        }

        private string GetCellPosition(ICell cell)
        {
            return string.Format("{0}{1}", (char)(cell.ColumnIndex + 'A'), cell.RowIndex + 1);
        }

        private class MyListBoxItem
        {
            public string Txt { get; set; }
            public TreeNode Node { get; set; }
            public MyListBoxItem(string txt, TreeNode node)
            {
                Txt = txt;
                Node = node;
            }
            public override string ToString()
            {
                return Txt;
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                var item = listBox1.Items[listBox1.SelectedIndex] as MyListBoxItem;
                treeView1.SelectedNode = item.Node;
            }
        }

        private void gv_SelectionChanged(object sender, EventArgs e)
        {
            var cells = gv.SelectedCells;
            if (cells.Count == 1)
            {
                var cell = cells[0];
                if (cell.ColumnIndex > 0 && cell.RowIndex < gv.Rows.Count - 1)
                {
                    label1.Text = GetCellFormula(GetICellOfThisCell(cell));
                }
            }
        }

        private ICell GetICellOfThisCell(DataGridViewCell cell)
        {
            var item = gv.Rows[cell.RowIndex].Tag as CItem;
            return typeof(CItem).GetProperty(gv.Columns[cell.ColumnIndex].Name).GetValue(item) as ICell;
        }

        private void gv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == gv.Rows.Count - 1) return;
            var citem = gv.Rows[e.RowIndex].Tag as CItem;
            BindRow(citem, gv.Rows[e.RowIndex]);
            RefreshSum();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
    }
}
