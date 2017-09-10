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
            gv.AutoGenerateColumns = false;
        }
        string FilePath { get; set; }

        XSSFWorkbook workbook;

        ExcelTableParser Parser;

        List<ICell> ChangedCells = new List<ICell>();

        public List<CItem> Items { get; set; } = new List<CItem>();

        private void button1_Click(object sender, EventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Filter = "表格文件|*.xlsx";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                FilePath = dlg.FileName;
                var dir = Path.GetDirectoryName(FilePath);
                var fname = Path.GetFileName(FilePath);
                var tmpfn = Path.Combine(dir, "~" + fname);
                if (File.Exists(tmpfn)) File.Delete(tmpfn);
                File.Copy(FilePath, tmpfn);
                ReadExcel(tmpfn);
            }
        }

        private void ReadExcel(string fn)
        {
            var templateFile = AppDomain.CurrentDomain.BaseDirectory + "//Template1.txt";
            Parser = new ExcelTableParser(fn);
            var items = Parser.Parse(templateFile);
            workbook = Parser.Workbook;
            var categories = items.GroupBy(i => i.TypeCategory).OrderBy(g => g.Key);
            foreach (var cate in categories)
            {
                var node = treeView1.Nodes.Add(cate.Key);
                node.Tag = cate;
                node.Nodes.AddRange(cate.Select(c => new TreeNode()
                {
                    Tag = c,
                    Text = string.Format("{0}({1}:R{2})", c.TypeCategory, c.SheetName, c.RowNum + 1)
                }).ToArray());
            }

            foreach (var col in VItem.Columns)
            {
                gv.Columns.Add(col.Key, col.Key);
            }
        }

        private void gv_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            var gvRow = gv.Rows[e.RowIndex];
            var gvCell = gv.Rows[e.RowIndex].Cells[e.ColumnIndex];
            var value = e.Value.ToString();
            var vitem = gvRow.Tag as VItem;
            var vcell = gvCell.Tag as ICell;
            var oldVal = GetCellValue(vcell);

            // if using a formula
            if (value.StartsWith("="))
            {
                vcell.SetCellType(CellType.Formula);
                vcell.SetCellFormula(value.Substring(1, value.Length - 1));
                var eval = workbook.GetCreationHelper().CreateFormulaEvaluator();
                eval.EvaluateFormulaCell(vcell);
            }
            else
            {
                // only consider number/string here
                double v;
                if (double.TryParse(value, out v))
                {
                    vcell.SetCellType(CellType.Numeric);
                    vcell.SetCellValue(v);
                }
                else
                {
                    vcell.SetCellValue(value.ToString());
                }
            }

            if (string.IsNullOrWhiteSpace(oldVal))
                oldVal = "空值";
            var node = treeView1.SelectedNode;
            while (node.Parent != null)
            {
                node = node.Parent;
            }
            listBox1.Items.Add(new MyListBoxItem(string.Format("{0}!{1}从【{2}】改变到【{3}】\r\n", vitem.SheetName, GetCellPosition(vcell), oldVal, GetCellValue(vcell)), node));
            gvCell.Style.BackColor = Color.SkyBlue;
            //RefreshRow(gvRow);
            ChangedCells.Add(vcell);
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Parent == null)
            {
                var items = (e.Node.Tag as IGrouping<string, VItem>).ToList();
                Bind(items);
                RefreshSum();
            }
        }

        private void Bind(IEnumerable<VItem> items)
        {
            gv.Rows.Clear();
            gv.Tag = items;
            foreach (var vitem in items)
            {
                var row = BindRow(vitem);
                gv.Rows.Add(row);
            }

            var sumRow = new DataGridViewRow();
            sumRow.DefaultCellStyle.Font = new Font("宋体", 10, FontStyle.Bold);
            sumRow.ReadOnly = true;
            sumRow.Tag = null;
            sumRow.CreateCells(gv);
            gv.Rows.Add(sumRow);
            RefreshSum();
        }

        private DataGridViewRow BindRow(VItem vitem, DataGridViewRow row = null)
        {
            if (row == null)
            {
                row = new DataGridViewRow();
                row.Tag = vitem;
                row.CreateCells(gv);
            }

            for (int i = 0; i < VItem.Columns.Count; i++)
            {
                var vcell = vitem.VCells[i];
                if (ChangedCells.Contains(vcell)) row.Cells[i].Style.BackColor = Color.SkyBlue;
                var cell = row.Cells[i];

                if (gv.Columns[i].Name == "TYPE & SPECIFICATION")
                {
                    cell.Value = vitem.TypeDescription;
                    cell.ReadOnly = true;
                }
                else if (gv.Columns[i].Name == "EQUIP №")
                {
                    cell.Value = vitem.EquipName;
                    cell.ReadOnly = true;
                }
                else
                {
                    cell.Value = ExcelTableParser.GetCellValue(vitem.VCells[i]);
                    cell.Tag = vitem.VCells[i];
                }

            }

            return row;
        }

        private void RefreshSum()
        {
            var items = gv.Tag as IEnumerable<VItem>;

            var sumRow = gv.Rows[gv.Rows.Count - 1];
            foreach (DataGridViewCell cell in sumRow.Cells)
            {
                var isNumber = true;

                if (cell.ColumnIndex == 0)
                {
                    cell.Value = "汇总";
                    continue;
                }

                if (cell.ColumnIndex <= 5) continue;

                double sum = 0;
                for (int i = 0; i < gv.Rows.Count - 1; i++)
                {
                    try
                    {
                        var vcell = gv.Rows[i].Cells[cell.ColumnIndex].Tag as ICell;
                        if (vcell.CellType == CellType.String)
                        {
                            isNumber = false;
                            break;
                        }
                        else if (vcell.CellType != CellType.Blank)
                            sum += vcell.NumericCellValue;
                    }
                    catch (Exception)
                    {

                    }
                }
                if (isNumber)
                    cell.Value = sum;
            }
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
                    label1.Text = "";
                    var vcell = cell.Tag as ICell;
                    if (vcell != null)
                        label1.Text = GetCellFormula(vcell);
                }

                BindDetail(cell.RowIndex);
            }
        }

        private void gv_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == gv.Rows.Count - 1) return;
            var vitem = gv.Rows[e.RowIndex].Tag as VItem;
            BindRow(vitem, gv.Rows[e.RowIndex]);
            RefreshSum();

            BindDetail(gv.Rows[e.RowIndex].Index);
        }

        private void gv_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            BindDetail(e.RowIndex);
        }

        private void BindDetail(int rowNum)
        {
            if (rowNum == -1 || rowNum >= gv.Rows.Count - 1) return;
            var row = gv.Rows[rowNum];
            var vitem = row.Tag as VItem;
            var tb = Parser.GetRange(vitem.SheetName, vitem.RowNum, vitem.RowNum + vitem.Rows, VItem.Columns);
            gvDetail.DataSource = tb;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                using (var s = File.Create(FilePath))
                {
                    workbook.Write(s);
                    MessageBox.Show("保存成功");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("请先关闭此Excel文件");
            }

        }
    }
}
