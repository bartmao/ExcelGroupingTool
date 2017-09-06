using NPOI.HSSF.UserModel;
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

        private void button1_Click(object sender, EventArgs e)
        {
            using (var stream = File.OpenRead("")) {
                var workbook = new HSSFWorkbook(stream);
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    if (Regex.IsMatch(sheet.SheetName, @"[\d|\.]+")) {
                        for (int j = 5; j < sheet.LastRowNum; j++)
                        {
                            var row = sheet.GetRow(j);
                            var designation = row.GetCell(2);
                        }
                    }
                }
            }
            
        }
    }
}
