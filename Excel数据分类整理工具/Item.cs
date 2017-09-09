using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel数据分类整理工具
{
    public class CItem
    {
        public string SheetName { get; set; }

        public string ItemNo { get; set; }

        public string Equip { get; set; }

        public ICell Designation { get; set; }

        public ICell[] Attributes { get; set; }

        public ICell Qty { get; set; }

        public ICell CS { get; set; }

        public ICell TS { get; set; }

        public ICell WeightTotal { get; set; }

        public ICell RMB { get; set; }

        public ICell Power { get; set; }

    }

    public class VItem
    {
        private Dictionary<string, int> Columns { get; set; }

        public string SheetName { get; set; }

        public int RowNum { get; set; }

        public List<ICell> Cells { get; set; } = new List<ICell>();

        public VItem(string sheetName, int rowNum)
        {
            SheetName = sheetName;
            RowNum = rowNum;
        }
    }

    public class CCell
    {
        public string Position { get; set; }

        public string Value { get; set; }
    }
}
