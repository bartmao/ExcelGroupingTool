using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
        public static Dictionary<string, int> Columns { get; set; }

        public string SheetName { get; set; }

        public int RowNum { get; set; }

        public int Rows { get; set; }

        public string EquipName { get; set; }

        public string TypeDescription { get; set; }

        public string TypeCategory
        {
            get;set;
        }

        public List<ICell> VCells { get; set; } = new List<ICell>();

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
