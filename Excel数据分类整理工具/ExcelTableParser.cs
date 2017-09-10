using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Excel数据分类整理工具
{
    public class ExcelTableParser
    {
        public string FilePath { get; set; }

        public XSSFWorkbook Workbook { get; set; }

        public ExcelTableParser(string filePath)
        {
            FilePath = filePath;
        }

        public List<VItem> Parse(string templateFile)
        {
            var items = new List<VItem>();

            var template = File.ReadAllText(templateFile).Trim();
            var lines = template.Split('\n').Select(t => t.Trim()).ToList();
            var rowStart = int.Parse(lines[1]);
            //var key = int.Parse(lines[3]);
            var columns = new Dictionary<string, int>();
            for (int i = 3; i < lines.Count; i++)
            {
                var pair = lines[i].Split(',');
                columns.Add(pair[0].Trim(), pair[1].Trim()[0] - 'A');
            }
            VItem.Columns = columns;

            using (var stream = File.OpenRead(FilePath))
            {
                Workbook = new XSSFWorkbook(stream);
                VItem lastItem = null;
                for (int i = 0; i < Workbook.Count; i++)
                {
                    var sheet = Workbook.GetSheetAt(i);
                    if (Regex.IsMatch(sheet.SheetName, @"[\d|\.]+"))
                    {
                        for (int r = rowStart - 1; r < sheet.LastRowNum; r++)
                        {
                            var row = sheet.GetRow(r);
                            var equipName = GetCellValue(row.GetCell(columns["EQUIP №"]));
                            var subItemName = GetCellValue(row.GetCell(columns["SUB №"]));
                            var typeName = GetCellValue(row.GetCell(columns["TYPE & SPECIFICATION"])).Trim();
                            var identity = row.GetCell(columns["ITEM №"]);

                            if (string.IsNullOrWhiteSpace(GetCellValue(identity)))
                            {
                                if (lastItem != null)
                                {
                                    lastItem.EquipName += equipName;
                                    if (subItemName != null) lastItem.TypeDescription += typeName + "\r\n";
                                }
                                continue;
                            }

                            var item = new VItem(sheet.SheetName, r);
                            item.EquipName = equipName;
                            item.TypeDescription = typeName;
                            item.TypeCategory = Regex.Match(item.EquipName, "[a-zA-Z]+").Groups[0].Value;
                            if (lastItem != null) lastItem.Rows = item.RowNum - lastItem.RowNum;
                            foreach (var col in columns)
                            {
                                item.VCells.Add(row.GetCell(col.Value));
                            }
                            items.Add(item);
                            lastItem = item;
                        }

                        lastItem.Rows = sheet.LastRowNum - lastItem.RowNum;
                    }
                }

            }

            return items;
        }

        public static string GetCellValue(ICell cell)
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

        public DataTable GetRange(string sheetName, int firstRow, int lastRow, Dictionary<string, int> columns)
        {
            var sheet = Workbook.GetSheet(sheetName);
            var tb = new DataTable();
            foreach (var col in columns.Keys)
            {
                tb.Columns.Add(col);
            }
            for (int i = firstRow; i < lastRow; i++)
            {
                var vrow = sheet.GetRow(i);
                var row = tb.NewRow();
                var j = 0;
                foreach (var col in columns.Values)
                {
                    row[j++] = GetCellValue(vrow.GetCell(col));
                }
                tb.Rows.Add(row);
            }

            return tb;
        }
    }
}
