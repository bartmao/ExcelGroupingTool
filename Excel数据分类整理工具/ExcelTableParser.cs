using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
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
            var key = int.Parse(lines[3]);
            var columns = new Dictionary<string, int>();
            for (int i = 5; i < lines.Count; i++)
            {
                var pair = lines[i].Split(',');
                columns.Add(pair[0].Trim(), pair[1].Trim()[0] - 'A');
            }

            using (var stream = File.OpenRead(FilePath))
            {
                var workbook = new XSSFWorkbook(stream);
                for (int i = 0; i < workbook.Count; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    if (Regex.IsMatch(sheet.SheetName, @"[\d|\.]+"))
                    {
                        for (int r = rowStart - 1; r < sheet.LastRowNum; r++)
                        {
                            var row = sheet.GetRow(r);
                            var item = new VItem(sheet.SheetName, r);

                            foreach (var col in columns)
                            {
                                item.Cells.Add(row.GetCell(col.Value));
                            }
                            items.Add(item);
                        }
                    }
                }

            }

            return items;
        }

        public void Parse(string sheetName, int headerRowNum)
        {

            using (var stream = File.OpenRead(FilePath))
            {
                var workbook = new XSSFWorkbook(stream);
                var sheet = workbook.GetSheet(sheetName);
                var mergeds = new List<CellRangeAddress>();
                for (int i = 0; i < sheet.NumMergedRegions; i++)
                {
                    mergeds.Add(sheet.GetMergedRegion(i));
                }

                var row = sheet.GetRow(headerRowNum - 1);

                var j = 0;
                var columns = new Dictionary<string, int>();
                while (j < row.LastCellNum)
                {
                    var cell = row.GetCell(j);
                    var columnName = "";
                    var columnIndex = 0;
                    if (cell.IsMergedCell)
                    {
                        var merged = mergeds.Single(m => m.FirstColumn <= cell.ColumnIndex
                            && m.LastColumn >= cell.ColumnIndex
                            && m.FirstRow <= cell.RowIndex
                            && m.LastRow >= cell.RowIndex);
                        for (var m = merged.FirstRow; m <= merged.LastRow; m++)
                        {
                            if (columnName != "") break;
                            for (int n = merged.FirstColumn; n <= merged.LastColumn; n++)
                            {
                                var r = sheet.GetRow(m);
                                var c = r.GetCell(n);
                                if (c.CellType == CellType.String && !string.IsNullOrWhiteSpace(c.StringCellValue))
                                {
                                    columnName = c.StringCellValue;
                                    columnIndex = n;
                                    break;
                                }
                            }
                        }
                        columns.Add(columnName, columnIndex);
                        j += merged.LastColumn - merged.FirstColumn + 1;
                    }
                    else if (cell.CellType == CellType.String && !string.IsNullOrWhiteSpace(cell.StringCellValue))
                    {
                        columns.Add(cell.StringCellValue, j++);
                    }
                    else
                    {
                        j++;
                    }
                }
            }
        }
    }
}
