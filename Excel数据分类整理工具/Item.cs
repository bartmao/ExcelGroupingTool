using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel数据分类整理工具
{
    public class CItem
    {
        public CCell Designation { get; set; }

        public CCell[] Attributes { get; set; }

        public CCell CS { get; set; }

        public CCell RMB { get; set; }

    }

    public class CCell {
        public string Position { get; set; }

        public string Value { get; set; }
    }
}
