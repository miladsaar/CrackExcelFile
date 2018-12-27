using System;
using System.Collections.Generic;
using System.Text;

namespace CrackExcelFile
{
    class TargetInfo
    {
        public string TargetName { get; set; }

        public string FileAddress { get; set; }

        public string TargetType { get; set; }

        public DateTime? CreateTime { get; set; }

        public CrackOption CrackOption { get; set; }
    }
}
