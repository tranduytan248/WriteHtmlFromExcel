using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WriteHtmlFromExcel.Models
{
    public class MergerModel
    {
        public int colStart { get; set; }
        public int colEnd { get; set; }
        public int rowStart { get; set; }
        public int rowEnd { get; set; }
        public string infoMerge { get; set; }
    }
}