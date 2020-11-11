using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CapitalStock
{
    class classStockBaseData:classStockList
    {
        public string DYear { get; set; }
        public double Dividend { get; set; }
        public double EPS { get; set; }
    }
}
