using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.ComponentModel.Com2Interop;

namespace CapitalStock
{
    class DataTick
    {
        public short Index { get; set; }
        public int Ptr { get; set; }
        public int Date { get; set; }
        public int Timehms { get; set; }
        public int Timemillins { get; set; }
        public int Bid { get; set; }
        public int Ask { get; set; }
        public int Close { get; set; }
        public int Qty { get; set; }
        public int Simulate { get; set; }
    }
}
