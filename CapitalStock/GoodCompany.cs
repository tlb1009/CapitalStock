using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;

namespace CapitalStock
{
    class GoodCompany:classStockList
    {
        private string DatabasePath = $"d:/data/stocks.sqlite";
        private string ConnectionString = "DATA SOURCE=";
        public GoodCompany()
        {
            ConnectionString += DatabasePath;
        }
        public double avgDividend { get; set; }
        /// <summary>
        /// 在取得close price之前要先設定checkDateTime,取得那一天的價格
        /// </summary>
        public double ClosePrice 
        {
            get 
            {
                double d = 0;
                using (SQLiteConnection cn=new SQLiteConnection(ConnectionString))
                {
                    cn.Open();
                    SQLiteCommand command = cn.CreateCommand();
                    command.Connection = cn;
                    command.CommandText = $"select closeprice from Stock{ID} where bsdate='{checkDateTime}'";
                    var sdr = command.ExecuteScalar();
                    if (sdr!=null)
                    {
                        d = (double)sdr;
                    }
                }
                return d;
            }
        }
        public string checkDateTime { get; set; }
    }
}
