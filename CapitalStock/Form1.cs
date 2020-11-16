using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SKCOMLib;
using System.Data.SQLite;
using System.Threading;
using System.Data.Entity.Infrastructure.Interception;
using HtmlAgilityPack;
using mshtml;
using MachineLog;
using System.Collections.Specialized;
using System.Xml;
using System.Diagnostics.Eventing.Reader;

namespace CapitalStock
{
    public partial class Form1 : Form
    {
        string DatabasePath = $"d:/data/stocks.sqlite";
        string ConnectionString = "DATA SOURCE=";
        string DatabasePathDivident = $"d:/data/stockDividends.sqlite";
        string ConnectionStringDividend = "Data SOURCE=";
        SKCenterLibClass skCenter = new SKCenterLibClass();
        SKQuoteLibClass skQuote = new SKQuoteLibClass();
        SKSTOCK skStock = new SKSTOCK();
        List<classStockData> _Stocks = new List<classStockData>();
        List<classStockList> _StockList = new List<classStockList>();
        csLog mLog = new csLog();
        public Form1()
        {
            InitializeComponent();
            textBox5.Text = DateTime.Now.ToString("yyyy/MM/dd");
            textBox7.Text = DateTime.Now.ToString("yyyy/MM/dd");
            textBox8.Text = DateTime.Now.ToString("yyyyMMdd");
            //initialize();
        }
        private void initialize()
        {
            string id = textBox9.Text;
            string pw = textBox10.Text;
            ConnectionString = ConnectionString + DatabasePath + ";  Version = 3;";
            ConnectionStringDividend = ConnectionStringDividend + DatabasePathDivident + ";  Version = 3;";
            int result = skCenter.SKCenterLib_Login(id, pw);
            this.Text = result.ToString();
            var filepath = $"d:/data";
            if (!Directory.Exists(filepath))
            {
                Directory.CreateDirectory(filepath);
            }
            filepath = $"d:/data/bigdeal";
            if (!Directory.Exists(filepath))
            {
                Directory.CreateDirectory(filepath);
            }
            var filename = $"d:/data/stocksNo.txt";
            if (File.Exists(filename))
            {
                string[] readlist = File.ReadAllLines(filename);
                foreach (var x in readlist)
                {
                    string[] ss = x.Split(',');
                    classStockList sklist = new classStockList();
                    sklist.ID = ss[0];
                    sklist.Name = ss[1];
                    _StockList.Add(sklist);
                    classStockData skData = new classStockData();
                    skData.ID = sklist.ID;
                    skData.Name = sklist.Name;
                    skData.OpDate = DateTime.Now.ToString("yyyy/MM/dd");
                    skData.Open = 0;
                    skData.High = 0;
                    skData.Low = 0;
                    skData.Close = 0;
                    skData.Vol = 0;
                    _Stocks.Add(skData);
                }
            }
            
            skQuote.OnNotifyStockList += SkQuote_OnNotifyStockList;
            skQuote.OnNotifyKLineData += SkQuote_OnNotifyKLineData;
            skQuote.OnNotifyQuote += SkQuote_OnNotifyQuote;
            skQuote.OnNotifyTicks += SkQuote_OnNotifyTicks;
            skQuote.OnNotifyHistoryTicks += SkQuote_OnNotifyHistoryTicks;
        }
        private void SkQuote_OnConnection(int nKind, int nCode)
        {
            this.Text = nKind.ToString();
        }
        #region --Form
        //上線
        private void button1_Click(object sender, EventArgs e)
        {
            if (this.Text.Equals("0") || this.Text.Equals("3002"))
            {
                skQuote.OnConnection += SkQuote_OnConnection;
                int result = skQuote.SKQuoteLib_EnterMonitor();
            }
        }
        //離線
        private void button2_Click(object sender, EventArgs e)
        {
            if (this.Text.Equals("3003"))
            {
                int result = skQuote.SKQuoteLib_LeaveMonitor();
            }
        }
        //取得商品列表
        private void button10_Click(object sender, EventArgs e)
        {
            _StockList.Clear();
            skQuote.SKQuoteLib_RequestStockList(0);
        }
        //存商品列表
        private void button11_Click(object sender, EventArgs e)
        {
            string s = "";
            foreach (var x in _StockList)
            {
                s += x.ID + "," + x.Name + Environment.NewLine;
            }
            string filename = $"d:/data/stocksNo.txt";
            bool fExist = File.Exists(filename);
            if (fExist)
            {
                File.Delete(filename);
            }
            File.AppendAllText(filename, s);
        }

        private void SkQuote_OnNotifyStockList(short sMarketNo, string bstrStockData)
        {
            if (this.Text.Equals("3003"))
            {
                string s = bstrStockData;
                string[] ss = s.Split(';');
                foreach (var x in ss)
                {
                    string[] ss2 = x.Split(',');
                    bool tryp = int.TryParse(ss2[0], out int i);
                    if (ss2[0].Length == 4 && tryp && i > 1000)
                    {
                        classStockList sk = new classStockList();
                        sk.ID = ss2[0];
                        sk.Name = ss2[1];
                        listBox1.Items.Add(sk.ID + " " + sk.Name);
                        _StockList.Add(sk);
                    }
                }
            }
        }
        #endregion
        #region Page1
        List<classStockData> myStocks;
        //create database
        private void button3_Click(object sender, EventArgs e)
        {
            CreateDatabase(DatabasePath);
            CreateDatabase(DatabasePathDivident);
        }
        //create table
        private void button19_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection cn = new SQLiteConnection(ConnectionString))
            {
                cn.Open();
                //創建數據表
                SQLiteCommand command = cn.CreateCommand();
                command.Connection = cn;
                // 添加參數
                command.Parameters.Add(command.CreateParameter());
                SQLiteTransaction transaction = cn.BeginTransaction();
                try
                {
                    foreach (var x in _StockList)
                    {
                        command.CommandText = $"CREATE TABLE IF NOT EXISTS STOCK{x.ID} (" +
                    $"bsDate TEXT PRIMARY KEY NOT NULL," +
                    $"OPENPRICE REAL NOT NULL," +
                    $"HIGHPRICE REAL NOT NULL," +
                    $"LOWPRICE REAL NOT NULL," +
                    $"CLOSEPRICE REAL NOT NULL," +
                    $"VOL INT NOT NULL)";
                        command.ExecuteNonQuery();
                    }
                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
                finally
                {
                    myStocks = null;
                }
            }
        }
        //取得日K
        private void button4_Click(object sender, EventArgs e)
        {
            myStocks = new List<classStockData>();
            listBox2.Items.Clear();
            foreach (var x in _StockList)
            {
                skQuote.SKQuoteLib_RequestKLine(x.ID, 4, 1);
            }
        }
        //存檔database
        private void button8_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection cn = new SQLiteConnection(ConnectionString))
            {
                cn.Open();
                //創建數據表
                SQLiteCommand command = cn.CreateCommand();
                command.Connection = cn;
                // 添加參數
                command.Parameters.Add(command.CreateParameter());
                SQLiteTransaction transaction = cn.BeginTransaction();
                try
                {
                    foreach (var x in myStocks)
                    {
                        command.CommandText = $"INSERT INTO STOCK{x.ID} (bsdate, openprice,highprice,lowprice,closeprice,vol) " +
                       $"VALUES ('{x.OpDate}',{x.Open},{x.High},{x.Low},{x.Close},{x.Vol});";
                        command.ExecuteNonQuery();
                    }
                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
                finally
                {
                    myStocks = null;
                }
            }
        }

        private void SkQuote_OnNotifyKLineData(string bstrStockNo, string bstrData)
        {
            string[] ss = bstrData.Split(',');
            if (ss.Length == 6)
            {
                //資料分解
                string dt = ss[0];
                double dOpen = 0;
                double.TryParse(ss[1], out dOpen);
                double dHigh = 0;
                double.TryParse(ss[2], out dHigh);
                double dLow = 0;
                double.TryParse(ss[3], out dLow);
                double dClose = 0;
                double.TryParse(ss[4], out dClose);
                int iVol = 0;
                int.TryParse(ss[5], out iVol);
                classStockData skdata = new classStockData();
                skdata.ID = bstrStockNo;
                skdata.OpDate = dt;
                skdata.Open = dOpen;
                skdata.High = dHigh;
                skdata.Low = dLow;
                skdata.Close = dClose;
                skdata.Vol = iVol;
                myStocks.Add(skdata);
            }
        }
        #endregion
        #region Page2
        //訂閱項目
        //取得股票股利AVG
        private void button18_Click(object sender, EventArgs e)
        {
            listBox5.Items.Clear();
            string index = textBox6.Text,checkDay=textBox7.Text;
            double avg = 0,closeprice=0,preAvg=0;
            using (SQLiteConnection cn=new SQLiteConnection(ConnectionStringDividend))
            {
                cn.Open();
                SQLiteCommand command = cn.CreateCommand();
                command.Connection = cn;
                command.CommandText=$"select dividend from dividend{index} order by year desc limit 8";
                SQLiteDataReader sdr = command.ExecuteReader();
                List<double> ds = new List<double>();
                while (sdr.Read())
                {
                    double d = sdr.GetDouble(0);
                    ds.Add(d);
                }
                avg = CalculationAvg(ds);
            }
            using (SQLiteConnection cn=new SQLiteConnection(ConnectionString))
            {
                cn.Open();
                SQLiteCommand command = cn.CreateCommand();
                command.Connection = cn;
                command.CommandText = $"select closeprice from stock{index} where bsdate='{checkDay}'";
                SQLiteDataReader sdr = command.ExecuteReader();
                while (sdr.Read())
                {
                    closeprice = sdr.GetDouble(0);
                }
                preAvg = avg / closeprice;
            }
            listBox5.Items.Add($"8年平均股利={avg},收盤價={closeprice},股利/收盤價={preAvg}");
        }
        private void button5_Click(object sender, EventArgs e)
        {
            string bstrStocks = "";
            foreach (var x in _StockList)
            {
                bstrStocks += x.ID + ',';
            }
            bstrStocks = bstrStocks.Remove(bstrStocks.Length - 1, 1);
            short pageNo = 1;
            skQuote.SKQuoteLib_RequestStocks(ref pageNo, bstrStocks);
        }
        //顯示符合項目
        private void button7_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox5.Items.Clear();
            foreach (var x in _Stocks)
            {
                string showString = $"{x.ID}{x.Name}\t開盤:{x.Open.ToString("0000.00")}\t最高:{x.High.ToString("0000.00")}\t最低:{x.Low.ToString("0000.00")}\t收盤:{x.Close.ToString("0000.00")}\t成交量:{x.Vol.ToString("000000")}";
                double ratio = 0;
                double topRatio = 0;
                double.TryParse(textBox2.Text, out ratio);
                double.TryParse(textBox1.Text, out topRatio);
                //尋找資料表
                bool tableExists = CheckTableExists($"STOCK{x.ID}");
                if (tableExists)
                {
                    var querstring = $"select vol from STOCK{x.ID} order by bsdate desc limit 10";
                    DataTable dt = GetDataTable(DatabasePath, querstring);
                    List<int> total = new List<int>();
                    foreach (DataRow y in dt.Rows)
                    {
                        total.Add((int)y.ItemArray[0]);
                    }

                    double avg = total.Average();
                    if (x.Vol < avg * topRatio && x.Close < 200 && x.Close > x.Open)
                    {
                        if (x.Vol > avg * ratio && avg * x.Close >= 30000)
                        {
                            showString = showString + "\t10日均量" + avg.ToString();
                            listBox1.Items.Add(showString);
                        }
                    }
                    if (x.Close < 200 && x.Close < x.Open)
                    {
                        if (x.Vol > avg * ratio && avg * x.Close >= 30000)
                        {
                            showString = showString + "\t10日均量" + avg.ToString();
                            listBox5.Items.Add(showString);
                        }
                    }
                }
            }
        }
        //顯示所有項目
        private void button9_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            foreach (var x in _Stocks)
            {
                string showString = $"{x.ID}{x.Name}\t開盤:{x.Open.ToString("0000.00")}\t最高:{x.High.ToString("0000.00")}\t最低:{x.Low.ToString("0000.00")}\t收盤:{x.Close.ToString("0000.00")}\t成交量:{x.Vol.ToString("000000")}";
                //尋找資料表
                bool tableExists = CheckTableExists($"STOCK{x.ID}");
                if (tableExists)
                {
                    var querstring = $"select vol from STOCK{x.ID} order by bsdate desc limit 10";
                    DataTable dt = GetDataTable(DatabasePath, querstring);
                    List<int> total = new List<int>();
                    foreach (DataRow y in dt.Rows)
                    {
                        total.Add((int)y.ItemArray[0]);
                    }
                    double avg = total.Average();
                    showString = showString + "\t10日均量" + avg.ToString();
                    listBox1.Items.Add(showString);
                }
            }
        }
        //存今日K
        private void button6_Click(object sender, EventArgs e)
        {
            using (SQLiteConnection cn = new SQLiteConnection(ConnectionString))
            {
                cn.Open();
                //創建數據表
                SQLiteCommand command = cn.CreateCommand();
                command.Connection = cn;
                // 添加參數
                command.Parameters.Add(command.CreateParameter());
                SQLiteTransaction transaction = cn.BeginTransaction();
                try
                {
                    foreach (var x in _Stocks)
                    {
                        command.CommandText = $"CREATE TABLE IF NOT EXISTS STOCK{x.ID} (" +
                    $"bsDate TEXT PRIMARY KEY NOT NULL," +
                    $"OPENPRICE REAL NOT NULL," +
                    $"HIGHPRICE REAL NOT NULL," +
                    $"LOWPRICE REAL NOT NULL," +
                    $"CLOSEPRICE REAL NOT NULL," +
                    $"VOL INT NOT NULL)";
                        command.ExecuteNonQuery();

                        command.CommandText = $"INSERT INTO STOCK{x.ID} (bsdate, openprice,highprice,lowprice,closeprice,vol) " +
                       $"VALUES ('{x.OpDate}',{x.Open},{x.High},{x.Low},{x.Close},{x.Vol});";
                        command.ExecuteNonQuery();
                    }
                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }
        }
        //刪除今日資料
        private void button16_Click(object sender, EventArgs e)
        {
            string today = DateTime.Now.ToString("yyyy/MM/dd");
            bool[] b = new bool[_StockList.Count];
            using (SQLiteConnection cn=new SQLiteConnection(ConnectionString))
            {
                cn.Open();
                SQLiteCommand command = cn.CreateCommand();
                command.Connection = cn;
                for (int i = 0; i < _StockList.Count; i++)
                {
                    command.CommandText = $"select * from stock{_StockList[i].ID} where bsdate='{today}'";
                    SQLiteDataReader sdr = command.ExecuteReader();
                    b[i] = sdr.HasRows;
                    sdr.Close();
                }
            }
            using (SQLiteConnection cn=new SQLiteConnection(ConnectionString))
            {
                cn.Open();
                SQLiteCommand command = cn.CreateCommand();
                command.Connection = cn;
                for (int i = 0; i < b.Length; i++)
                {
                    if (b[i])
                    {
                        command.CommandText = $"delete from stock{_StockList[i].ID} where bsdate='{today}'";
                        command.ExecuteNonQuery();
                    }
                }
            }
        }
        private void SkQuote_OnNotifyQuote(short sMarketNo, short sStockIdx)
        {
            //按INDEX取得DATA
            skQuote.SKQuoteLib_GetStockByIndex(sMarketNo, sStockIdx, ref skStock);

            bool dataNotExistsFlag = true;
            foreach (var x in _Stocks)
            {
                //更新資料
                if (skStock.bstrStockNo == x.ID)
                {
                    x.Open = (double)skStock.nOpen / 100;
                    x.High = (double)skStock.nHigh / 100;
                    x.Low = (double)skStock.nLow / 100;
                    x.Close = (double)skStock.nClose / 100;
                    x.Vol = skStock.nTQty;
                    dataNotExistsFlag = false;
                    break;
                }
            }
            //新增資料
            if (dataNotExistsFlag)
            {
                classStockData skdata = new classStockData();
                skdata.ID = skStock.bstrStockNo;
                skdata.Name = skStock.bstrStockName;
                skdata.Vol = skStock.nTQty;
                _Stocks.Add(skdata);
            }
        }
        #endregion
        #region database
        /// <summary>建立資料庫連線</summary>
        /// <param name="database">資料庫名稱</param>
        /// <returns></returns>
        public SQLiteConnection OpenConnection(string database)
        {
            var conntion = new SQLiteConnection() { ConnectionString = $"Data Source={database}" };
            if (conntion.State == ConnectionState.Open) conntion.Close();
            conntion.Open();
            return conntion;
        }
        /// <summary>建立新資料庫</summary>
        /// <param name="database">資料庫名稱</param>
        public void CreateDatabase(string database)
        {
            var connection = new SQLiteConnection()
            {
                ConnectionString = $"Data Source={database}"
            };
            connection.Open();
            connection.Close();
        }
        /// <summary>建立新資料表</summary>
        /// <param name="database">資料庫名稱</param>
        /// <param name="sqlCreateTable">建立資料表的 SQL 語句</param>
        public void CreateTable(string database, string sqlCreateTable)
        {
            var connection = OpenConnection(database);
            //connection.Open();
            var command = new SQLiteCommand(sqlCreateTable, connection);
            var mySqlTransaction = connection.BeginTransaction();
            try
            {
                command.Transaction = mySqlTransaction;
                command.ExecuteNonQuery();
                mySqlTransaction.Commit();
            }
            catch (Exception ex)
            {
                mySqlTransaction.Rollback();
                throw (ex);
            }
            if (connection.State == ConnectionState.Open) connection.Close();
        }

        /// <summary>新增\修改\刪除資料</summary>
        /// <param name="database">資料庫名稱</param>
        /// <param name="sqlManipulate">資料操作的 SQL 語句</param>
        public void Manipulate(string database, string sqlManipulate)
        {
            var connection = OpenConnection(database);
            var command = new SQLiteCommand(sqlManipulate, connection);
            var mySqlTransaction = connection.BeginTransaction();
            try
            {
                command.Transaction = mySqlTransaction;
                command.ExecuteNonQuery();
                mySqlTransaction.Commit();
            }
            catch (Exception ex)
            {
                mySqlTransaction.Rollback();
                throw (ex);
            }
            if (connection.State == ConnectionState.Open) connection.Close();
        }
        /// <summary>讀取資料</summary>
        /// <param name="database">資料庫名稱</param>
        /// <param name="sqlQuery">資料查詢的 SQL 語句</param>
        /// <returns></returns>
        public DataTable GetDataTable(string database, string sqlQuery)
        {
            var connection = OpenConnection(database);
            var dataAdapter = new SQLiteDataAdapter(sqlQuery, connection);
            var myDataTable = new DataTable();
            var myDataSet = new DataSet();
            myDataSet.Clear();
            dataAdapter.Fill(myDataSet);
            myDataTable = myDataSet.Tables[0];
            if (connection.State == ConnectionState.Open) connection.Close();
            return myDataTable;
        }
        /// <summary>
        /// 檢查有無TABLE
        /// </summary>
        /// <param name="tableName">資料表名稱</param>
        /// <returns></returns>
        private bool CheckTableExists(string tableName)
        {
            bool b = false;
            var querString = $"select * from sqlite_master where type='table' and name='{tableName}'";
            using (SQLiteConnection cn = new SQLiteConnection(ConnectionString))
            {
                cn.Open();
                using (SQLiteCommand com = new SQLiteCommand(querString, cn))
                {
                    SQLiteDataReader reader = com.ExecuteReader();
                    b = reader.HasRows;
                }
                cn.Close();
            }
            return b;
        }
        /// <summary>
        /// 檢查有無資料
        /// </summary>
        /// <param name="tableName">資料表名稱</param>
        /// <param name="bsDate">資料列日期</param>
        /// <returns></returns>
        private bool CheckDataRowExists(string tableName, string bsDate)
        {
            bool b = false;
            var querString = $"select * from {tableName} where bsDate='{bsDate}'";
            using (SQLiteConnection cn = new SQLiteConnection(ConnectionString))
            {
                cn.Open();
                using (SQLiteCommand com = new SQLiteCommand(querString, cn))
                {
                    SQLiteDataReader reader = com.ExecuteReader();
                    b = reader.HasRows;
                }
                cn.Close();
            }
            return b;
        }


        #endregion
        #region Page3
        int buyQty = 0;
        int sellQty = 0;
        int totalQty = 0;
        int avgQty = 0;
        int avgCount = 0;
        private void button13_Click(object sender, EventArgs e)
        {
            string s = textBox3.Text;
            short pageNo = 2;
            skQuote.SKQuoteLib_RequestTicks(ref pageNo, s);
        }
        private void button15_Click(object sender, EventArgs e)
        {
            listBox4.Items.Add(DateTime.Now.ToString("HH:mm:ss") + "\t" + (buyQty - sellQty).ToString());
        }
        private void SkQuote_OnNotifyTicks(short sMarketNo, short sIndex, int nPtr, int nDate, int nTimehms, int nTimemillismicros, int nBid, int nAsk, int nClose, int nQty, int nSimulate)
        {
            //listBox3.Items.Clear();
            if ((((nTimehms >= 90000 && !checkBox2.Checked) || (nTimehms >= 84300 && checkBox2.Checked)) && nTimehms < 132500) || (checkBox1.Checked && (nTimehms >= 150000 || nTimehms < 60000)))
            {
                avgCount++;
                if (avgCount > 1)
                {
                    totalQty += nQty;
                    avgQty = totalQty / avgCount;
                    DataTick dt = new DataTick();
                    dt.Index = sIndex;
                    dt.Ptr = nPtr;
                    dt.Date = nDate;
                    dt.Timehms = nTimehms;
                    dt.Timemillins = nTimemillismicros;
                    dt.Bid = nBid;
                    dt.Ask = nAsk;
                    dt.Close = nClose;
                    dt.Qty = nQty;
                    dt.Simulate = nSimulate;
                    if (dt.Close <= dt.Bid)
                    {
                        sellQty += dt.Qty;
                    }
                    if (dt.Close >= dt.Ask)
                    {
                        buyQty += dt.Qty;
                    }
                    listBox3.Items.Add($"代號:{dt.Index}\t時間{dt.Timehms}\t現價:{dt.Close}\t現量:{dt.Qty}\t{buyQty}-{sellQty}={buyQty - sellQty}\t均量{avgQty}");
                    if (dt.Qty >= avgQty * 5)
                    {
                        listBox4.Items.Add($"代號:{dt.Index}\t時間{dt.Timehms}\t{dt.Close}\t{buyQty}-{sellQty}={buyQty - sellQty}");
                    }
                }
            }
            listBox3.TopIndex = listBox3.Items.Count - 1;
            listBox4.TopIndex = listBox4.Items.Count - 1;
        }
        private void SkQuote_OnNotifyHistoryTicks(short sMarketNo, short sStockIdx, int nPtr, int nDate, int nTimehms, int nTimemillismicros, int nBid, int nAsk, int nClose, int nQty, int nSimulate)
        {
            //listBox3.Items.Clear();
            if ((((nTimehms >= 90000 && !checkBox2.Checked) || (nTimehms >= 84300 && checkBox2.Checked)) && nTimehms < 132500) || (checkBox1.Checked && (nTimehms >= 150000 || nTimehms < 60000)))
            {
                avgCount++;
                if (avgCount > 1)
                {
                    totalQty += nQty;
                    avgQty = totalQty / avgCount;
                    DataTick dt = new DataTick();
                    dt.Index = sStockIdx;
                    dt.Ptr = nPtr;
                    dt.Date = nDate;
                    dt.Timehms = nTimehms;
                    dt.Timemillins = nTimemillismicros;
                    dt.Bid = nBid;
                    dt.Ask = nAsk;
                    dt.Close = nClose;
                    dt.Qty = nQty;
                    dt.Simulate = nSimulate;
                    if (dt.Close <= dt.Bid)
                    {
                        sellQty += dt.Qty;
                    }
                    if (dt.Close >= dt.Ask)
                    {
                        buyQty += dt.Qty;
                    }
                    listBox3.Items.Add($"代號:{dt.Index}\t時間{dt.Timehms}\t現價:{dt.Close}\t現量:{dt.Qty}\t{buyQty}-{sellQty}={buyQty - sellQty}\t均量{avgQty}");
                    if (dt.Qty >= avgQty * 5)
                    {
                        listBox4.Items.Add($"代號:{dt.Index}\t時間{dt.Timehms}\t{dt.Close}\t{buyQty}-{sellQty}={buyQty - sellQty}");
                    }
                }
            }
        }
        #endregion
        #region Page4
        WebBrowser web1 = new WebBrowser();
        //等待網頁回傳完
        private void waitTillLoad(WebBrowser webBrControl)
        {
            WebBrowserReadyState loadStatus;
            int waittime = 100000;
            int counter = 0;
            while (true)
            {
                loadStatus = webBrControl.ReadyState;
                Application.DoEvents();
                if ((counter > waittime) || (loadStatus == WebBrowserReadyState.Uninitialized) || (loadStatus == WebBrowserReadyState.Loading) || (loadStatus == WebBrowserReadyState.Interactive))
                {
                    break;
                }
                counter++;
            }

            counter = 0;
            while (true)
            {
                loadStatus = webBrControl.ReadyState;
                Application.DoEvents();
                if (loadStatus == WebBrowserReadyState.Complete && webBrControl.IsBusy != true)
                {
                    break;
                }
                counter++;
            }
        }
        //取得個股十年股利
        private HtmlAgilityPack.HtmlDocument GetDividendListWithBrowser(String url)
        {
            web1.ScriptErrorsSuppressed = true;
            web1.Navigate(url);

            waitTillLoad(this.web1);

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            var documentAsIHtmlDocument3 = (mshtml.IHTMLDocument3)web1.Document.DomDocument;
            StringReader sr = new StringReader(documentAsIHtmlDocument3.documentElement.outerHTML);
            doc.Load(sr);

            return doc;
        }
        private void button12_Click(object sender, EventArgs e)
        {
            GetAllStocksDividend();
        }
        private void GetAllStocksDividend()
        {
            List<classStockBaseData> sds;
            foreach (var x in _StockList)
            {
                int count = 0;
                button12.Text = x.ID + x.Name;
                sds = new List<classStockBaseData>();
                //股利頁面
                var url = $"https://www.cmoney.tw/finance/f00027.aspx?s=" +
                    $"{x.ID}";
                GetNothingRetry:
                var document = GetDividendListWithBrowser(url);
                for (int i = 0; i < 10; i++)
                {
                    classStockBaseData sd = new classStockBaseData();
                    var nodeYear = document.DocumentNode.SelectSingleNode($"//tr[{i + 2}]/td[1]");
                    var nodeCost = document.DocumentNode.SelectSingleNode($"//tr[{i + 2}]/td[2]");
                    if (nodeYear != null && nodeCost != null)
                    {
                        sd.DYear = nodeYear.InnerText;
                        string s = nodeCost.InnerText;
                        double d = 0;
                        bool pOK = double.TryParse(s, out d);
                        if (pOK)
                        {
                            sd.Dividend = d;
                            sds.Add(sd);
                        }
                        else
                        {
                            mLog.LogText = $"小數轉換失敗:{x.ID}{x.Name}:{sd.DYear}";
                            sd.Dividend = 0;
                            sds.Add(sd);
                        }
                    }
                    else
                    {
                        //無資料
                        if (i==0)
                        {
                            if (count<5)
                            {
                                Thread.Sleep(10000);
                                count++;
                                goto GetNothingRetry;
                            }
                            else
                            {
                                mLog.LogText = $"多次取不到資料:{x.ID}";
                                count = 0;
                                break;
                            }
                        }
                        else
                        {
                            count = 0;
                            break;
                        }
                    }
                }

                //database
                using (SQLiteConnection cn = new SQLiteConnection(ConnectionStringDividend))
                {
                    cn.Open();
                    //創建數據表
                    SQLiteCommand command = cn.CreateCommand();
                    command.Connection = cn;
                    // 添加參數
                    command.Parameters.Add(command.CreateParameter());
                    SQLiteTransaction transaction = cn.BeginTransaction();
                    try
                    {
                        foreach (var y in sds)
                        {
                            command.CommandText = $"CREATE TABLE IF NOT EXISTS dividend{x.ID} (" +
                        $"Year TEXT PRIMARY KEY NOT NULL," +
                        $"Dividend REAL NOT NULL)";
                            command.ExecuteNonQuery();
                            command.CommandText = $"INSERT INTO dividend{x.ID} (Year, Dividend) " +
                           $"VALUES ('{y.DYear}',{y.Dividend});";
                            command.ExecuteNonQuery();
                        }
                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
                Thread.Sleep(300);
            }
        }
        //單一個股取股利
        private void button14_Click(object sender, EventArgs e)
        {
            List<classStockBaseData> sds = new List<classStockBaseData>();
            string x = textBox4.Text;
            var url = $"https://www.cmoney.tw/finance/f00027.aspx?s=" +
                $"{x}";
            var document = GetDividendListWithBrowser(url);
            for (int i = 0; i < 10; i++)
            {
                classStockBaseData sd = new classStockBaseData();
                var nodeYear = document.DocumentNode.SelectSingleNode($"//tr[{i + 2}]/td[1]");
                var nodeCost = document.DocumentNode.SelectSingleNode($"//tr[{i + 2}]/td[2]");
                if (nodeYear != null && nodeCost != null)
                {
                    sd.DYear = nodeYear.InnerText;
                    string s = nodeCost.InnerText;
                    double d = 0;
                    bool pOK = double.TryParse(s, out d);
                    if (pOK)
                    {
                        sd.Dividend = d;
                    }
                    else
                    {
                        sd.Dividend = 0;
                        mLog.LogText = $"小數轉換失敗:{x}:{sd.DYear}";
                    }
                    sds.Add(sd);
                }
                else
                {
                    //無資料
                    mLog.LogText = $"無當年資料:{x}:{i + 2}";
                    break;
                }
            }
            //database
            using (SQLiteConnection cn = new SQLiteConnection(ConnectionStringDividend))
            {
                cn.Open();
                //創建數據表
                SQLiteCommand command = cn.CreateCommand();
                command.Connection = cn;
                // 添加參數
                command.Parameters.Add(command.CreateParameter());
                SQLiteTransaction transaction = cn.BeginTransaction();
                try
                {
                    foreach (var y in sds)
                    {
                        command.CommandText = $"CREATE TABLE IF NOT EXISTS dividend{x} (" +
                    $"Year TEXT PRIMARY KEY NOT NULL," +
                    $"Dividend REAL NOT NULL)";
                        command.ExecuteNonQuery();
                        command.CommandText = $"INSERT INTO dividend{x} (Year, Dividend) " +
                       $"VALUES ('{y.DYear}',{y.Dividend});";
                        command.ExecuteNonQuery();
                    }
                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }
        }
        //選出股票
        private void button17_Click(object sender, EventArgs e)
        {
            listBox6.Items.Clear();
            listBox7.Items.Clear();
            List<GoodCompany> companys = new List<GoodCompany>();
            foreach (var x in _StockList)
            {
                //取得8年股利
                string queryString = $"select dividend from dividend{x.ID} order by year desc limit 8";
                DataTable dt = GetDataTable(DatabasePathDivident, queryString);
                //8年以上公司
                if (dt.Rows.Count>=8)
                {
                    List<double> ds = new List<double>();
                    bool CanAdd = true;
                    foreach (DataRow y in dt.Rows)
                    {
                        double d = (double)y.ItemArray[0];
                        //每年配息大於1塊
                        if (d>=1)
                        {
                            ds.Add(d);
                        }
                        else
                        {
                            CanAdd = false;
                            break;
                        }
                    }
                    //8年配息平均大於3塊
                    double avg = CalculationAvg(ds);
                    //add company
                    GoodCompany gc = new GoodCompany();
                    if (CanAdd && avg>=3)
                    {
                        gc.ID = x.ID;
                        gc.Name = x.Name;
                        gc.avgDividend = avg;
                        companys.Add(gc);
                        gc.checkDateTime = textBox5.Text;
                        listBox6.Items.Add($"{gc.ID}{gc.Name}  \t平均股利{gc.avgDividend.ToString("0.00")}\t今日收盤價{gc.ClosePrice}");
                    }
                }
            }
            List<GoodCompany> myCompanys = new List<GoodCompany>();
            foreach (var z in companys)
            {
                //現價/平均股息>5%
                double pre = z.avgDividend / z.ClosePrice;
                if (pre >= 0.05)
                {
                    myCompanys.Add(z);
                    listBox7.Items.Add($"{z.ID}{z.Name}  \t平均股利{z.avgDividend.ToString("0.00")}\t殖利率{pre.ToString("0.0000")}");
                }
            }
        }
        private double CalculationAvg(List<double> values)
        {
            double k = 0.3;
            double newValue = 0;
            double oldValue = 0;
            for (int i = values.Count; i > 0; i--)
            {
                newValue = values[i-1];
                if (newValue >= oldValue)
                {
                    oldValue = (newValue - oldValue) * k + oldValue;
                }
                else
                {
                    oldValue = (newValue - oldValue)*(k+0.3) + oldValue;
                }
            }
            return oldValue;
        }


        #endregion

        #region Page5
        private void button20_Click(object sender, EventArgs e)
        {
            string x = textBox8.Text;
            var url = $"https://www.twse.com.tw/fund/T86?response=html&date={x}&selectType=ALL";
            var document = GetDividendListWithBrowser(url);
            var stringa = document.DocumentNode.SelectNodes("//tr");
            if (stringa!=null)
            {
                for (int i = 2; i < stringa.Count; i++)
                {
                    var stringb = document.DocumentNode.SelectNodes($"//tr[{i}]//td");
                    if (stringb != null && stringb.Count <= 19)
                    {
                        
                        int nu;
                        bool b = int.TryParse(stringb.First().InnerText, out nu);
                        if ( b && nu < 9999)
                        {
                            string s1 = "";
                            foreach (var y in stringb)
                            {
                                s1 += y.InnerText + ";";
                            }
                            listBox8.Items.Add(s1);
                            File.AppendAllText($"d:/data/bigdeal/{x}.csv", s1+Environment.NewLine);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("no data");
            }
        }
        #endregion

        private void button21_Click(object sender, EventArgs e)
        {
            initialize();
        }
    }
}
