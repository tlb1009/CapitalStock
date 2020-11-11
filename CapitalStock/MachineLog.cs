using System;
using System.IO;

namespace MachineLog
{
    
    public class csLog
    {
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="logName">Log檔名稱</param>
        /// <param name="errlogName">Error Log檔名稱</param>
        private void initialFuntion(string logName,string errlogName)
        {
            bool dirExist = Directory.Exists(_filePath);
            if (!dirExist)
            {
                Directory.CreateDirectory(_filePath);
            }
            dirExist = Directory.Exists(_errFilePath);
            if (!dirExist)
            {
                Directory.CreateDirectory(_errFilePath);
            }
            bool fileExist = File.Exists(_filePath + logName);
            if (!fileExist)
            {
                File.AppendAllText(_filePath + logName, "");
            }
            fileExist = File.Exists(_errFilePath + errlogName);
            if (!fileExist)
            {
                File.AppendAllText(_errFilePath + errlogName, "");
            }
        }
        /// <summary>
        /// 建構子(無附加檔名)
        /// </summary>
        public csLog()
        {
            _fileName = "Log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            _errFileName = "ErrorLog" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            initialFuntion(_fileName,_errFileName);
        }
        /// <summary>
        /// 建構子(有附加檔名)
        /// </summary>
        /// <param name="_logName">要附加的檔名</param>
        public csLog(string _logName)
        {
            LogName = _logName;
            _fileName = "Log_" + LogName + "_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            _errFileName = "ErrorLog" + LogName + "_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            initialFuntion(_fileName, _errFileName);
        }
        #region --define--
        string LogName = "";
        private string _filePath = @"D:\Log\", _errFilePath = @"D:\ErrorLog\";
        private string _fileName = "", _errFileName = "";
        private string _logText="", _errorLogText="", _alarmhistory="";
        #endregion
        /// <summary>
        /// 操作記錄
        /// </summary>
        public string LogText
        {
            get { return _logText; }
            set
            {
                string ss =DateTime.Now.ToString("HH:mm:ss.fff=>")+ value;
                string name= _filePath + _fileName;
                SaveLog(name, ss);
                //_logText =_logText+ss+ Environment.NewLine;
            }
        }
        
        /// <summary>
        /// 當前警報
        /// </summary>
        public string ErrorLogText
        {
            get {return _errorLogText; }
            set
            {
                string ss = DateTime.Now.ToString("HH:mm:ss.fff=>") + value;
                SaveLog(_errFilePath + _errFileName, ss);
                _errorLogText = ss + Environment.NewLine + _errorLogText;
                _alarmhistory = ss + Environment.NewLine + _alarmhistory;
            }
        }

        /// <summary>
        /// 歷史警報
        /// </summary>
        public string AlarmHistory
        {
            get { return _alarmhistory; }
            set { _alarmhistory = value; }
        }

        /// <summary>
        /// 清除異常訊息
        /// </summary>
        public void ClearErrorLogText()
        {
            _errorLogText = "";
        }

        //存檔
        private void SaveLog(string ifullname,string text)
        {
            string ss = ifullname;
            File.AppendAllText(ss, text+Environment.NewLine);
        }
    }
}
