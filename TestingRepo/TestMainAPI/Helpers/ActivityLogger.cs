using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace TestAPI.Helpers
{
    public static class ActivityLogger
    {
        //private readonly IConfiguration _config;
        public static void ErrorFileLog(string Message)
        {
            var logPath = AppDomain.CurrentDomain.BaseDirectory + "App_Log";

            string _timestamp = DateTime.Now.ToString("yyyyMMddHHmmssffff");
            string logFile = "\\log_" + _timestamp + ".txt";
            try
            {

                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }
                var filePath = logPath + logFile;

                FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                StreamWriter sw = new StreamWriter(fs);
                try
                {
                    sw.Write(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + $"\tMessage: {Message}" + Environment.NewLine);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"ActivityLogger Failed: {ex.Message}");

                    sw.Close();
                    fs.Close();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ActivityLogger Failed: {ex.Message}");
            }
        }

        //public static void ErrorFileLog(string Message)
        //{
        //    //var logPth = _config.GetChildren GetValue<string>("AllowedOrigins");
        //    //var logPath = Directory.GetCurrentDirectory() + "\\Logs";
        //    var logPath = AppDomain.CurrentDomain.BaseDirectory + "App_Log";

        //    string _timestamp = DateTime.Now.ToString("yyyyMMdd");
        //    string logFile = "\\log_" + _timestamp + ".txt";

        //    if (!Directory.Exists(logPath))
        //    {
        //        Directory.CreateDirectory(logPath);
        //    }
        //    var filePath = logPath + logFile;

        //    if (!File.Exists(filePath))
        //    {
        //        File.Create(filePath);
        //    }

        //    FileStream fs = new FileStream(filePath, FileMode.Append, FileAccess.Write);
        //    StreamWriter sw = new StreamWriter(fs);

        //    try
        //    {


        //        sw.Write(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + $"\tMessage: {Message}" + Environment.NewLine);
        //        sw.Close();
        //        fs.Close();
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine($"ActivityLogger Failed: {ex.Message}");

        //        sw.Close();
        //        fs.Close();
        //    }
        //}

        public static void OldErrorFileLog(string Message)
        {
            //var logPth = _config.GetChildren GetValue<string>("AllowedOrigins");
            //var logPath = Directory.GetCurrentDirectory() + "\\Logs";
            var logPath = AppDomain.CurrentDomain.BaseDirectory + "App_Log";

            string _timestamp = DateTime.Now.ToString("yyyyMMdd");
            string logFile = "\\log_" + _timestamp + ".txt";

            try
            {
                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }
                var filePath = logPath + logFile;

                if (!File.Exists(filePath))
                {
                    File.Create(filePath);
                }

                FileStream fs = new FileStream(filePath, FileMode.Append, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);
                sw.Close();
                fs.Close();

                FileStream fs1 = new FileStream(filePath, FileMode.Append, FileAccess.Write);
                StreamWriter sw1 = new StreamWriter(fs1);

                sw1.Write(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + $"\tMessage: {Message}" + Environment.NewLine);
                sw1.Close();
                fs1.Close();
            }
            catch
            { }
        }
    }
}


