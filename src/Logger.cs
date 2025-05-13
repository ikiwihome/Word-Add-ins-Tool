using System;
using System.IO;
using System.Diagnostics;

namespace Word_AddIns
{
    /// <summary>
    /// 日志记录工具类，提供本地和远程日志记录功能
    /// </summary>
    public static class Logger
    {
        /// <summary>
        /// 日志文件路径(位于用户目录下的Word_AddIns.log)
        /// </summary>
        public static readonly string LogFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Word_AddIns.log");

        /// <summary>
        /// 用户ID存储文件路径(位于用户目录下的userid.txt)
        /// </summary>
        public static readonly string useridPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "userid.txt");

        private static string userId = null; // 缓存的用户ID

        /// <summary>
        /// 记录异常信息(不带异常对象)
        /// </summary>
        /// <param name="functionName">函数名称</param>
        /// <param name="errorName">错误名称</param>
        public static void LogException(string functionName, string errorName)
        {
            try
            {
                // 构造错误日志消息
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [ERROR] " +
                                   $"{functionName}, " +
                                   $"{errorName}";

                // 写入本地日志文件
                File.AppendAllText(LogFilePath, logMessage + Environment.NewLine);
            }
            catch (Exception logEx)
            {
                Debug.WriteLine($"Failed to log error: {logEx.Message}");
            }
        }

        /// <summary>
        /// 记录异常信息(带异常对象)
        /// </summary>
        /// <param name="ex">异常对象</param>
        /// <param name="functionName">函数名称</param>
        /// <param name="errorName">错误名称</param>
        public static void LogException(Exception ex, string functionName, string errorName)
        {
            try
            {
                // 构造包含异常详细信息的日志消息
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [ERROR] " +
                                   $"{functionName}, " +
                                   $"{errorName}, " +
                                   $"Error: {ex.Message}";

                // 写入本地日志文件
                File.AppendAllText(LogFilePath, logMessage + Environment.NewLine);
            }
            catch (Exception logEx)
            {
                Debug.WriteLine($"Failed to log exception: {logEx.Message}");
            }
        }
    }
}
