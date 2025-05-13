using Microsoft.Toolkit.Uwp.Notifications;
using Windows.UI.Notifications;
using System;

namespace Word_AddIns
{
    /// <summary>
    /// Windows通知工具类，提供显示普通通知和带进度条通知的功能
    /// </summary>
    public class Notification
    {
        private static ToastNotification _progressNotification; // 当前活动的进度条通知实例
        private static ToastNotification _currentNotification; // 当前活动的普通通知实例

        /// <summary>
        /// 显示普通通知
        /// </summary>
        /// <param name="title">通知标题</param>
        /// <param name="message">通知内容</param>
        public static void Show(string title, string message)
        {
            // 清除所有历史通知
            ToastNotificationManagerCompat.History.Clear();
            
            // 如果有正在显示的普通通知，先显式地移除它
            if (_currentNotification != null)
            {
                ToastNotificationManagerCompat.CreateToastNotifier().Hide(_currentNotification);
                _currentNotification = null;
            }
            
            // 如果有正在显示的进度条通知，先显式地移除它
            if (_progressNotification != null)
            {
                ToastNotificationManagerCompat.CreateToastNotifier().Hide(_progressNotification);
                _progressNotification = null;  // 清空引用
            }

            // 构建Toast通知内容
            var toastContent = new ToastContentBuilder()
                .AddText(title)    // 设置通知标题
                .AddText(message)  // 设置通知内容
                .GetToastContent();
                
            // 创建通知对象并设置属性
            var toast = new ToastNotification(toastContent.GetXml());
            string uniqueId = Guid.NewGuid().ToString();
            toast.Tag = uniqueId;
            toast.Group = "Addin_Notifications";
            toast.Priority = ToastNotificationPriority.High;
            toast.ExpirationTime = DateTime.Now.AddSeconds(1);
            
            // 显示通知并保存引用
            ToastNotificationManagerCompat.CreateToastNotifier().Show(toast);
            _currentNotification = toast;
        }

        /// <summary>
        /// 创建带进度条的通知
        /// </summary>
        /// <param name="title">通知标题</param>
        /// <param name="message">通知内容</param>
        /// <param name="progressId">进度条标识ID</param>
        public static void CreateProgress(string title, string progressId)
        {
            // 清除所有历史通知，避免通知堆积
            ToastNotificationManagerCompat.History.Clear();

            // 如果有正在显示的进度条通知，先显式地移除它
            if (_progressNotification != null)
            {
                ToastNotificationManagerCompat.CreateToastNotifier().Hide(_progressNotification);
                _progressNotification = null;  // 清空引用
            }

            // 构建带进度条的Toast内容
            var toastContent = new ToastContentBuilder()
                .AddText(title)    // 设置通知标题
                //.AddText(message)  // 设置通知内容
                .AddVisualChild(new AdaptiveProgressBar()  // 添加进度条控件
                {
                    Title = "进度",
                    Value = new BindableProgressBarValue("progressValue"),  // 绑定进度值
                    ValueStringOverride = new BindableString("progressValueString"),  // 绑定进度文本
                    Status = new BindableString("progressStatus")  // 绑定状态文本
                })
                .GetToastContent();

            // 创建Toast通知对象
            var toast = new ToastNotification(toastContent.GetXml());
            toast.Tag = progressId;  // 设置通知标签用于后续更新
            toast.Data = new NotificationData();  // 初始化通知数据
            toast.Data.Values["progressValue"] = "0";  // 初始进度0%
            toast.Data.Values["progressValueString"] = "0%";  // 初始进度文本
            toast.Data.Values["progressStatus"] = Globals.ThisAddIn.GetResourceString("Msg_InProcessing");  // 初始状态文本

            // 显示通知并保存引用
            ToastNotificationManagerCompat.CreateToastNotifier().Show(toast);
            _progressNotification = toast;  // 保存当前进度条通知引用
        }

        /// <summary>
        /// 更新进度条通知的进度和状态
        /// </summary>
        /// <param name="progress">进度值(0-1之间的小数)</param>
        /// <param name="status">可选的状态文本</param>
        /// <param name="timeoutSeconds">超时时间(秒)，进度100%时为3秒，否则为5秒</param>
        public static void UpdateProgress(double progress, string status = null, int timeoutSeconds = 5)
        {
            // 如果没有活动的进度条通知，直接返回
            if (_progressNotification == null) return;

            // 准备更新数据
            var data = new NotificationData();
            data.Values["progressValue"] = progress.ToString();  // 设置进度值(0-1)
            data.Values["progressValueString"] = $"{(int)(progress * 100)}%";  // 设置百分比文本

            // 如果有状态文本，更新状态
            if (!string.IsNullOrEmpty(status))
            {
                data.Values["progressStatus"] = status;
            }

            // 根据进度值设置超时时间(100%时为3秒，否则为5秒)
            int timeout = progress >= 1.0 ? 3 : timeoutSeconds;
            _progressNotification.ExpirationTime = DateTime.Now.AddSeconds(timeout);

            // 通过标签更新通知
            ToastNotificationManagerCompat.CreateToastNotifier().Update(data, _progressNotification.Tag);
        }
    }
}
