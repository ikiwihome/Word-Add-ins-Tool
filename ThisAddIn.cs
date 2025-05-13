using System;
using System.IO;
using System.Threading;
using System.Globalization;
using System.Resources;
using System.Reflection;
using Office = Microsoft.Office.Core;

namespace Word_AddIns
{
    /// <summary>
    /// Word插件主类，负责插件生命周期管理和核心功能
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>资源管理器，用于加载本地化字符串</summary>
        private readonly ResourceManager resourceManager = new ResourceManager("Word_AddIns.Properties.Resources", Assembly.GetExecutingAssembly());

        /// <summary>当前语言设置</summary>
        private CultureInfo currentLanguage = CultureInfo.GetCultureInfo("zh-CN");

        /// <summary>
        /// 插件启动事件处理程序
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Set initial language based on Office UI language
            UpdateLanguageSettings();

        }

        /// <summary>
        /// 更新语言设置，根据Office UI语言切换插件语言
        /// </summary>
        private void UpdateLanguageSettings()
        {
            try
            {
                int uiLanguage = Globals.ThisAddIn.Application.LanguageSettings.LanguageID[
                    Office.MsoAppLanguageID.msoLanguageIDUI];

                // Set culture based on language ID
                currentLanguage = uiLanguage == 2052 ?  // 2052 is Chinese (PRC)
                    CultureInfo.GetCultureInfo("zh-CN") :
                    CultureInfo.GetCultureInfo(uiLanguage);

                Thread.CurrentThread.CurrentUICulture = currentLanguage;

                // Refresh UI if ribbon is already loaded
                Globals.Ribbons.Ribbon1?.UpdateUIResources();
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "UpdateLanguageSettings", "Language settings update failed");
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"), "Language settings update failed");

                // Final fallback to Chinese if anything fails
                currentLanguage = CultureInfo.GetCultureInfo("zh-CN");
                Thread.CurrentThread.CurrentUICulture = currentLanguage;
            }
        }

        /// <summary>
        /// 获取本地化资源字符串
        /// </summary>
        /// <param name="resourceName">资源名称</param>
        /// <returns>本地化字符串</returns>
        public string GetResourceString(string resourceName)
        {
            return resourceManager.GetString(resourceName, currentLanguage);
        }

        /// <summary>
        /// 插件关闭事件处理程序
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // 检查是否有待处理的模板替换
                if (Globals.Ribbons.Ribbon1?.IsTemplateReplacePending == true)
                {
                    string batPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "replace_template.bat");

                    if (File.Exists(batPath))
                    {
                        // 使用隐藏窗口模式执行，延迟5秒确保Word完全退出
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = "cmd.exe",
                            Arguments = $"/C timeout 5 && \"{batPath}\"",
                            WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden,
                            CreateNoWindow = true
                        });
                    }
                    else
                    {
                        Logger.LogException("ThisAddIn_Shutdown", Globals.ThisAddIn.GetResourceString("Error_TemplateScriptMissing"));
                        Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                            Globals.ThisAddIn.GetResourceString("Error_TemplateScriptMissing"));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "ThisAddIn_Shutdown", Globals.ThisAddIn.GetResourceString("Error_TemplateScriptFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_TemplateScriptFailed"));
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
