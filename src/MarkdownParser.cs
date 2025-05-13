using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace Word_AddIns
{
    /// <summary>
    /// Markdown解析器，用于将Markdown文件转换为Word文档内容
    /// </summary>
    public class MarkdownParser
    {
        private readonly Microsoft.Office.Interop.Word.Application _wordApp;

        /// <summary>
        /// 构造函数，初始化Markdown解析器
        /// </summary>
        /// <param name="wordApp">Word应用程序实例</param>
        public MarkdownParser(Microsoft.Office.Interop.Word.Application wordApp)
        {
            _wordApp = wordApp;
        }

        /// <summary>
        /// 获取Pandoc可执行文件路径
        /// </summary>
        /// <returns>返回找到的pandoc.exe路径，如果找不到则返回null</returns>
        private string GetPandocPath()
        {
            // 首先在Resources目录查找
            string pandocPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "pandoc.exe");
            if (File.Exists(pandocPath))
            {
                return pandocPath;
            }

            // 然后检查PATH中是否有pandoc命令
            using (Process process = new Process())
            {
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = "pandoc",
                    Arguments = "--version",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true
                };
                process.Start();
                if (process.WaitForExit(1000) && process.ExitCode == 0)
                {
                    return "pandoc"; // 返回命令名称
                }
            }


            // 最后在Program Files目录查找
            string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            pandocPath = Path.Combine(programFiles, "Pandoc", "pandoc.exe");
            if (File.Exists(pandocPath))
            {
                return pandocPath;
            }

            // 如果都找不到，提示用户下载
            Logger.LogException("GetPandocPath", Globals.ThisAddIn.GetResourceString("Error_PandocMissing"));
            MessageBox.Show(Globals.ThisAddIn.GetResourceString("Error_PandocMissing"),
                          Globals.ThisAddIn.GetResourceString("Title_Error"),
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Error);
            return null;
        }

        /// <summary>
        /// 异步解析Markdown文件并插入到Word文档
        /// </summary>
        /// <param name="markdownFilePath">Markdown文件路径</param>
        /// <param name="range">Word文档中要插入内容的范围</param>
        /// <returns>异步任务</returns>
        public async System.Threading.Tasks.Task MarkdownParserAsync(string markdownFilePath, Range range)
        {
            // 验证路径是否包含非法字符
            if (string.IsNullOrWhiteSpace(markdownFilePath))
            {
                Logger.LogException("ProcessAndInsertMarkdownAsync", string.Format(Globals.ThisAddIn.GetResourceString("Error_InvalidPathChars"),
                    Path.GetFileName(markdownFilePath)));
                return;
            }

            // 验证路径是否包含非法字符
            var invalidChars = markdownFilePath.Where(c => Path.GetInvalidPathChars().Contains(c)).ToArray();
            if (invalidChars.Length > 0)
            {
                Logger.LogException("ProcessAndInsertMarkdownAsync", string.Format(Globals.ThisAddIn.GetResourceString("Error_InvalidPathChars"),
                    string.Join(", ", invalidChars)));
                return;
            }

            // 验证markdown文件是否存在
            if (!File.Exists(markdownFilePath))
            {
                Logger.LogException("ProcessAndInsertMarkdownAsync", string.Format(Globals.ThisAddIn.GetResourceString("Error_FileNotFound"),
                    Path.GetFileName(markdownFilePath)));
                return;
            }

            // 调用Pandoc开始转换
            await StartPandocConvertAsync(markdownFilePath, range);
        }

        /// <summary>
        /// 启动Pandoc转换过程
        /// </summary>
        /// <param name="markdownFilePath">Markdown文件路径</param>
        /// <param name="range">Word文档中要插入内容的范围</param>
        /// <returns>异步任务</returns>
        private async System.Threading.Tasks.Task StartPandocConvertAsync(string markdownFilePath, Range range)
        {
            string pandocPath = GetPandocPath();
            string docxPath = Path.Combine(Path.GetTempPath(), "Addin_Markdown",
                $"{Path.GetFileNameWithoutExtension(markdownFilePath)}.docx");

            // 使用pandoc-reference.docx作为参考文档进行转换
            // 确保样式和格式与参考文档一致
            string referenceDoc = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Resources", "pandoc-reference.docx");

            // 计算文件总行数
            int totalLines = 0;

            try
            {
                using (var reader = File.OpenText(markdownFilePath))
                {
                    while (reader.ReadLine() != null)
                    {
                        totalLines++;
                    }
                }

                // 设置Pandoc参数
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = pandocPath,
                    Arguments = $"\"{markdownFilePath}\" -o \"{docxPath}\" --reference-doc \"{referenceDoc}\" " +
                        $"--resource-path \"{Path.GetDirectoryName(markdownFilePath)}\" " +
                        "-f markdown+emoji+backtick_code_blocks+east_asian_line_breaks+" +
                        "hard_line_breaks+lists_without_preceding_blankline-" +
                        "blank_before_blockquote-blank_before_header-ignore_line_breaks " +
                        "--trace=true",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetDirectoryName(markdownFilePath),
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    StandardErrorEncoding = System.Text.Encoding.UTF8,
                    StandardOutputEncoding = System.Text.Encoding.UTF8
                };

                // 如果md文件大于200行，则创建进度条
                if (totalLines > 200)
                {
                    // 创建进度条通知
                    Notification.CreateProgress(
                        Globals.ThisAddIn.GetResourceString("Title_ImportMarkdown"),
                        $"import_markdown_{DateTime.Now.Ticks}");
                }


                // 启动Pandoc进程进行转换
                // 使用TaskCompletionSource等待进程完成
                using (Process process = Process.Start(psi))
                {
                    // 使用StreamReader直接读取输出
                    int currentLine = 0;

                    // 创建一个单独的任务来读取标准错误输出
                    System.Threading.Tasks.Task errorReadTask = System.Threading.Tasks.Task.Run(() =>
                    {
                        string line;
                        while ((line = process.StandardError.ReadLine()) != null)
                        {
                            // 跳过不包含trace信息的行
                            if (!line.Contains("[trace] Parsed [Header")) continue;

                            // 跳过不包含行号信息的行
                            int atLineIndex = line.LastIndexOf("at line ");
                            if (atLineIndex < 0) continue;

                            // 尝试解析行号
                            string lineInfo = line.Substring(atLineIndex + 8);
                            if (!int.TryParse(lineInfo, out int lineNumber)) continue;

                            currentLine = lineNumber;
                            // 计算进度百分比
                            double progressPercentage = (double)currentLine / totalLines;

                            // 仅当文件较大时才更新进度条
                            if (totalLines <= 200) continue;

                            Notification.UpdateProgress(
                                progressPercentage,
                                    string.Format(Globals.ThisAddIn.GetResourceString("Msg_ProcessingFile"),
                                    Path.GetFileName(markdownFilePath)));
                        }
                    });

                    // 创建TaskCompletionSource用于异步等待进程退出
                    var tcs = new TaskCompletionSource<bool>();

                    // 启用进程事件触发
                    process.EnableRaisingEvents = true;

                    // 注册进程退出事件处理程序，进程退出时设置任务结果为完成
                    process.Exited += (s, e) => tcs.TrySetResult(true);

                    // 检查进程是否已经退出（可能在注册事件前就已经退出）
                    if (process.HasExited)
                    {
                        // 如果已经退出，直接设置任务结果为完成
                        tcs.TrySetResult(true);
                    }

                    // 异步等待进程退出完成
                    await tcs.Task;

                    // 如果转换出错，则输出错误内容
                    if (process.ExitCode != 0)
                    {
                        string fullError = process.StandardError.ReadToEnd();
                        string firstLineError = fullError.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? fullError;
                        Logger.LogException("StartPandocConvertAsync",
                            $"{Globals.ThisAddIn.GetResourceString("Error_PandocFailed")} Pandoc错误输出:{fullError}");

                        // 显示错误的通知 (错误)
                        Notification.Show(
                            Globals.ThisAddIn.GetResourceString("Title_ImportFailed"),
                            $"{Globals.ThisAddIn.GetResourceString("Error_PandocFailed")}:{firstLineError}");
                        return;
                    }
                }

                // 将转换后的临时docx文件内容插入到Word文档
                // 使用InsertFile方法直接插入文件内容，避免内存中的复制粘贴操作
                // 这种方法对大型文档特别有效，可以减少内存占用并提高处理效率
                range.InsertFile(docxPath);

                // 清理转换后的docx文件
                File.Delete(docxPath);

                // 导入成功，更新进度为100% (成功)
                // 如果md文件大于200行，则更新进度条，否则只显示通知
                if (totalLines > 200)
                {
                    Notification.UpdateProgress(1.0, Globals.ThisAddIn.GetResourceString("Msg_MarkdownImportSuccess"));
                }
                else
                {
                    Notification.Show(
                        Globals.ThisAddIn.GetResourceString("Title_Success"),
                        Globals.ThisAddIn.GetResourceString("Msg_MarkdownImportSuccess"));
                }
                return;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "StartPandocConvertAsync", Globals.ThisAddIn.GetResourceString("Error_PandocFailed"));
            }
        }

    }
}
