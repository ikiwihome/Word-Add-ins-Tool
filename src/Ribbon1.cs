﻿using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace Word_AddIns
{
    /// <summary>
    /// Word插件功能区实现类，提供文档处理、水印管理等功能
    /// </summary>
    public partial class Ribbon1
    {
        /// <summary>标记是否需要替换Word模板</summary>
        public bool IsTemplateReplacePending { get; set; }

        /// <summary>水印检查定时器</summary>
        private System.Windows.Forms.Timer watermarkCheckTimer;

        /// <summary>当前水印文本</summary>
        private string currentWatermarkText = "Confidential"; // Default value, will be updated from resources

        /// <summary>当前水印字体大小</summary>
        private int currentWatermarkSize = 54;

        /// <summary>水印是否加粗</summary>
        private bool isWatermarkBold = true;

        /// <summary>水印是否斜体</summary>
        private bool isWatermarkItalic = false;

        /// <summary>
        /// 更新界面资源，从资源文件加载所有按钮标签和提示文本
        /// </summary>
        public void UpdateUIResources()
        {
            // Update tab label
            tabAddinTool.Label = Globals.ThisAddIn.GetResourceString("RibbonTab_Label");

            // Update group labels
            grpImport.Label = Globals.ThisAddIn.GetResourceString("ImportGroup_Label");
            grpDesign.Label = Globals.ThisAddIn.GetResourceString("DesignGroup_Label");
            grpExport.Label = Globals.ThisAddIn.GetResourceString("ExportGroup_Label");
            grpTemplate.Label = Globals.ThisAddIn.GetResourceString("TemplateGroup_Label");
            grpMore.Label = Globals.ThisAddIn.GetResourceString("MoreGroup_Label");

            // Update all button labels and tooltips
            UpdateButtonResources(btnImportMarkdown, "btnImportMarkdown");
            UpdateButtonResources(btnImportText, "btnImportText");
            UpdateButtonResources(btnUniformFont, "btnUniformFont");
            UpdateButtonResources(btnUniformTableStyle, "btnUniformTableStyle");
            UpdateButtonResources(btnRemoveStyles, "btnRemoveStyles");
            UpdateButtonResources(btnApplyTemplate, "btnApplyTemplate");
            UpdateButtonResources(mnuWatermark, "mnuWatermark");
            UpdateButtonResources(btnAddWatermark, "btnAddWatermark");
            UpdateButtonResources(btnModifyWatermark, "btnModifyWatermark");
            UpdateButtonResources(btnRemoveWatermark, "btnRemoveWatermark");
            UpdateButtonResources(btnExportPDF, "btnExportPDF");
            UpdateButtonResources(btnNewDocument, "btnNewDocument");
            UpdateButtonResources(btnUpdateFields, "btnUpdateFields");
            UpdateButtonResources(btnSetDefault, "btnSetDefault");
            UpdateButtonResources(btnAbout, "btnAbout");
            UpdateButtonResources(btnFeedback, "btnFeedback");
        }

        /// <summary>
        /// 更新单个按钮的资源文本
        /// </summary>
        /// <param name="button">功能区按钮</param>
        /// <param name="resourcePrefix">资源前缀</param>
        private void UpdateButtonResources(RibbonButton button, string resourcePrefix)
        {
            button.Label = Globals.ThisAddIn.GetResourceString($"{resourcePrefix}_Label");
            button.ScreenTip = Globals.ThisAddIn.GetResourceString($"{resourcePrefix}_ScreenTip");
            button.SuperTip = Globals.ThisAddIn.GetResourceString($"{resourcePrefix}_SuperTip");
        }

        /// <summary>
        /// 更新菜单的资源文本
        /// </summary>
        /// <param name="menu">功能区菜单</param>
        /// <param name="resourcePrefix">资源前缀</param>
        private void UpdateButtonResources(RibbonMenu menu, string resourcePrefix)
        {
            menu.Label = Globals.ThisAddIn.GetResourceString($"{resourcePrefix}_Label");
            menu.ScreenTip = Globals.ThisAddIn.GetResourceString($"{resourcePrefix}_ScreenTip");
            menu.SuperTip = Globals.ThisAddIn.GetResourceString($"{resourcePrefix}_SuperTip");
        }

        /// <summary>
        /// 获取按钮图标
        /// </summary>
        /// <param name="buttonId">按钮ID</param>
        /// <returns>按钮图像</returns>
        private System.Drawing.Image GetButtonImage(string buttonId)
        {
            return Properties.Resources.ResourceManager.GetObject(buttonId) as System.Drawing.Image;
        }

        /// <summary>
        /// 导入Markdown文件按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private async void OnImportMarkdown(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Documents.Count == 0)
                {
                    Logger.LogException("OnImportMarkdown", Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"), Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    return;
                }

                using (var openDialog = new OpenFileDialog())
                {
                    openDialog.Filter = "Markdown文件 (*.md)|*.md";
                    openDialog.Multiselect = false;
                    openDialog.Title = Globals.ThisAddIn.GetResourceString("Title_SelectMarkdownFiles");

                    if (openDialog.ShowDialog() == DialogResult.OK)
                    {
                        var doc = app.ActiveDocument;
                        var range = doc.Content;
                        var processor = new MarkdownParser(app);

                        // 移动到文档末尾
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

                        foreach (string filePath in openDialog.FileNames)
                        {
                            if (!File.Exists(filePath))
                            {
                                Logger.LogException("OnImportMarkdown", string.Format(Globals.ThisAddIn.GetResourceString("Error_FileNotFound"),
                                    Path.GetFileName(filePath)));

                                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                                    string.Format(Globals.ThisAddIn.GetResourceString("Error_FileNotFound"),
                                    Path.GetFileName(filePath)));
                                continue;
                            }

                            // 检查是否已在空白页开头
                            bool isAtPageStart = range.Information[Microsoft.Office.Interop.Word.WdInformation.wdFirstCharacterColumnNumber] == 1 &&
                                                range.Information[Microsoft.Office.Interop.Word.WdInformation.wdFirstCharacterLineNumber] == 1;
                            bool isPageBlank = string.IsNullOrWhiteSpace(range.Text) || range.Characters.Count <= 1;

                            // 如果不是空白页开头则插入分页符
                            if (!(isAtPageStart && isPageBlank))
                            {
                                range.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
                            }

                            // 处理并插入Markdown内容
                            await processor.MarkdownParserAsync(filePath, range);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnImportMarkdown", Globals.ThisAddIn.GetResourceString("Error_ImportFailed"));
            }
        }

        /// <summary>
        /// 导入纯文本文件按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void OnImportText(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Documents.Count == 0)
                {
                    Logger.LogException("OnImportText", Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    return;
                }

                using (var openDialog = new OpenFileDialog())
                {
                    openDialog.Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*";
                    openDialog.Multiselect = true;
                    openDialog.Title = Globals.ThisAddIn.GetResourceString("Title_SelectTextFiles");

                    if (openDialog.ShowDialog() == DialogResult.OK)
                    {
                        var doc = app.ActiveDocument;
                        var range = doc.Content;
                        int successCount = 0;

                        // 移动到文档末尾，如果需要从当前光标处导入则注释掉下面这行代码
                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);

                        int totalFiles = openDialog.FileNames.Length;
                        int currentFile = 0;

                        foreach (string filePath in openDialog.FileNames)
                        {
                            currentFile++;
                            double fileProgress = 0;

                            if (!File.Exists(filePath))
                            {
                                Logger.LogException("OnImportText", string.Format(Globals.ThisAddIn.GetResourceString("Error_FileNotFound"),
                                    Path.GetFileName(filePath)));

                                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                                    string.Format(Globals.ThisAddIn.GetResourceString("Error_FileNotFound"),
                                    Path.GetFileName(filePath)));
                                continue;
                            }

                            // 更新进度条显示当前文件名
                            Notification.UpdateProgress(
                                (currentFile - 1) / (double)totalFiles,
                                string.Format(Globals.ThisAddIn.GetResourceString("Msg_ProcessingFile"),
                                Path.GetFileName(filePath)));

                            // 检查是否已在空白页开头
                            bool isAtPageStart = range.Information[Microsoft.Office.Interop.Word.WdInformation.wdFirstCharacterColumnNumber] == 1 &&
                                                range.Information[Microsoft.Office.Interop.Word.WdInformation.wdFirstCharacterLineNumber] == 1;
                            bool isPageBlank = string.IsNullOrWhiteSpace(range.Text) || range.Characters.Count <= 1;

                            // 如果不是空白页开头则插入分页符
                            if (!(isAtPageStart && isPageBlank))
                            {
                                range.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
                            }

                            // 创建进度条通知
                            Notification.CreateProgress(
                                Globals.ThisAddIn.GetResourceString("Title_ImportText"),
                                $"import_text_{DateTime.Now.Ticks}");

                            // 使用分段方式读取大文本文件
                            EncodingUtility.ReadTextFileWithEncodingChunked(filePath, (chunk, progress) =>
                            {
                                Debug.WriteLine($"正在导入文件: {Path.GetFileName(filePath)}, 进度: {progress:F2}%");

                                // 计算总体进度：(已完成文件数 + 当前文件进度)/总文件数
                                fileProgress = progress / 100.0;
                                double totalProgress = (currentFile - 1 + fileProgress) / totalFiles;

                                // 更新进度条
                                Notification.UpdateProgress(
                                    totalProgress,
                                    string.Format(Globals.ThisAddIn.GetResourceString("Msg_ProcessingFile"),
                                    Path.GetFileName(filePath)));

                                try
                                {
                                    // 检查range是否有效
                                    if (range != null && range.Start >= 0 && range.End >= 0)
                                    {
                                        // 插入当前分段内容
                                        range.InsertAfter(chunk);

                                        // 移动到新插入内容的末尾
                                        range.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                                    }
                                    else
                                    {
                                        Debug.WriteLine("无效的range对象，无法插入内容");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"插入内容时出错: {ex.Message}");
                                    throw; // 重新抛出异常以便外层捕获
                                }

                                // 更新UI让Word有机会处理其他事件
                                Globals.ThisAddIn.Application.System.Cursor = Microsoft.Office.Interop.Word.WdCursorType.wdCursorWait;
                                Globals.ThisAddIn.Application.ScreenRefresh();
                                Globals.ThisAddIn.Application.System.Cursor = Microsoft.Office.Interop.Word.WdCursorType.wdCursorNormal;
                            });

                            successCount++;
                        }

                        // 成功导入x/y个文本文件
                        Notification.UpdateProgress(1.0, string.Format(Globals.ThisAddIn.GetResourceString("Msg_TextImportSuccess"),
                            successCount, openDialog.FileNames.Length));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnImportText", Globals.ThisAddIn.GetResourceString("Error_ImportFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_ImportFailed"));
            }
        }

        /// <summary>
        /// 统一文档字体按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void OnUniformFont(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Documents.Count == 0)
                {
                    Logger.LogException("OnUniformFont", Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    return;
                }

                // 添加文档刷新和延迟
                var doc = app.ActiveDocument;

                // 确保文档完成所有挂起操作
                Thread.Sleep(1000); // 1秒延迟

                string chineseFont = Globals.ThisAddIn.GetResourceString("DefaultFontName");

                // 直接设置整个文档的字体，不依赖样式
                doc.Content.Font.Name = chineseFont;
                doc.Content.Font.NameFarEast = chineseFont;

                // 确保所有文本范围都应用字体
                foreach (Microsoft.Office.Interop.Word.Range range in doc.StoryRanges)
                {
                    range.Font.Name = chineseFont;
                    range.Font.NameFarEast = chineseFont;
                }

                Notification.Show(
                    Globals.ThisAddIn.GetResourceString("Title_Success"),
                    string.Format(Globals.ThisAddIn.GetResourceString("Msg_UniformFontSuccess"), chineseFont));
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnUniformFont", Globals.ThisAddIn.GetResourceString("Error_UniformFontFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_UniformFontFailed"));
            }
        }

        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        /// <summary>
        /// 统一表格样式按钮点击事件
        /// </summary>
        private void OnUniformTableStyle(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Documents.Count == 0)
                {
                    Logger.LogException("OnUniformTableStyle", Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    return;
                }

                var doc = app.ActiveDocument;

                // Apply styles to all tables
                if (doc.Tables.Count > 0)
                {
                    foreach (Microsoft.Office.Interop.Word.Table table in doc.Tables)
                    {
                        // Apply built-in table style based on language
                        object styleName;
                        if (app.LanguageSettings.LanguageID[Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI] == 2052) // Chinese
                        {
                            styleName = "网格型";
                        }
                        else // Default to English
                        {
                            styleName = "Table Grid";
                        }
                        table.set_Style(ref styleName);

                        // Set table width to 100% of page
                        table.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        table.PreferredWidth = 100f;
                    }
                    Notification.Show(
                        Globals.ThisAddIn.GetResourceString("Title_Success"),
                        Globals.ThisAddIn.GetResourceString("Msg_UniformTableStyleSuccess"));
                }
                else
                {
                    Notification.Show(
                        Globals.ThisAddIn.GetResourceString("Title_Notice"),
                        Globals.ThisAddIn.GetResourceString("Msg_NoTable"));
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnUniformTableStyle",
                    Globals.ThisAddIn.GetResourceString("Error_UniformTableStyleFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_UniformTableStyleFailed"));
            }
        }

        /// <summary>
        /// 移除样式按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        // 保护列表 - 不应删除的样式
        private static readonly string[] ProtectedStyles = new[]
        {
            "Normal", "Body Text", "First Paragraph", "Compact", "Title", "Subtitle",
            "Author", "Date", "Abstract", "Abstract Title", "Bibliography",
            "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5",
            "Heading 6", "Heading 7", "Heading 8", "Heading 9",
            "Block Text", "Footnote Block Text", "Source Code", "Footnote Text",
            "Definition Term", "Definition", "Caption", "Table Caption", "Image Caption",
            "Figure", "Captioned Figure", "TOC Heading", "Default Paragraph Font",
            "Body Text Char", "Verbatim Char", "Footnote Reference", "Hyperlink",
            "Section Number", "有序列表", "正文居中"
        };

        private void OnRemoveStyles(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 获取Word应用程序实例
                var app = Globals.ThisAddIn.Application;

                // 检查是否有打开的文档
                if (app.Documents.Count == 0)
                {
                    Logger.LogException("OnRemoveStyles", Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    return;
                }

                var doc = app.ActiveDocument;

                // 1. 收集所有将被删除的样式名称
                var stylesToDelete = new System.Collections.Generic.List<string>();
                foreach (Microsoft.Office.Interop.Word.Style style in doc.Styles)
                {
                    // 跳过内置样式
                    if (style.BuiltIn) continue;

                    string styleName = style.NameLocal;

                    // 跳过保护列表中的样式
                    if (ProtectedStyles.Contains(styleName) ||
                        styleName.StartsWith("Style First Paragraph") ||
                        styleName.StartsWith("Table") ||
                        styleName.EndsWith("Tok"))
                    {
                        continue;
                    }

                    stylesToDelete.Add(styleName);
                }

                // 如果没有要删除的样式，直接返回
                if (stylesToDelete.Count == 0)
                {
                    Notification.Show(
                        Globals.ThisAddIn.GetResourceString("Title_Notice"),
                        Globals.ThisAddIn.GetResourceString("Msg_NoStylesToRemove"));
                    return;
                }

                // 2. 显示样式删除对话框
                using (var dialog = new StyleDeleteDialog(stylesToDelete))
                {
                    var result = dialog.ShowDialog();

                    // 3. 如果用户确认，则删除选中的样式
                    if (result == DialogResult.OK && dialog.SelectedStyles.Count > 0)
                    {
                        foreach (Microsoft.Office.Interop.Word.Style style in doc.Styles)
                        {
                            if (dialog.SelectedStyles.Contains(style.NameLocal))
                            {
                                style.Delete();
                            }
                        }
                    }
                    else
                    {
                        return; // 用户取消操作
                    }
                }

                Notification.Show(
                    Globals.ThisAddIn.GetResourceString("Title_Success"),
                    Globals.ThisAddIn.GetResourceString("Msg_StyleRemovedSuccess"));
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnRemoveStyles", Globals.ThisAddIn.GetResourceString("Error_DeleteStyleFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_DeleteStyleFailed"));
            }
        }

        /// <summary>
        /// 应用模板样式按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void OnApplyTemplate(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Documents.Count == 0)
                {
                    Logger.LogException("OnApplyTemplate", Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    return;
                }

                var doc = app.ActiveDocument;

                // 1. 获取模板文件路径
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Normal.dotm");

                if (!File.Exists(templatePath))
                {
                    Logger.LogException("OnApplyTemplate", string.Format(Globals.ThisAddIn.GetResourceString("Error_TemplateNotFound"),
                        templatePath));

                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        string.Format(Globals.ThisAddIn.GetResourceString("Error_TemplateNotFound"),
                        templatePath));
                    return;
                }

                // 2. 应用模板样式
                doc.CopyStylesFromTemplate(templatePath);

                Notification.Show(
                    Globals.ThisAddIn.GetResourceString("Title_Success"),
                    Globals.ThisAddIn.GetResourceString("Msg_TemplateApplySuccess"));
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnApplyTemplate", Globals.ThisAddIn.GetResourceString("Error_ApplyTemplateFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_ApplyTemplateFailed"));
            }
        }

        /// <summary>
        /// 添加水印按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void OnAddWatermark(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Documents.Count == 0)
                {
                    Logger.LogException("OnAddWatermark", Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    return;
                }

                // Show dialog with current values
                var style = isWatermarkBold ? FontStyle.Bold : FontStyle.Regular;
                if (isWatermarkItalic) style |= FontStyle.Italic;

                using (var dialog = new WatermarkDialog(currentWatermarkText, currentWatermarkSize, style))
                {
                    if (dialog.ShowDialog() == DialogResult.OK &&
                        !string.IsNullOrWhiteSpace(dialog.WatermarkText))
                    {
                        var doc = app.ActiveDocument;

                        if (doc.Sections.Count == 1)
                        {
                            AddWatermarkToSection(doc.Sections[1], dialog);
                        }
                        else if (doc.Sections.Count > 1)
                        {
                            // 从第2节开始给所有节添加水印
                            for (int i = 2; i <= doc.Sections.Count; i++)
                            {
                                AddWatermarkToSection(doc.Sections[i], dialog);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnAddWatermark", Globals.ThisAddIn.GetResourceString("Error_AddWatermarkFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_AddWatermarkFailed"));
            }
        }

        private void AddWatermarkToSection(Microsoft.Office.Interop.Word.Section section, WatermarkDialog dialog)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;

                // 获取页眉范围
                var header = section.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                // 创建文本框水印
                var watermark = header.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    Left: 0, Top: 0,
                    Width: (float)(32 * 28.35), // 32厘米转换为磅
                    Height: (float)(12 * 28.35)); // 12厘米转换为磅

                // 设置水印文本和样式
                watermark.Name = "watermark";
                var textRange = watermark.TextFrame.TextRange;
                textRange.Text = dialog.WatermarkText;
                textRange.Font.Name = "Microsoft YaHei";
                textRange.Font.Size = dialog.FontSize;
                textRange.Font.Bold = dialog.CurrentFontStyle.HasFlag(FontStyle.Bold) ? 1 : 0;
                textRange.Font.Italic = dialog.CurrentFontStyle.HasFlag(FontStyle.Italic) ? 1 : 0;
                textRange.Font.Fill.ForeColor.RGB = 0xE7E6E6;
                textRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // 垂直居中位置
                watermark.TextFrame.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;


                // 设置文本框格式
                watermark.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                watermark.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                watermark.Rotation = 315;
                watermark.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapNone;
                watermark.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);

                // 设置水印位置
                watermark.Left = (float)Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter;
                watermark.Top = (float)Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter;
                watermark.RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
                watermark.RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
                watermark.WrapFormat.AllowOverlap = -1; // 允许重叠

                btnAddWatermark.Enabled = false;
                btnModifyWatermark.Enabled = true;
                btnRemoveWatermark.Enabled = true;

                Notification.Show(
                    Globals.ThisAddIn.GetResourceString("Title_Success"),
                    Globals.ThisAddIn.GetResourceString("Msg_WatermarkAdded"));
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "AddWatermarkToSection", "Add Watermark To Section Failed");
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"), "Add Watermark To Section Failed");
            }
        }

        /// <summary>
        /// 修改水印按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void OnModifyWatermark(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Documents.Count == 0)
                {
                    Logger.LogException("OnModifyWatermark", Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    return;
                }

                var doc = app.ActiveDocument;

                // 使用辅助函数查找水印
                var watermark = FindWatermarkShape();

                if (watermark == null)
                {
                    Logger.LogException("OnModifyWatermark", Globals.ThisAddIn.GetResourceString("Error_WatermarkNotFound"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_WatermarkNotFound"));
                    return;
                }

                // Get current watermark properties
                var textRange = watermark.TextFrame.TextRange;
                currentWatermarkText = textRange.Text;
                currentWatermarkSize = (int)textRange.Font.Size;
                isWatermarkBold = textRange.Font.Bold == -1;
                isWatermarkItalic = textRange.Font.Italic == -1;

                // Show dialog with current values
                var style = isWatermarkBold ? FontStyle.Bold : FontStyle.Regular;
                if (isWatermarkItalic) style |= FontStyle.Italic;

                using (var dialog = new WatermarkDialog(currentWatermarkText, currentWatermarkSize, style))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        if (!string.IsNullOrWhiteSpace(dialog.WatermarkText))
                        {
                            // Update watermark properties
                            textRange.Text = dialog.WatermarkText;
                            textRange.Font.Size = dialog.FontSize;
                            textRange.Font.Bold = dialog.CurrentFontStyle.HasFlag(FontStyle.Bold) ? -1 : 0;
                            textRange.Font.Italic = dialog.CurrentFontStyle.HasFlag(FontStyle.Italic) ? -1 : 0;

                            btnAddWatermark.Enabled = false;
                            btnRemoveWatermark.Enabled = true;

                            Notification.Show(
                                Globals.ThisAddIn.GetResourceString("Title_Success"),
                                Globals.ThisAddIn.GetResourceString("Msg_WatermarkModified"));
                        }
                        else
                        {
                            // Remove if Text is empty
                            OnRemoveWatermark(sender, e);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnModifyWatermark", Globals.ThisAddIn.GetResourceString("Error_ModifyWatermarkFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_ModifyWatermarkFailed"));
            }
        }

        /// <summary>
        /// 移除水印按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void OnRemoveWatermark(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Documents.Count == 0)
                {
                    Logger.LogException("OnRemoveWatermark", Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    return;
                }

                var doc = app.ActiveDocument;
                bool removed = false;

                // 使用辅助函数查找水印
                var watermark = FindWatermarkShape();
                if (watermark != null)
                {
                    watermark.Delete();
                    removed = true;
                }

                if (removed)
                {
                    btnAddWatermark.Enabled = true;
                    btnModifyWatermark.Enabled = false;
                    btnRemoveWatermark.Enabled = false;

                    Notification.Show(
                        Globals.ThisAddIn.GetResourceString("Title_Success"),
                        Globals.ThisAddIn.GetResourceString("Msg_WatermarkRemoved"));
                }
                else
                {
                    Logger.LogException("OnRemoveWatermark", Globals.ThisAddIn.GetResourceString("Error_WatermarkNotFound"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_WatermarkNotFound"));
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnRemoveWatermark", Globals.ThisAddIn.GetResourceString("Error_RemoveWatermarkFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_RemoveWatermarkFailed"));
            }
        }

        /// <summary>
        /// 导出PDF按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private async void OnExportPDF(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Documents.Count == 0)
                {
                    Logger.LogException("OnExportPDF", Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    return;
                }

                var doc = app.ActiveDocument;
                string filePath = null;
                
                using (var saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "PDF 文件 (*.pdf)|*.pdf";
                    saveDialog.FileName = Path.GetFileNameWithoutExtension(doc.Name) + ".pdf";
                    saveDialog.DefaultExt = "pdf";
                    saveDialog.Title = Globals.ThisAddIn.GetResourceString("Title_ExportPDF");

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        filePath = saveDialog.FileName;
                        // 显示导出进度通知
                        Notification.CreateProgress(
                            Globals.ThisAddIn.GetResourceString("Title_ExportPDF"),
                            "export_pdf_progress");
                        Notification.UpdateProgress(0.3, 
                            Globals.ThisAddIn.GetResourceString("Msg_ExportingPDF"));

                        // 异步执行PDF导出
                        await Task.Run(() => 
                        {
                            doc.SaveAs2(filePath,
                                FileFormat: Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                        });

                        Notification.UpdateProgress(0.8, 
                            Globals.ThisAddIn.GetResourceString("Msg_OpeningPDF"));

                        // 异步打开PDF文件
                        await Task.Run(() => Process.Start(filePath));

                        Notification.UpdateProgress(1.0, 
                            Globals.ThisAddIn.GetResourceString("Msg_ExportSuccess"));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnExportPDF", Globals.ThisAddIn.GetResourceString("Error_ExportPDFFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_ExportPDFFailed"));
            }
            finally
            {
                // 标记进度完成，通知会自动超时消失
                Notification.UpdateProgress(1.0, 
                    Globals.ThisAddIn.GetResourceString("Msg_ExportSuccess"));
            }
        }

        /// <summary>
        /// 新建文档按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void OnNewDocument(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;

                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "addin-template.docx");

                if (!File.Exists(templatePath))
                {
                    Logger.LogException("OnNewDocument", string.Format(Globals.ThisAddIn.GetResourceString("Error_TemplateNotFound"),
                        templatePath));

                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        string.Format(Globals.ThisAddIn.GetResourceString("Error_TemplateNotFound"),
                        templatePath));
                    return;
                }

                // Create new document using template
                app.Documents.Add(templatePath);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnNewDocument", Globals.ThisAddIn.GetResourceString("Error_CreateDocumentFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_CreateDocumentFailed"));
            }
        }

        /// <summary>
        /// 更新域按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void OnUpdateFields(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Documents.Count == 0)
                {
                    Logger.LogException("OnUpdateFields", Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_NoDocument"));
                    return;
                }

                var doc = app.ActiveDocument;

                // 更新所有域
                foreach (Microsoft.Office.Interop.Word.Field field in doc.Fields)
                {
                    field.Update();
                }

                // 特别更新所有目录
                foreach (Microsoft.Office.Interop.Word.TableOfContents toc in doc.TablesOfContents)
                {
                    toc.Update();
                }

                Notification.Show(
                    Globals.ThisAddIn.GetResourceString("Title_Success"),
                    Globals.ThisAddIn.GetResourceString("Msg_FieldsUpdated"));
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnUpdateFields", Globals.ThisAddIn.GetResourceString("Error_UpdateFieldsFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_UpdateFieldsFailed"));
            }
        }

        /// <summary>
        /// 设为默认模板按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void OnSetDefault(object sender, RibbonControlEventArgs e)
        {
            var result = MessageBox.Show(
                Globals.ThisAddIn.GetResourceString("Msg_ConfirmTemplateReplace"),
                Globals.ThisAddIn.GetResourceString("Title_TemplateReplace"),
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);

            if (result == DialogResult.OK)
            {
                // 当用户点击确认后，将在Word退出后自动替换模板
                IsTemplateReplacePending = true;

                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Notice"),
                    Globals.ThisAddIn.GetResourceString("Msg_ReplaceTemplatePending"));
            }
        }

        /// <summary>
        /// 关于按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void OnAbout(object sender, RibbonControlEventArgs e)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var version = assembly.GetName().Version.ToString();
            var title = assembly.GetCustomAttribute<AssemblyTitleAttribute>().Title;
            var copyright = assembly.GetCustomAttribute<AssemblyCopyrightAttribute>().Copyright;

            string buildDate = File.GetLastWriteTime(assembly.Location).ToString("yyyy-MM-dd");
            string aboutText = $@"{title}
版本: v{version} (构建于 {buildDate})
作者: ikiwi
Email: ikiwicc@gmail.com
{copyright}";


            MessageBox.Show(aboutText,
               Globals.ThisAddIn.GetResourceString("Title_About"),
               MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 反馈按钮点击事件
        /// </summary>
        /// <param name="sender">事件源</param>
        /// <param name="e">事件参数</param>
        private void OnFeedback(object sender, RibbonControlEventArgs e)
        {
            string email = "ikiwicc@gmail.com";
            string subject = "Word Add-ins Tool 反馈意见 " + DateTime.Now.ToString("yyyy-MM-dd");

            // 获取系统信息
            string dotnetVersion = Environment.Version.ToString();
            string vstoVersion = typeof(Microsoft.Office.Tools.Ribbon.RibbonBase).Assembly.GetName().Version.ToString();
            string appVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // 构建邮件正文
            // 获取详细的Windows版本信息
            string osName = Environment.OSVersion.Version.Major >= 10 ?
                          (Environment.OSVersion.Version.Build >= 22000 ? "Windows 11" : "Windows 10") :
                          "Windows";
            string versionNumber = "";
            string buildNumber = Environment.OSVersion.Version.Build.ToString();

            // 从注册表获取显示版本号(如23H2)
            try
            {
                using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
                {
                    if (key != null)
                    {
                        var displayVersion = key.GetValue("DisplayVersion")?.ToString();
                        if (!string.IsNullOrEmpty(displayVersion))
                        {
                            versionNumber = displayVersion;
                        }
                        else
                        {
                            versionNumber = $"Build {buildNumber}";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                versionNumber = $"Build {buildNumber}";
                Logger.LogException(ex, "OnFeedback", "Windows 版本识别错误");
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"), "Windows 版本识别错误");
            }

            // 获取Word版本信息
            string wordVersion = Globals.ThisAddIn.Application.Version;
            string wordBuild = Globals.ThisAddIn.Application.Build.ToString();

            // 根据版本号映射Word产品名称
            string wordProductName;
            try
            {
                if (wordVersion.StartsWith("16.0"))
                {
                    if (!string.IsNullOrEmpty(wordBuild))
                    {
                        int wordBuildNumber = int.Parse(wordBuild.Split('.')[2]);
                        // Office 2021 (16.0.14326.xxxxx 及以上)
                        if (wordBuildNumber >= 14326)
                        {
                            wordProductName = "Microsoft Word 2021";
                        }
                        // Office 2019 (16.0.10336.xxxxx - 16.0.14325.xxxxx)
                        else if (wordBuildNumber >= 10336)
                        {
                            wordProductName = "Microsoft Word 2019";
                        }
                        // Office 2016 (16.0.4229.xxxxx - 16.0.10335.xxxxx)
                        else if (wordBuildNumber >= 4229)
                        {
                            wordProductName = "Microsoft Word 2016";
                        }
                        else
                        {
                            wordProductName = "Microsoft Word 2016 早期版本";
                        }
                    }
                    else
                    {
                        wordProductName = "Microsoft Word 2016/2019/2021";
                    }
                }
                else if (wordVersion.StartsWith("15.0"))
                {
                    wordProductName = "Microsoft Word 2013";
                }
                else if (wordVersion.StartsWith("14.0"))
                {
                    wordProductName = "Microsoft Word 2010";
                }
                else if (wordVersion.StartsWith("12.0"))
                {
                    wordProductName = "Microsoft Word 2007";
                }
                else if (wordVersion.StartsWith("11.0"))
                {
                    wordProductName = "Microsoft Word 2003";
                }
                // 添加 Word 2000
                else if (wordVersion.StartsWith("9.0"))
                {
                    wordProductName = "Microsoft Word 2000";
                }
                // 添加 Word 97
                else if (wordVersion.StartsWith("8.0"))
                {
                    wordProductName = "Microsoft Word 97";
                }
                else
                {
                    wordProductName = $"Microsoft Word (未知版本: {wordVersion})";
                }
            }
            catch (Exception ex)
            {
                wordProductName = "Microsoft Word (版本识别错误)";
                Logger.LogException(ex, "OnFeedback", "Microsoft Word 版本识别错误");
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"), "Microsoft Word 版本识别错误");
            }

            string buildDate = File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location).ToString("yyyy-MM-dd");
            string body = $@"请在此写下您使用中遇到的问题，或者对本插件的新功能期待和改进优化意见：

----------------------------------------

系统信息：
- 插件版本: {appVersion} (构建于 {buildDate})
- 计算机名: {Environment.MachineName}
- 用户名: {Environment.UserName}
- 操作系统: {osName} {versionNumber} (内部版本 {buildNumber})
- Word版本: {wordProductName} (版本 {wordVersion} Build {wordBuild})
- .NET Framework运行时版本: {dotnetVersion}
- VSTO运行时版本: {vstoVersion}
- 反馈时间: {currentTime}";

            string logFilePath = Logger.LogFilePath;

            try
            {
                if (File.Exists(logFilePath))
                {
                    string logContent = File.ReadAllText(logFilePath);
                    body += $"\n\n----------------------------------------\n\n插件日志:\n{logContent}";
                }

                string mailtoUri = $"mailto:{email}?subject={Uri.EscapeDataString(subject)}&body={Uri.EscapeDataString(body)}";
                if (File.Exists(logFilePath))
                {
                    // 尝试直接添加附件（仅适用于某些邮件客户端）
                    mailtoUri = $"mailto:{email}?subject={Uri.EscapeDataString(subject)}&body={Uri.EscapeDataString(body)}&attach={Uri.EscapeDataString(logFilePath)}";
                }
                Process.Start(mailtoUri);
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "OnFeedback", Globals.ThisAddIn.GetResourceString("Error_OpenEmailFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_OpenEmailFailed"));
            }
        }
        private string CalculateFileHash(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            using (var sha256 = SHA256.Create())
            using (var stream = File.OpenRead(filePath))
            {
                byte[] hashBytes = sha256.ComputeHash(stream);
                return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// 查找名为"watermark"的水印文本框
        /// </summary>
        /// <returns>找到则返回Shape对象，否则返回null</returns>
        private Microsoft.Office.Interop.Word.Shape FindWatermarkShape()
        {
            var app = Globals.ThisAddIn.Application;
            if (app.Documents.Count == 0)
                return null;

            var doc = app.ActiveDocument;

            foreach (Microsoft.Office.Interop.Word.Section section in doc.Sections)
            {
                foreach (Microsoft.Office.Interop.Word.HeaderFooter hf in section.Footers)
                {
                    foreach (Microsoft.Office.Interop.Word.Shape shape in hf.Shapes)
                    {
                        if (shape.Name == "watermark")
                        {
                            return shape;
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Ribbon加载事件
        /// </summary>
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // Update watermark text from resources if available
            try
            {
                currentWatermarkText = Globals.ThisAddIn.GetResourceString("DefaultWatermarkText") ?? currentWatermarkText;
            }
            catch (Exception ex)
            {
                // Keep default value if resource access fails
                Logger.LogException(ex, "Ribbon1_Load", Globals.ThisAddIn.GetResourceString("Error_LoadWatermarkFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_LoadWatermarkFailed"));
            }

            // Initialize watermark check timer
            watermarkCheckTimer = new System.Windows.Forms.Timer
            {
                Interval = 5000 // 5 seconds
            };
            watermarkCheckTimer.Tick += WatermarkCheckTimer_Tick;
            watermarkCheckTimer.Start();

            // Check watermark existence and set button states
            try
            {
                var app = Globals.ThisAddIn.Application;
                bool hasWatermark = false;

                if (app.Documents.Count > 0)
                {
                    var doc = app.ActiveDocument;

                    // 使用辅助函数检查水印
                    hasWatermark = FindWatermarkShape() != null;
                }

                // Set button states based on watermark existence
                btnAddWatermark.Enabled = !hasWatermark;
                btnModifyWatermark.Enabled = hasWatermark;
                btnRemoveWatermark.Enabled = hasWatermark;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "Ribbon1_Load", Globals.ThisAddIn.GetResourceString("Error_CheckWatermarkFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_CheckWatermarkFailed"));
            }

            // Check if template files are identical
            try
            {
                string wordTemplatePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Microsoft", "Templates", "Normal.dotm");

                string addinTemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Normal.dotm");

                string tempPath = null;
                if (File.Exists(wordTemplatePath))
                {
                    // Create temp copy with meaningful name
                    string tempDir = Path.Combine(Path.GetTempPath(), "Word_AddIns");
                    Directory.CreateDirectory(tempDir);
                    tempPath = Path.Combine(tempDir, $"Normal_{DateTime.Now:yyyyMMddHHmmss}.dotm");
                    File.Copy(wordTemplatePath, tempPath, true);
                }
                else
                {
                    Logger.LogException("Ribbon1_Load", Globals.ThisAddIn.GetResourceString("Error_WordTemplateMissing"));
                    Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                        Globals.ThisAddIn.GetResourceString("Error_WordTemplateMissing"));
                }

                string wordTemplateHash = CalculateFileHash(tempPath);
                string addinTemplateHash = CalculateFileHash(addinTemplatePath);

                if (wordTemplateHash != null && addinTemplateHash != null &&
                    wordTemplateHash.Equals(addinTemplateHash, StringComparison.OrdinalIgnoreCase))
                {
                    btnSetDefault.Enabled = false;
                }
                else
                {
                    btnSetDefault.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "Ribbon1_Load", Globals.ThisAddIn.GetResourceString("Error_CompareTemplateFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_CompareTemplateFailed"));
            }
        }

        private void WatermarkCheckTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                bool hasWatermark = false;

                if (app.Documents.Count > 0)
                {
                    hasWatermark = FindWatermarkShape() != null;
                }

                // Update button states based on watermark existence
                btnAddWatermark.Enabled = !hasWatermark;
                btnModifyWatermark.Enabled = hasWatermark;
                btnRemoveWatermark.Enabled = hasWatermark;
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "WatermarkCheckTimer_Tick", Globals.ThisAddIn.GetResourceString("Error_WatermarkCheckFailed"));
                Notification.Show(Globals.ThisAddIn.GetResourceString("Title_Error"),
                    Globals.ThisAddIn.GetResourceString("Error_WatermarkCheckFailed"));
            }
        }
    }
}
