﻿using System;
using System.Drawing;
using System.Windows.Forms;
using System.Resources;

namespace Word_AddIns
{
    /// <summary>
    /// 水印设置对话框，用于设置Word文档的水印文字和样式
    /// </summary>
    public partial class WatermarkDialog : Form
    {
        /// <summary>
        /// 获取或设置水印文字内容
        /// </summary>
        public string WatermarkText { get; set; }

        /// <summary>
        /// 获取或设置水印字体大小
        /// </summary>
        public int FontSize { get; set; }

        /// <summary>
        /// 获取或设置当前字体样式
        /// </summary>
        public FontStyle CurrentFontStyle { get; set; }

        private bool isBold = true;
        private bool isItalic = false;

        /// <summary>
        /// 构造函数，初始化水印对话框
        /// </summary>
        /// <param name="text">初始水印文字</param>
        /// <param name="fontSize">初始字体大小</param>
        /// <param name="style">初始字体样式</param>
        public WatermarkDialog(string text, int fontSize, FontStyle style)
        {
            InitializeComponent();
            // 删除传入文本中最后一个\r，并将其他的\r换成支持Windows多行显示的\r\n
            if (text.EndsWith("\r"))
            {
                text = text.Substring(0, text.Length - 1);  // 移除最后一个字符
            }
            WatermarkText = text.Replace("\r", "\r\n");
            FontSize = fontSize;
            CurrentFontStyle = style;

            // 确保字号有效
            if (FontSize <= 0 || FontSize > 99)
            {
                FontSize = 54;
            }

            txtWatermark.Text = WatermarkText;
            cboFontSize.SelectedItem = FontSize.ToString();
            isBold = style.HasFlag(FontStyle.Bold);
            isItalic = style.HasFlag(FontStyle.Italic);
            UpdatePreview();
        }

        /// <summary>
        /// 更新水印预览效果
        /// </summary>
        private void UpdatePreview()
        {
            // 确保字号有效
            int safeFontSize = FontSize > 0 ? FontSize : 54;

            FontStyle style = CurrentFontStyle;
            txtWatermark.Font = new Font(
                new ResourceManager("Word_AddIns.Properties.Resources", typeof(WatermarkDialog).Assembly)
                    .GetString("DefaultFontName"),
                safeFontSize,
                style);
            txtWatermark.Text = txtWatermark.Text;
            btnBold.BackColor = CurrentFontStyle.HasFlag(FontStyle.Bold) ? SystemColors.AppWorkspace : SystemColors.Control;
            btnItalic.BackColor = CurrentFontStyle.HasFlag(FontStyle.Italic) ? SystemColors.AppWorkspace : SystemColors.Control;
            isBold = CurrentFontStyle.HasFlag(FontStyle.Bold);
            isItalic = CurrentFontStyle.HasFlag(FontStyle.Italic);
        }

        /// <summary>
        /// 字体大小选择框改变事件处理
        /// </summary>
        private void CboFontSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (int.TryParse(cboFontSize.Text, out int size) && size >= 1 && size <= 99)
            {
                FontSize = size;
                UpdatePreview();
            }
            else
            {
                cboFontSize.Text = FontSize.ToString();
            }
        }

        private void CboFontSize_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Only allow numbers and control keys
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void CboFontSize_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!int.TryParse(cboFontSize.Text, out int size) || size < 1 || size > 99)
            {
                cboFontSize.Text = FontSize.ToString();
                e.Cancel = true;
            }
        }

        /// <summary>
        /// 粗体按钮点击事件处理
        /// </summary>
        private void BtnBold_Click(object sender, EventArgs e)
        {
            isBold = !isBold;
            btnBold.BackColor = isBold ? SystemColors.AppWorkspace : SystemColors.Control;
            CurrentFontStyle = isBold ? CurrentFontStyle | FontStyle.Bold : CurrentFontStyle & ~FontStyle.Bold;
            UpdatePreview();
        }

        /// <summary>
        /// 斜体按钮点击事件处理
        /// </summary>
        private void BtnItalic_Click(object sender, EventArgs e)
        {
            isItalic = !isItalic;
            btnItalic.BackColor = isItalic ? SystemColors.AppWorkspace : SystemColors.Control;
            CurrentFontStyle = isItalic ? CurrentFontStyle | FontStyle.Italic : CurrentFontStyle & ~FontStyle.Italic;
            UpdatePreview();
        }

        /// <summary>
        /// 确定按钮点击事件处理
        /// </summary>
        private void BtnOK_Click(object sender, EventArgs e)
        {
            WatermarkText = txtWatermark.Text;
            if (cboFontSize.SelectedItem != null)
            {
                FontSize = int.Parse(cboFontSize.SelectedItem.ToString());
            }
            CurrentFontStyle = txtWatermark.Font.Style;
            DialogResult = DialogResult.OK;
            Close();
        }

        /// <summary>
        /// 取消按钮点击事件处理
        /// </summary>
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
