using System.Collections.Generic;
using System.Windows.Forms;

namespace Word_AddIns
{
    /// <summary>
    /// 样式删除对话框，用于让用户选择要删除的Word样式
    /// </summary>
    public partial class StyleDeleteDialog : Form
    {
        /// <summary>
        /// 获取用户选择的样式列表
        /// </summary>
        public List<string> SelectedStyles { get; private set; }

        /// <summary>
        /// 构造函数，初始化样式删除对话框
        /// </summary>
        /// <param name="stylesToDelete">可删除的样式列表</param>
        public StyleDeleteDialog(List<string> stylesToDelete)
        {
            InitializeComponent();
            InitializeStylesList(stylesToDelete);
            UpdateCountLabel();
        }

        /// <summary>
        /// 初始化样式列表控件
        /// </summary>
        /// <param name="styles">要显示的样式列表</param>
        private void InitializeStylesList(List<string> styles)
        {
            stylesListBox.Items.Clear();
            foreach (var style in styles)
            {
                stylesListBox.Items.Add(style, true);
            }
            stylesListBox.ItemCheck += (s, e) =>
            {
                // 延迟调用以确保状态已更新
                BeginInvoke((MethodInvoker)UpdateCountLabel);
            };

            selectAllButton.Click += (s, e) =>
            {
                for (int i = 0; i < stylesListBox.Items.Count; i++)
                    stylesListBox.SetItemChecked(i, true);
                UpdateCountLabel();
            };

            unselectAllButton.Click += (s, e) =>
            {
                for (int i = 0; i < stylesListBox.Items.Count; i++)
                    stylesListBox.SetItemChecked(i, false);
                UpdateCountLabel();
            };
        }

        /// <summary>
        /// 更新已选样式数量统计标签
        /// </summary>
        private void UpdateCountLabel()
        {
            int selected = 0;
            for (int i = 0; i < stylesListBox.Items.Count; i++)
            {
                if (stylesListBox.GetItemChecked(i))
                    selected++;
            }
            int total = stylesListBox.Items.Count;
            countLabel.Text = $"已选中: {selected}/{total} 个样式";
        }

        /// <summary>
        /// 窗体关闭事件处理，保存用户选择的样式
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (this.DialogResult == DialogResult.OK)
            {
                SelectedStyles = new List<string>();
                foreach (var item in stylesListBox.CheckedItems)
                {
                    SelectedStyles.Add(item.ToString());
                }
            }
            base.OnFormClosing(e);
        }
    }
}
