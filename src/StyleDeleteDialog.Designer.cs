namespace Word_AddIns
{
    partial class StyleDeleteDialog
    {
        private System.Windows.Forms.CheckedListBox stylesListBox;
        private System.Windows.Forms.Panel mainPanel;
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.stylesListBox = new System.Windows.Forms.CheckedListBox();
            this.mainPanel = new System.Windows.Forms.Panel();
            this.countLabel = new System.Windows.Forms.Label();
            this.selectAllButton = new System.Windows.Forms.Button();
            this.unselectAllButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.okButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.mainPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // stylesListBox
            // 
            this.stylesListBox.BackColor = System.Drawing.SystemColors.Control;
            this.stylesListBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.stylesListBox.CheckOnClick = true;
            this.stylesListBox.Location = new System.Drawing.Point(14, 76);
            this.stylesListBox.Name = "stylesListBox";
            this.stylesListBox.Size = new System.Drawing.Size(583, 364);
            this.stylesListBox.TabIndex = 0;
            // 
            // mainPanel
            // 
            this.mainPanel.Controls.Add(this.label1);
            this.mainPanel.Controls.Add(this.countLabel);
            this.mainPanel.Controls.Add(this.selectAllButton);
            this.mainPanel.Controls.Add(this.unselectAllButton);
            this.mainPanel.Controls.Add(this.cancelButton);
            this.mainPanel.Controls.Add(this.okButton);
            this.mainPanel.Controls.Add(this.stylesListBox);
            this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainPanel.Location = new System.Drawing.Point(0, 0);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Size = new System.Drawing.Size(609, 566);
            this.mainPanel.TabIndex = 0;
            // 
            // countLabel
            // 
            this.countLabel.AutoSize = true;
            this.countLabel.Location = new System.Drawing.Point(10, 467);
            this.countLabel.Name = "countLabel";
            this.countLabel.Size = new System.Drawing.Size(103, 24);
            this.countLabel.TabIndex = 9;
            this.countLabel.Text = "已选中: 0/0";
            // 
            // selectAllButton
            // 
            this.selectAllButton.Location = new System.Drawing.Point(12, 504);
            this.selectAllButton.Name = "selectAllButton";
            this.selectAllButton.Size = new System.Drawing.Size(93, 50);
            this.selectAllButton.TabIndex = 5;
            this.selectAllButton.Text = "全选";
            // 
            // unselectAllButton
            // 
            this.unselectAllButton.Location = new System.Drawing.Point(111, 504);
            this.unselectAllButton.Name = "unselectAllButton";
            this.unselectAllButton.Size = new System.Drawing.Size(93, 50);
            this.unselectAllButton.TabIndex = 6;
            this.unselectAllButton.Text = "全不选";
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(477, 504);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(120, 50);
            this.cancelButton.TabIndex = 7;
            this.cancelButton.Text = "取消";
            // 
            // okButton
            // 
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.okButton.Location = new System.Drawing.Point(351, 504);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(120, 50);
            this.okButton.TabIndex = 8;
            this.okButton.Text = "确认删除";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(532, 24);
            this.label1.TabIndex = 10;
            this.label1.Text = "注意：样式可能包括文档中正在使用的样式，会影响已有内容显示";
            // 
            // StyleDeleteDialog
            // 
            this.ClientSize = new System.Drawing.Size(609, 566);
            this.Controls.Add(this.mainPanel);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "StyleDeleteDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "选择需要删除的样式";
            this.mainPanel.ResumeLayout(false);
            this.mainPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button selectAllButton;
        private System.Windows.Forms.Button unselectAllButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Label countLabel;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Label label1;
    }
}
