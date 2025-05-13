﻿namespace Word_AddIns
{
    partial class WatermarkDialog
    {
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
            this.txtWatermark = new System.Windows.Forms.TextBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cboFontSize = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnBold = new System.Windows.Forms.Button();
            this.btnItalic = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtWatermark
            // 
            this.txtWatermark.AcceptsReturn = true;
            this.txtWatermark.Font = new System.Drawing.Font("微软雅黑", 54F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtWatermark.Location = new System.Drawing.Point(18, 57);
            this.txtWatermark.Margin = new System.Windows.Forms.Padding(4);
            this.txtWatermark.Multiline = true;
            this.txtWatermark.Name = "txtWatermark";
            this.txtWatermark.Size = new System.Drawing.Size(1229, 450);
            this.txtWatermark.TabIndex = 0;
            this.txtWatermark.Text = global::Word_AddIns.Properties.Resources.DefaultWatermarkText;
            // 
            // btnOK
            // 
            this.btnOK.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnOK.Location = new System.Drawing.Point(993, 562);
            this.btnOK.Margin = new System.Windows.Forms.Padding(4);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(112, 34);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = global::Word_AddIns.Properties.Resources.btn_OK;
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.BtnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnCancel.Location = new System.Drawing.Point(1135, 562);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(4);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(112, 34);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = global::Word_AddIns.Properties.Resources.btn_Cancel;
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(18, 30);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(140, 24);
            this.label1.TabIndex = 3;
            this.label1.Text = global::Word_AddIns.Properties.Resources.Label_WatermarkText;
            // 
            // cboFontSize
            // 
            this.cboFontSize.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboFontSize.FormattingEnabled = true;
            this.cboFontSize.Items.AddRange(new object[] {
            "20",
            "22",
            "24",
            "26",
            "28",
            "36",
            "48",
            "54"});
            this.cboFontSize.Location = new System.Drawing.Point(22, 556);
            this.cboFontSize.Name = "cboFontSize";
            this.cboFontSize.Size = new System.Drawing.Size(122, 32);
            this.cboFontSize.TabIndex = 1;
            this.cboFontSize.Text = "54";
            this.cboFontSize.SelectedIndexChanged += new System.EventHandler(this.CboFontSize_SelectedIndexChanged);
            this.cboFontSize.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.CboFontSize_KeyPress);
            this.cboFontSize.Validating += new System.ComponentModel.CancelEventHandler(this.CboFontSize_Validating);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(22, 529);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(122, 24);
            this.label2.TabIndex = 4;
            this.label2.Text = global::Word_AddIns.Properties.Resources.Label_FontSize;
            // 
            // btnBold
            // 
            this.btnBold.FlatAppearance.BorderSize = 2;
            this.btnBold.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnBold.Location = new System.Drawing.Point(175, 540);
            this.btnBold.Margin = new System.Windows.Forms.Padding(0);
            this.btnBold.Name = "btnBold";
            this.btnBold.Size = new System.Drawing.Size(48, 48);
            this.btnBold.TabIndex = 5;
            this.btnBold.Text = "B";
            this.btnBold.UseVisualStyleBackColor = false;
            this.btnBold.Click += new System.EventHandler(this.BtnBold_Click);
            // 
            // btnItalic
            // 
            this.btnItalic.FlatAppearance.BorderSize = 2;
            this.btnItalic.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnItalic.Location = new System.Drawing.Point(229, 540);
            this.btnItalic.Margin = new System.Windows.Forms.Padding(0);
            this.btnItalic.Name = "btnItalic";
            this.btnItalic.Size = new System.Drawing.Size(48, 48);
            this.btnItalic.TabIndex = 6;
            this.btnItalic.Text = "I";
            this.btnItalic.UseVisualStyleBackColor = false;
            this.btnItalic.Click += new System.EventHandler(this.BtnItalic_Click);
            // 
            // WatermarkDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(1278, 624);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtWatermark);
            this.Controls.Add(this.cboFontSize);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnBold);
            this.Controls.Add(this.btnItalic);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "WatermarkDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = global::Word_AddIns.Properties.Resources.Title_WatermarkDialog;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtWatermark;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cboFontSize;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnBold;
        private System.Windows.Forms.Button btnItalic;
    }
}
