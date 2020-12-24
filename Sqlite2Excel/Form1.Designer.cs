namespace Sqlite2Excel
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxSourceName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxOutputName = new System.Windows.Forms.TextBox();
            this.buttonSelectSource = new System.Windows.Forms.Button();
            this.buttonSelectOutput = new System.Windows.Forms.Button();
            this.buttonExport = new System.Windows.Forms.Button();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.ToolStripMenuItemRestore = new System.Windows.Forms.ToolStripMenuItem();
            this.ToolStripMenuItemExit = new System.Windows.Forms.ToolStripMenuItem();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "源文件:";
            // 
            // textBoxSourceName
            // 
            this.textBoxSourceName.Location = new System.Drawing.Point(86, 15);
            this.textBoxSourceName.Name = "textBoxSourceName";
            this.textBoxSourceName.Size = new System.Drawing.Size(381, 21);
            this.textBoxSourceName.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(24, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "输出文件：";
            // 
            // textBoxOutputName
            // 
            this.textBoxOutputName.Enabled = false;
            this.textBoxOutputName.Location = new System.Drawing.Point(86, 42);
            this.textBoxOutputName.Name = "textBoxOutputName";
            this.textBoxOutputName.Size = new System.Drawing.Size(381, 21);
            this.textBoxOutputName.TabIndex = 3;
            // 
            // buttonSelectSource
            // 
            this.buttonSelectSource.Location = new System.Drawing.Point(486, 15);
            this.buttonSelectSource.Name = "buttonSelectSource";
            this.buttonSelectSource.Size = new System.Drawing.Size(87, 23);
            this.buttonSelectSource.TabIndex = 4;
            this.buttonSelectSource.Text = "选择源文件";
            this.buttonSelectSource.UseVisualStyleBackColor = true;
            this.buttonSelectSource.Click += new System.EventHandler(this.buttonSelectSource_Click);
            // 
            // buttonSelectOutput
            // 
            this.buttonSelectOutput.Enabled = false;
            this.buttonSelectOutput.Location = new System.Drawing.Point(486, 42);
            this.buttonSelectOutput.Name = "buttonSelectOutput";
            this.buttonSelectOutput.Size = new System.Drawing.Size(87, 23);
            this.buttonSelectOutput.TabIndex = 5;
            this.buttonSelectOutput.Text = "选择输出文件";
            this.buttonSelectOutput.UseVisualStyleBackColor = true;
            this.buttonSelectOutput.Click += new System.EventHandler(this.buttonSelectOutput_Click);
            // 
            // buttonExport
            // 
            this.buttonExport.Location = new System.Drawing.Point(486, 80);
            this.buttonExport.Name = "buttonExport";
            this.buttonExport.Size = new System.Drawing.Size(84, 23);
            this.buttonExport.TabIndex = 6;
            this.buttonExport.Text = "导出";
            this.buttonExport.UseVisualStyleBackColor = true;
            this.buttonExport.Click += new System.EventHandler(this.buttonExport_Click);
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.ContextMenuStrip = this.contextMenuStrip1;
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "Sqlite2Excel";
            this.notifyIcon1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseDoubleClick);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolStripMenuItemRestore,
            this.ToolStripMenuItemExit});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(101, 48);
            // 
            // ToolStripMenuItemRestore
            // 
            this.ToolStripMenuItemRestore.Name = "ToolStripMenuItemRestore";
            this.ToolStripMenuItemRestore.Size = new System.Drawing.Size(100, 22);
            this.ToolStripMenuItemRestore.Text = "还原";
            this.ToolStripMenuItemRestore.Click += new System.EventHandler(this.ToolStripMenuItemRestore_Click);
            // 
            // ToolStripMenuItemExit
            // 
            this.ToolStripMenuItemExit.Name = "ToolStripMenuItemExit";
            this.ToolStripMenuItemExit.Size = new System.Drawing.Size(100, 22);
            this.ToolStripMenuItemExit.Text = "退出";
            this.ToolStripMenuItemExit.Click += new System.EventHandler(this.ToolStripMenuItemExit_Click);
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(93, 80);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 21);
            this.dateTimePicker1.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 85);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 8;
            this.label3.Text = "起始时间：";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(582, 236);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.buttonExport);
            this.Controls.Add(this.buttonSelectOutput);
            this.Controls.Add(this.buttonSelectSource);
            this.Controls.Add(this.textBoxOutputName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxSourceName);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Sqite2Excel";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxSourceName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxOutputName;
        private System.Windows.Forms.Button buttonSelectSource;
        private System.Windows.Forms.Button buttonSelectOutput;
        private System.Windows.Forms.Button buttonExport;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemRestore;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemExit;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label3;
    }
}

