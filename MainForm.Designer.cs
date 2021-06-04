namespace ExcelTobdf
{
    partial class MainForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.textBox_filename = new System.Windows.Forms.TextBox();
            this.button_open = new System.Windows.Forms.Button();
            this.listView = new DigitalPlatform.GUI.ListViewNF();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.button_output = new System.Windows.Forms.Button();
            this.textBox_output = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button_start = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox_location = new System.Windows.Forms.TextBox();
            this.textBox_bookType = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox_itemBarcodePrefix = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox_libraryCode = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.webBrowser2 = new System.Windows.Forms.WebBrowser();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel_msg = new System.Windows.Forms.ToolStripStatusLabel();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Excel文件:";
            // 
            // textBox_filename
            // 
            this.textBox_filename.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_filename.Location = new System.Drawing.Point(83, 12);
            this.textBox_filename.Name = "textBox_filename";
            this.textBox_filename.ReadOnly = true;
            this.textBox_filename.Size = new System.Drawing.Size(843, 21);
            this.textBox_filename.TabIndex = 1;
            // 
            // button_open
            // 
            this.button_open.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_open.Location = new System.Drawing.Point(932, 10);
            this.button_open.Name = "button_open";
            this.button_open.Size = new System.Drawing.Size(75, 23);
            this.button_open.TabIndex = 2;
            this.button_open.Text = "...";
            this.button_open.UseVisualStyleBackColor = true;
            this.button_open.Click += new System.EventHandler(this.button_open_Click);
            // 
            // listView
            // 
            this.listView.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.listView.FullRowSelect = true;
            this.listView.GridLines = true;
            this.listView.Location = new System.Drawing.Point(12, 39);
            this.listView.Name = "listView";
            this.listView.Size = new System.Drawing.Size(995, 212);
            this.listView.TabIndex = 3;
            this.listView.UseCompatibleStateImageBehavior = false;
            this.listView.View = System.Windows.Forms.View.Details;
            this.listView.ColumnContextMenuClicked += new DigitalPlatform.GUI.ListViewNF.ColumnContextMenuHandler(this.listView_ColumnContextMenuClicked);
            this.listView.MouseUp += new System.Windows.Forms.MouseEventHandler(this.listView_MouseUp);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "1";
            // 
            // button_output
            // 
            this.button_output.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_output.Location = new System.Drawing.Point(932, 255);
            this.button_output.Name = "button_output";
            this.button_output.Size = new System.Drawing.Size(75, 23);
            this.button_output.TabIndex = 6;
            this.button_output.Text = "...";
            this.button_output.UseVisualStyleBackColor = true;
            this.button_output.Click += new System.EventHandler(this.button_output_Click);
            // 
            // textBox_output
            // 
            this.textBox_output.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_output.Location = new System.Drawing.Point(87, 257);
            this.textBox_output.Name = "textBox_output";
            this.textBox_output.ReadOnly = true;
            this.textBox_output.Size = new System.Drawing.Size(839, 21);
            this.textBox_output.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 260);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "输出文件名:";
            // 
            // button_start
            // 
            this.button_start.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button_start.Location = new System.Drawing.Point(249, 516);
            this.button_start.Name = "button_start";
            this.button_start.Size = new System.Drawing.Size(75, 23);
            this.button_start.TabIndex = 7;
            this.button_start.Text = "开始";
            this.button_start.UseVisualStyleBackColor = true;
            this.button_start.Click += new System.EventHandler(this.button_start_Click);
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(34, 494);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 10;
            this.label4.Text = "馆藏地点：";
            // 
            // textBox_location
            // 
            this.textBox_location.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_location.Location = new System.Drawing.Point(105, 491);
            this.textBox_location.Name = "textBox_location";
            this.textBox_location.Size = new System.Drawing.Size(138, 21);
            this.textBox_location.TabIndex = 11;
            // 
            // textBox_bookType
            // 
            this.textBox_bookType.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_bookType.Location = new System.Drawing.Point(105, 518);
            this.textBox_bookType.Name = "textBox_bookType";
            this.textBox_bookType.Size = new System.Drawing.Size(138, 21);
            this.textBox_bookType.TabIndex = 13;
            this.textBox_bookType.Text = "普通图书";
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(34, 521);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 12);
            this.label5.TabIndex = 12;
            this.label5.Text = "册类型：";
            // 
            // textBox_itemBarcodePrefix
            // 
            this.textBox_itemBarcodePrefix.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_itemBarcodePrefix.Location = new System.Drawing.Point(105, 433);
            this.textBox_itemBarcodePrefix.Name = "textBox_itemBarcodePrefix";
            this.textBox_itemBarcodePrefix.Size = new System.Drawing.Size(138, 21);
            this.textBox_itemBarcodePrefix.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(10, 436);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(89, 12);
            this.label6.TabIndex = 14;
            this.label6.Text = "册条码号前缀：";
            // 
            // textBox_libraryCode
            // 
            this.textBox_libraryCode.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_libraryCode.Location = new System.Drawing.Point(105, 464);
            this.textBox_libraryCode.Name = "textBox_libraryCode";
            this.textBox_libraryCode.Size = new System.Drawing.Size(138, 21);
            this.textBox_libraryCode.TabIndex = 17;
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(34, 467);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 12);
            this.label7.TabIndex = 16;
            this.label7.Text = "分馆代码：";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox1.Controls.Add(this.webBrowser1);
            this.groupBox1.Location = new System.Drawing.Point(12, 284);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(502, 138);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "责任者与责任方式拆分情况";
            // 
            // webBrowser1
            // 
            this.webBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser1.Location = new System.Drawing.Point(3, 17);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(496, 118);
            this.webBrowser1.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.webBrowser2);
            this.groupBox2.Location = new System.Drawing.Point(522, 284);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(485, 138);
            this.groupBox2.TabIndex = 19;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "条码号重复记录";
            // 
            // webBrowser2
            // 
            this.webBrowser2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser2.Location = new System.Drawing.Point(3, 17);
            this.webBrowser2.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser2.Name = "webBrowser2";
            this.webBrowser2.Size = new System.Drawing.Size(479, 118);
            this.webBrowser2.TabIndex = 0;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel_msg});
            this.statusStrip1.Location = new System.Drawing.Point(0, 607);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1019, 22);
            this.statusStrip1.TabIndex = 20;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel_msg
            // 
            this.toolStripStatusLabel_msg.Name = "toolStripStatusLabel_msg";
            this.toolStripStatusLabel_msg.Size = new System.Drawing.Size(131, 17);
            this.toolStripStatusLabel_msg.Text = "toolStripStatusLabel1";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1019, 629);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.textBox_libraryCode);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.textBox_itemBarcodePrefix);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textBox_bookType);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textBox_location);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.button_start);
            this.Controls.Add(this.button_output);
            this.Controls.Add(this.textBox_output);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.listView);
            this.Controls.Add(this.button_open);
            this.Controls.Add(this.textBox_filename);
            this.Controls.Add(this.label1);
            this.Name = "MainForm";
            this.Text = "MainForm";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_filename;
        private System.Windows.Forms.Button button_open;
        private DigitalPlatform.GUI.ListViewNF listView;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.Button button_output;
        private System.Windows.Forms.TextBox textBox_output;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button_start;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox_location;
        private System.Windows.Forms.TextBox textBox_bookType;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox_itemBarcodePrefix;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox_libraryCode;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.WebBrowser webBrowser2;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel_msg;
    }
}

