namespace ImportXlsToDataTable
{
    partial class Form1
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
            this.dataSet1 = new System.Data.DataSet();
            this.richTextBoxMain = new System.Windows.Forms.RichTextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.buttonImpDevList = new System.Windows.Forms.Button();
            this.buttonImpMode = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.btnImport = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label10 = new System.Windows.Forms.Label();
            this.comboBox_syslist = new System.Windows.Forms.ComboBox();
            this.buttonCheckModeTable = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.comboBox_stationlist = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.buttonCreateModePanel = new System.Windows.Forms.Button();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.label3 = new System.Windows.Forms.Label();
            this.buttonScreenClear = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonScreenReflash = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "NewDataSet";
            // 
            // richTextBoxMain
            // 
            this.richTextBoxMain.BackColor = System.Drawing.SystemColors.WindowText;
            this.richTextBoxMain.ForeColor = System.Drawing.SystemColors.Window;
            this.richTextBoxMain.Location = new System.Drawing.Point(238, 10);
            this.richTextBoxMain.Name = "richTextBoxMain";
            this.richTextBoxMain.Size = new System.Drawing.Size(414, 342);
            this.richTextBoxMain.TabIndex = 2;
            this.richTextBoxMain.Text = "";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(10, 10);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(223, 342);
            this.tabControl1.TabIndex = 3;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.button3);
            this.tabPage1.Controls.Add(this.button2);
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.label8);
            this.tabPage1.Controls.Add(this.buttonImpDevList);
            this.tabPage1.Controls.Add(this.buttonImpMode);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.label7);
            this.tabPage1.Controls.Add(this.btnImport);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(215, 316);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "数据导入";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(85, 219);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(86, 23);
            this.button3.TabIndex = 15;
            this.button3.Text = "清空清单配置";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(85, 132);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(86, 23);
            this.button2.TabIndex = 14;
            this.button2.Text = "清空类配置";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(85, 56);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(86, 23);
            this.button1.TabIndex = 13;
            this.button1.Text = "清空模式配置";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 195);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(77, 12);
            this.label8.TabIndex = 12;
            this.label8.Text = "导入设备清单";
            // 
            // buttonImpDevList
            // 
            this.buttonImpDevList.Location = new System.Drawing.Point(85, 190);
            this.buttonImpDevList.Name = "buttonImpDevList";
            this.buttonImpDevList.Size = new System.Drawing.Size(86, 23);
            this.buttonImpDevList.TabIndex = 11;
            this.buttonImpDevList.Text = "设备清单";
            this.buttonImpDevList.UseVisualStyleBackColor = true;
            this.buttonImpDevList.Click += new System.EventHandler(this.buttonImpDevList_Click);
            // 
            // buttonImpMode
            // 
            this.buttonImpMode.Location = new System.Drawing.Point(85, 26);
            this.buttonImpMode.Name = "buttonImpMode";
            this.buttonImpMode.Size = new System.Drawing.Size(86, 23);
            this.buttonImpMode.TabIndex = 10;
            this.buttonImpMode.Text = "模式信息";
            this.buttonImpMode.UseVisualStyleBackColor = true;
            this.buttonImpMode.Click += new System.EventHandler(this.buttonImpMode_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 108);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 9;
            this.label1.Text = "导入设备类表";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 32);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(65, 12);
            this.label7.TabIndex = 8;
            this.label7.Text = "导入模式表";
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(85, 102);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(86, 23);
            this.btnImport.TabIndex = 6;
            this.btnImport.Text = "设备类表";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.label10);
            this.tabPage2.Controls.Add(this.comboBox_syslist);
            this.tabPage2.Controls.Add(this.buttonCheckModeTable);
            this.tabPage2.Controls.Add(this.label9);
            this.tabPage2.Controls.Add(this.comboBox_stationlist);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.label5);
            this.tabPage2.Controls.Add(this.buttonCreateModePanel);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(215, 316);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "工具栏";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(13, 161);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(53, 12);
            this.label10.TabIndex = 16;
            this.label10.Text = "选择系统";
            // 
            // comboBox_syslist
            // 
            this.comboBox_syslist.FormattingEnabled = true;
            this.comboBox_syslist.Items.AddRange(new object[] {
            "空"});
            this.comboBox_syslist.Location = new System.Drawing.Point(98, 158);
            this.comboBox_syslist.Margin = new System.Windows.Forms.Padding(2);
            this.comboBox_syslist.Name = "comboBox_syslist";
            this.comboBox_syslist.Size = new System.Drawing.Size(92, 20);
            this.comboBox_syslist.TabIndex = 15;
            // 
            // buttonCheckModeTable
            // 
            this.buttonCheckModeTable.Location = new System.Drawing.Point(98, 64);
            this.buttonCheckModeTable.Name = "buttonCheckModeTable";
            this.buttonCheckModeTable.Size = new System.Drawing.Size(91, 23);
            this.buttonCheckModeTable.TabIndex = 14;
            this.buttonCheckModeTable.Text = "校对模式表";
            this.buttonCheckModeTable.UseVisualStyleBackColor = true;
            this.buttonCheckModeTable.Click += new System.EventHandler(this.buttonCheckModeTable_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(13, 70);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(65, 12);
            this.label9.TabIndex = 13;
            this.label9.Text = "模式表校对";
            // 
            // comboBox_stationlist
            // 
            this.comboBox_stationlist.FormattingEnabled = true;
            this.comboBox_stationlist.Items.AddRange(new object[] {
            "空"});
            this.comboBox_stationlist.Location = new System.Drawing.Point(98, 113);
            this.comboBox_stationlist.Margin = new System.Windows.Forms.Padding(2);
            this.comboBox_stationlist.Name = "comboBox_stationlist";
            this.comboBox_stationlist.Size = new System.Drawing.Size(92, 20);
            this.comboBox_stationlist.TabIndex = 12;
            this.comboBox_stationlist.SelectedIndexChanged += new System.EventHandler(this.comboBox_stationlist_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 114);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 12);
            this.label4.TabIndex = 11;
            this.label4.Text = "选择站";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(13, 29);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(77, 12);
            this.label5.TabIndex = 9;
            this.label5.Text = "模式对比画面";
            // 
            // buttonCreateModePanel
            // 
            this.buttonCreateModePanel.Location = new System.Drawing.Point(98, 24);
            this.buttonCreateModePanel.Name = "buttonCreateModePanel";
            this.buttonCreateModePanel.Size = new System.Drawing.Size(91, 23);
            this.buttonCreateModePanel.TabIndex = 8;
            this.buttonCreateModePanel.Text = "生成画面";
            this.buttonCreateModePanel.UseVisualStyleBackColor = true;
            this.buttonCreateModePanel.Click += new System.EventHandler(this.buttonCreateModePanel_Click);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.label3);
            this.tabPage3.Controls.Add(this.buttonScreenClear);
            this.tabPage3.Controls.Add(this.label2);
            this.tabPage3.Controls.Add(this.buttonScreenReflash);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Margin = new System.Windows.Forms.Padding(2);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(215, 316);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "其他";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(18, 86);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 11;
            this.label3.Text = "日志显示";
            // 
            // buttonScreenClear
            // 
            this.buttonScreenClear.Location = new System.Drawing.Point(77, 81);
            this.buttonScreenClear.Name = "buttonScreenClear";
            this.buttonScreenClear.Size = new System.Drawing.Size(95, 23);
            this.buttonScreenClear.TabIndex = 10;
            this.buttonScreenClear.Text = "清空显示";
            this.buttonScreenClear.UseVisualStyleBackColor = true;
            this.buttonScreenClear.Click += new System.EventHandler(this.buttonScreenClear_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(18, 43);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 9;
            this.label2.Text = "日志显示";
            // 
            // buttonScreenReflash
            // 
            this.buttonScreenReflash.Location = new System.Drawing.Point(77, 38);
            this.buttonScreenReflash.Name = "buttonScreenReflash";
            this.buttonScreenReflash.Size = new System.Drawing.Size(95, 23);
            this.buttonScreenReflash.TabIndex = 8;
            this.buttonScreenReflash.Text = "停止刷新";
            this.buttonScreenReflash.UseVisualStyleBackColor = true;
            this.buttonScreenReflash.Click += new System.EventHandler(this.buttonScreenReflash_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(662, 362);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.richTextBoxMain);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "模式控制界面生成系统";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Data.DataSet dataSet1;
        private System.Windows.Forms.RichTextBox richTextBoxMain;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button buttonCreateModePanel;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button buttonScreenClear;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonScreenReflash;
        private System.Windows.Forms.Button buttonImpMode;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button buttonImpDevList;
        private System.Windows.Forms.ComboBox comboBox_stationlist;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button buttonCheckModeTable;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox comboBox_syslist;
    }
}

