namespace ExcelReplaceTool
{
    partial class ExcelReplaceTool
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
            this.m_textPath = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.m_culName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.m_output = new System.Windows.Forms.ListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.m_textSource = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.m_textTarget = new System.Windows.Forms.TextBox();
            this.m_btnReplace = new System.Windows.Forms.Button();
            this.m_asyncWorker = new System.ComponentModel.BackgroundWorker();
            this.m_textState = new System.Windows.Forms.Label();
            this.m_notIgnoreCase = new System.Windows.Forms.CheckBox();
            this.m_caseAll = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "文件路径：";
            // 
            // m_textPath
            // 
            this.m_textPath.AllowDrop = true;
            this.m_textPath.Location = new System.Drawing.Point(93, 10);
            this.m_textPath.Name = "m_textPath";
            this.m_textPath.Size = new System.Drawing.Size(354, 21);
            this.m_textPath.TabIndex = 1;
            this.m_textPath.TextChanged += new System.EventHandler(this.m_textPath_TextChanged);
            this.m_textPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.m_textPath_DragDrop);
            this.m_textPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.m_textPath_DragEnter);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "列名：";
            // 
            // m_culName
            // 
            this.m_culName.Location = new System.Drawing.Point(93, 37);
            this.m_culName.Name = "m_culName";
            this.m_culName.Size = new System.Drawing.Size(354, 21);
            this.m_culName.TabIndex = 4;
            this.m_culName.TextChanged += new System.EventHandler(this.m_culName_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 94);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 5;
            this.label3.Text = "输出：";
            // 
            // m_output
            // 
            this.m_output.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_output.FormattingEnabled = true;
            this.m_output.ItemHeight = 12;
            this.m_output.Location = new System.Drawing.Point(24, 118);
            this.m_output.Name = "m_output";
            this.m_output.Size = new System.Drawing.Size(423, 244);
            this.m_output.TabIndex = 6;
            this.m_output.SelectedIndexChanged += new System.EventHandler(this.m_output_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(22, 68);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 7;
            this.label4.Text = "源字符串：";
            // 
            // m_textSource
            // 
            this.m_textSource.Location = new System.Drawing.Point(93, 65);
            this.m_textSource.Name = "m_textSource";
            this.m_textSource.Size = new System.Drawing.Size(126, 21);
            this.m_textSource.TabIndex = 8;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(238, 68);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(77, 12);
            this.label5.TabIndex = 9;
            this.label5.Text = "目标字符串：";
            // 
            // m_textTarget
            // 
            this.m_textTarget.Location = new System.Drawing.Point(321, 65);
            this.m_textTarget.Name = "m_textTarget";
            this.m_textTarget.Size = new System.Drawing.Size(126, 21);
            this.m_textTarget.TabIndex = 10;
            // 
            // m_btnReplace
            // 
            this.m_btnReplace.Location = new System.Drawing.Point(372, 92);
            this.m_btnReplace.Name = "m_btnReplace";
            this.m_btnReplace.Size = new System.Drawing.Size(75, 23);
            this.m_btnReplace.TabIndex = 11;
            this.m_btnReplace.Text = "替换";
            this.m_btnReplace.UseVisualStyleBackColor = true;
            this.m_btnReplace.Click += new System.EventHandler(this.m_btnReplace_Click);
            // 
            // m_asyncWorker
            // 
            this.m_asyncWorker.WorkerReportsProgress = true;
            this.m_asyncWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.m_asyncWorker_DoWork);
            this.m_asyncWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.m_asyncWorker_ProgressChanged);
            // 
            // m_textState
            // 
            this.m_textState.AutoSize = true;
            this.m_textState.Location = new System.Drawing.Point(22, 365);
            this.m_textState.Name = "m_textState";
            this.m_textState.Size = new System.Drawing.Size(0, 12);
            this.m_textState.TabIndex = 12;
            // 
            // m_notIgnoreCase
            // 
            this.m_notIgnoreCase.AutoSize = true;
            this.m_notIgnoreCase.Location = new System.Drawing.Point(93, 92);
            this.m_notIgnoreCase.Name = "m_notIgnoreCase";
            this.m_notIgnoreCase.Size = new System.Drawing.Size(84, 16);
            this.m_notIgnoreCase.TabIndex = 13;
            this.m_notIgnoreCase.Text = "区分大小写";
            this.m_notIgnoreCase.UseVisualStyleBackColor = true;
            // 
            // m_caseAll
            // 
            this.m_caseAll.AutoSize = true;
            this.m_caseAll.Location = new System.Drawing.Point(183, 92);
            this.m_caseAll.Name = "m_caseAll";
            this.m_caseAll.Size = new System.Drawing.Size(84, 16);
            this.m_caseAll.TabIndex = 14;
            this.m_caseAll.Text = "单元格匹配";
            this.m_caseAll.UseVisualStyleBackColor = true;
            // 
            // ExcelReplaceTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(471, 391);
            this.Controls.Add(this.m_caseAll);
            this.Controls.Add(this.m_notIgnoreCase);
            this.Controls.Add(this.m_textState);
            this.Controls.Add(this.m_btnReplace);
            this.Controls.Add(this.m_textTarget);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.m_textSource);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.m_output);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.m_culName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.m_textPath);
            this.Controls.Add(this.label1);
            this.MinimumSize = new System.Drawing.Size(487, 430);
            this.Name = "ExcelReplaceTool";
            this.Text = "Excel表精确替换工具";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox m_textPath;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox m_culName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListBox m_output;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox m_textSource;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox m_textTarget;
        private System.Windows.Forms.Button m_btnReplace;
        private System.ComponentModel.BackgroundWorker m_asyncWorker;
        private System.Windows.Forms.Label m_textState;
        private System.Windows.Forms.CheckBox m_notIgnoreCase;
        private System.Windows.Forms.CheckBox m_caseAll;
    }
}

