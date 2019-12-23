namespace BigFixExcelConnector
{
    partial class FormOpen
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormOpen));
            this.dataGridViewSaveList = new System.Windows.Forms.DataGridView();
            this.QueryName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LastQueryTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.buttonOpenRun = new System.Windows.Forms.Button();
            this.buttonClose = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.labelStatus = new System.Windows.Forms.Label();
            this.buttonOpen = new System.Windows.Forms.Button();
            this.buttonDelete = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSaveList)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewSaveList
            // 
            this.dataGridViewSaveList.AllowUserToAddRows = false;
            this.dataGridViewSaveList.AllowUserToDeleteRows = false;
            this.dataGridViewSaveList.AllowUserToResizeRows = false;
            this.dataGridViewSaveList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewSaveList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewSaveList.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.QueryName,
            this.Type,
            this.LastQueryTime});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewSaveList.DefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridViewSaveList.Location = new System.Drawing.Point(21, 47);
            this.dataGridViewSaveList.MultiSelect = false;
            this.dataGridViewSaveList.Name = "dataGridViewSaveList";
            this.dataGridViewSaveList.ReadOnly = true;
            this.dataGridViewSaveList.RowHeadersVisible = false;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.LightSteelBlue;
            this.dataGridViewSaveList.RowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridViewSaveList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewSaveList.Size = new System.Drawing.Size(659, 335);
            this.dataGridViewSaveList.TabIndex = 1;
            this.dataGridViewSaveList.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewSaveList_CellDoubleClick);
            // 
            // QueryName
            // 
            this.QueryName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.QueryName.FillWeight = 120F;
            this.QueryName.HeaderText = "Name";
            this.QueryName.Name = "QueryName";
            this.QueryName.ReadOnly = true;
            // 
            // Type
            // 
            this.Type.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Type.HeaderText = "Type";
            this.Type.Name = "Type";
            this.Type.ReadOnly = true;
            // 
            // LastQueryTime
            // 
            this.LastQueryTime.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.LastQueryTime.HeaderText = "Last Query Time (Total / Relevance)";
            this.LastQueryTime.Name = "LastQueryTime";
            this.LastQueryTime.ReadOnly = true;
            // 
            // buttonOpenRun
            // 
            this.buttonOpenRun.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOpenRun.Location = new System.Drawing.Point(401, 395);
            this.buttonOpenRun.Name = "buttonOpenRun";
            this.buttonOpenRun.Size = new System.Drawing.Size(95, 23);
            this.buttonOpenRun.TabIndex = 2;
            this.buttonOpenRun.Tag = "";
            this.buttonOpenRun.Text = "Open and Run";
            this.buttonOpenRun.UseVisualStyleBackColor = true;
            this.buttonOpenRun.Click += new System.EventHandler(this.buttonOpenAndRun_Click);
            // 
            // buttonClose
            // 
            this.buttonClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonClose.Location = new System.Drawing.Point(626, 395);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(54, 23);
            this.buttonClose.TabIndex = 3;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Saved Queries:";
            // 
            // labelStatus
            // 
            this.labelStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelStatus.AutoSize = true;
            this.labelStatus.Location = new System.Drawing.Point(21, 400);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(0, 13);
            this.labelStatus.TabIndex = 6;
            // 
            // buttonOpen
            // 
            this.buttonOpen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOpen.Location = new System.Drawing.Point(502, 395);
            this.buttonOpen.Name = "buttonOpen";
            this.buttonOpen.Size = new System.Drawing.Size(50, 23);
            this.buttonOpen.TabIndex = 7;
            this.buttonOpen.Tag = "";
            this.buttonOpen.Text = "Open";
            this.buttonOpen.UseVisualStyleBackColor = true;
            this.buttonOpen.Click += new System.EventHandler(this.buttonOpen_Click);
            // 
            // buttonDelete
            // 
            this.buttonDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonDelete.Location = new System.Drawing.Point(558, 395);
            this.buttonDelete.Name = "buttonDelete";
            this.buttonDelete.Size = new System.Drawing.Size(62, 23);
            this.buttonDelete.TabIndex = 8;
            this.buttonDelete.Tag = "";
            this.buttonDelete.Text = "Delete";
            this.buttonDelete.UseVisualStyleBackColor = true;
            this.buttonDelete.Click += new System.EventHandler(this.buttonDelete_Click);
            // 
            // FormOpen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(699, 428);
            this.Controls.Add(this.buttonDelete);
            this.Controls.Add(this.buttonOpen);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.buttonOpenRun);
            this.Controls.Add(this.dataGridViewSaveList);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(604, 466);
            this.Name = "FormOpen";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Open Query Definition";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSaveList)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewSaveList;
        private System.Windows.Forms.Button buttonOpenRun;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.Button buttonOpen;
        private System.Windows.Forms.Button buttonDelete;
        private System.Windows.Forms.DataGridViewTextBoxColumn QueryName;
        private System.Windows.Forms.DataGridViewTextBoxColumn Type;
        private System.Windows.Forms.DataGridViewTextBoxColumn LastQueryTime;
    }
}