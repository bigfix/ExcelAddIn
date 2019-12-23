namespace BigFixExcelConnector
{
    partial class Fixlets
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
            this.a1Panel1 = new Owf.Controls.A1Panel();
            this.labelFixletsSelected = new System.Windows.Forms.Label();
            this.buttonUnselectAll = new System.Windows.Forms.Button();
            this.buttonSelectAll = new System.Windows.Forms.Button();
            this.checkedListBoxSites = new System.Windows.Forms.CheckedListBox();
            this.buttonGetSites = new System.Windows.Forms.Button();
            this.a1Panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // a1Panel1
            // 
            this.a1Panel1.BorderColor = System.Drawing.Color.Gray;
            this.a1Panel1.Controls.Add(this.labelFixletsSelected);
            this.a1Panel1.Controls.Add(this.buttonUnselectAll);
            this.a1Panel1.Controls.Add(this.buttonSelectAll);
            this.a1Panel1.Controls.Add(this.checkedListBoxSites);
            this.a1Panel1.Controls.Add(this.buttonGetSites);
            this.a1Panel1.GradientEndColor = System.Drawing.Color.Silver;
            this.a1Panel1.GradientStartColor = System.Drawing.Color.SlateGray;
            this.a1Panel1.Image = null;
            this.a1Panel1.ImageLocation = new System.Drawing.Point(4, 4);
            this.a1Panel1.Location = new System.Drawing.Point(3, 2);
            this.a1Panel1.Name = "a1Panel1";
            this.a1Panel1.Size = new System.Drawing.Size(401, 364);
            this.a1Panel1.TabIndex = 0;
            // 
            // labelFixletsSelected
            // 
            this.labelFixletsSelected.AutoSize = true;
            this.labelFixletsSelected.BackColor = System.Drawing.Color.Transparent;
            this.labelFixletsSelected.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelFixletsSelected.Location = new System.Drawing.Point(252, 333);
            this.labelFixletsSelected.Name = "labelFixletsSelected";
            this.labelFixletsSelected.Size = new System.Drawing.Size(98, 14);
            this.labelFixletsSelected.TabIndex = 4;
            this.labelFixletsSelected.Text = "Fixlets selected: ";
            // 
            // buttonUnselectAll
            // 
            this.buttonUnselectAll.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonUnselectAll.Location = new System.Drawing.Point(171, 328);
            this.buttonUnselectAll.Name = "buttonUnselectAll";
            this.buttonUnselectAll.Size = new System.Drawing.Size(75, 23);
            this.buttonUnselectAll.TabIndex = 3;
            this.buttonUnselectAll.Text = "Unselect All";
            this.buttonUnselectAll.UseVisualStyleBackColor = true;
            this.buttonUnselectAll.Click += new System.EventHandler(this.buttonUnselectAll_Click);
            // 
            // buttonSelectAll
            // 
            this.buttonSelectAll.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonSelectAll.Location = new System.Drawing.Point(90, 328);
            this.buttonSelectAll.Name = "buttonSelectAll";
            this.buttonSelectAll.Size = new System.Drawing.Size(75, 23);
            this.buttonSelectAll.TabIndex = 2;
            this.buttonSelectAll.Text = "Select All";
            this.buttonSelectAll.UseVisualStyleBackColor = true;
            this.buttonSelectAll.Click += new System.EventHandler(this.buttonSelectAll_Click);
            // 
            // checkedListBoxSites
            // 
            this.checkedListBoxSites.BackColor = System.Drawing.SystemColors.Window;
            this.checkedListBoxSites.CheckOnClick = true;
            this.checkedListBoxSites.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkedListBoxSites.ForeColor = System.Drawing.SystemColors.WindowText;
            this.checkedListBoxSites.FormattingEnabled = true;
            this.checkedListBoxSites.HorizontalScrollbar = true;
            this.checkedListBoxSites.Location = new System.Drawing.Point(9, 10);
            this.checkedListBoxSites.Name = "checkedListBoxSites";
            this.checkedListBoxSites.Size = new System.Drawing.Size(378, 292);
            this.checkedListBoxSites.TabIndex = 0;
            this.checkedListBoxSites.ThreeDCheckBoxes = true;
            this.checkedListBoxSites.SelectedValueChanged += new System.EventHandler(this.checkedListBoxSites_SelectedValueChanged);
            // 
            // buttonGetSites
            // 
            this.buttonGetSites.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonGetSites.Location = new System.Drawing.Point(9, 328);
            this.buttonGetSites.Name = "buttonGetSites";
            this.buttonGetSites.Size = new System.Drawing.Size(75, 23);
            this.buttonGetSites.TabIndex = 1;
            this.buttonGetSites.Text = "Get Sites";
            this.buttonGetSites.UseVisualStyleBackColor = true;
            this.buttonGetSites.Click += new System.EventHandler(this.buttonGetSites_Click);
            // 
            // Wizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(817, 453);
            this.Controls.Add(this.a1Panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Wizard";
            this.Text = "Wizard";
            this.a1Panel1.ResumeLayout(false);
            this.a1Panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Owf.Controls.A1Panel a1Panel1;
        private System.Windows.Forms.CheckedListBox checkedListBoxSites;
        private System.Windows.Forms.Button buttonGetSites;
        private System.Windows.Forms.Button buttonSelectAll;
        private System.Windows.Forms.Button buttonUnselectAll;
        private System.Windows.Forms.Label labelFixletsSelected;
    }
}