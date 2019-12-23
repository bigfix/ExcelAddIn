namespace BigFixExcelConnector
{
    partial class ShowRelevanceCode
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ShowRelevanceCode));
            this.syntaxEditBES = new QWhale.Editor.SyntaxEdit(this.components);
            this.parserBES = new QWhale.Syntax.Parser();
            this.buttonCopy = new System.Windows.Forms.Button();
            this.imageListShowCodeWindow = new System.Windows.Forms.ImageList(this.components);
            this.buttonClose = new System.Windows.Forms.Button();
            this.buttonFlatten = new System.Windows.Forms.Button();
            this.buttonIndent = new System.Windows.Forms.Button();
            this.buttonIndentLW = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // syntaxEditBES
            // 
            this.syntaxEditBES.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.syntaxEditBES.BackColor = System.Drawing.SystemColors.Window;
            this.syntaxEditBES.Braces.BackColor = System.Drawing.Color.OrangeRed;
            this.syntaxEditBES.Braces.BracesOptions = QWhale.Editor.TextSource.BracesOptions.Highlight;
            this.syntaxEditBES.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.syntaxEditBES.Font = new System.Drawing.Font("Courier New", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.syntaxEditBES.Gutter.Options = ((QWhale.Editor.GutterOptions)(((QWhale.Editor.GutterOptions.PaintLineNumbers | QWhale.Editor.GutterOptions.PaintBookMarks) 
            | QWhale.Editor.GutterOptions.PaintLineModificators)));
            this.syntaxEditBES.Lexer = this.parserBES;
            this.syntaxEditBES.Location = new System.Drawing.Point(0, 0);
            this.syntaxEditBES.Name = "syntaxEditBES";
            this.syntaxEditBES.Selection.Options = ((QWhale.Editor.SelectionOptions)(((((QWhale.Editor.SelectionOptions.OverwriteBlocks | QWhale.Editor.SelectionOptions.SmartFormat) 
            | QWhale.Editor.SelectionOptions.RtfClipboard) 
            | QWhale.Editor.SelectionOptions.ClearOnDrag) 
            | QWhale.Editor.SelectionOptions.CopyLineWhenEmpty)));
            this.syntaxEditBES.Size = new System.Drawing.Size(881, 369);
            this.syntaxEditBES.TabIndex = 1;
            this.syntaxEditBES.Text = "";
            this.syntaxEditBES.WordWrap = true;
            // 
            // parserBES
            // 
            this.parserBES.DefaultState = 0;
            this.parserBES.XmlScheme = resources.GetString("parserBES.XmlScheme");
            // 
            // buttonCopy
            // 
            this.buttonCopy.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.buttonCopy.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonCopy.ImageKey = "clipboard-copy-document-text.png";
            this.buttonCopy.ImageList = this.imageListShowCodeWindow;
            this.buttonCopy.Location = new System.Drawing.Point(604, 379);
            this.buttonCopy.Name = "buttonCopy";
            this.buttonCopy.Size = new System.Drawing.Size(122, 23);
            this.buttonCopy.TabIndex = 2;
            this.buttonCopy.Text = "Copy to Clipboard";
            this.buttonCopy.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonCopy.UseVisualStyleBackColor = true;
            this.buttonCopy.Click += new System.EventHandler(this.buttonCopy_Click);
            // 
            // imageListShowCodeWindow
            // 
            this.imageListShowCodeWindow.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListShowCodeWindow.ImageStream")));
            this.imageListShowCodeWindow.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListShowCodeWindow.Images.SetKeyName(0, "cross.png");
            this.imageListShowCodeWindow.Images.SetKeyName(1, "clipboard-copy-document-text.png");
            this.imageListShowCodeWindow.Images.SetKeyName(2, "application_side_contract.png");
            this.imageListShowCodeWindow.Images.SetKeyName(3, "application_side_expand.png");
            // 
            // buttonClose
            // 
            this.buttonClose.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.buttonClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonClose.ImageKey = "cross.png";
            this.buttonClose.ImageList = this.imageListShowCodeWindow;
            this.buttonClose.Location = new System.Drawing.Point(744, 379);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(62, 23);
            this.buttonClose.TabIndex = 3;
            this.buttonClose.Text = "Close";
            this.buttonClose.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // buttonFlatten
            // 
            this.buttonFlatten.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.buttonFlatten.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonFlatten.ImageKey = "application_side_contract.png";
            this.buttonFlatten.ImageList = this.imageListShowCodeWindow;
            this.buttonFlatten.Location = new System.Drawing.Point(464, 379);
            this.buttonFlatten.Name = "buttonFlatten";
            this.buttonFlatten.Size = new System.Drawing.Size(122, 23);
            this.buttonFlatten.TabIndex = 4;
            this.buttonFlatten.Text = "Flatten Relevance";
            this.buttonFlatten.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonFlatten.UseVisualStyleBackColor = true;
            this.buttonFlatten.Click += new System.EventHandler(this.buttonFlatten_Click);
            // 
            // buttonIndent
            // 
            this.buttonIndent.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.buttonIndent.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonIndent.ImageKey = "application_side_expand.png";
            this.buttonIndent.ImageList = this.imageListShowCodeWindow;
            this.buttonIndent.Location = new System.Drawing.Point(240, 379);
            this.buttonIndent.Name = "buttonIndent";
            this.buttonIndent.Size = new System.Drawing.Size(206, 23);
            this.buttonIndent.TabIndex = 5;
            this.buttonIndent.Text = "Indent Relevance (Standard BigFix)";
            this.buttonIndent.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonIndent.UseVisualStyleBackColor = true;
            this.buttonIndent.Click += new System.EventHandler(this.buttonIndent_Click);
            // 
            // buttonIndentLW
            // 
            this.buttonIndentLW.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.buttonIndentLW.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonIndentLW.ImageKey = "application_side_expand.png";
            this.buttonIndentLW.ImageList = this.imageListShowCodeWindow;
            this.buttonIndentLW.Location = new System.Drawing.Point(80, 379);
            this.buttonIndentLW.Name = "buttonIndentLW";
            this.buttonIndentLW.Size = new System.Drawing.Size(142, 23);
            this.buttonIndentLW.TabIndex = 6;
            this.buttonIndentLW.Text = "Indent Relevance (LW)";
            this.buttonIndentLW.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonIndentLW.UseVisualStyleBackColor = true;
            this.buttonIndentLW.Click += new System.EventHandler(this.buttonIndentLW_Click);
            // 
            // ShowRelevanceCode
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(881, 410);
            this.Controls.Add(this.buttonIndentLW);
            this.Controls.Add(this.buttonIndent);
            this.Controls.Add(this.buttonFlatten);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.buttonCopy);
            this.Controls.Add(this.syntaxEditBES);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(845, 200);
            this.Name = "ShowRelevanceCode";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Show Relevance Code";
            this.ResumeLayout(false);

        }

        #endregion

        private QWhale.Editor.SyntaxEdit syntaxEditBES;
        private System.Windows.Forms.Button buttonCopy;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Button buttonFlatten;
        private System.Windows.Forms.Button buttonIndent;
        private System.Windows.Forms.ImageList imageListShowCodeWindow;
        private QWhale.Syntax.Parser parserBES;
        private System.Windows.Forms.Button buttonIndentLW;

    }
}