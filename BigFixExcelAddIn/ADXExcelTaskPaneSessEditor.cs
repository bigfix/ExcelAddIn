using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing;
using AddinExpress.MSO;
using QWhale.Common;
using QWhale.Syntax;
using QWhale.Syntax.CodeCompletion;
using QWhale.Editor;
using QWhale.Editor.TextSource;
using Microsoft.Win32;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using System.Collections;
using Ader.Text;

namespace BigFixExcelConnector
{
    /// <summary>
    /// Summary description for ADXExcelTaskPaneSessEditor.
    /// </summary>
    public class ADXExcelTaskPaneSessEditor : AddinExpress.XL.ADXExcelTaskPane
    {
        // RelevanceService bes = new RelevanceService();
        RelevanceBindingEx bes = new RelevanceBindingEx();

        String webReportsURL = "";
        String userName = "";
        String password = "";
        String comboBoxTaskPaneLocation = "Top";

        int maxColumnWidth = 80;
        String[] oneRow;

        private TabControl tabControl1;
        private QWhale.Syntax.Parser parserBES;
        private QWhale.Editor.TextSource.TextSource textSourceBES;
        private ToolStripStatusLabel toolStripStatusLabelMessage;
        private ToolStripStatusLabel toolStripStatusLabelEvalTime;
        private ToolStripStatusLabel toolStripStatusLabelConnectedUser;
        private ToolStripDropDownButton toolStripDropDownButtonRun;
        private TableLayoutPanel tableLayoutPanel1;
        private StatusStrip statusStrip1;
        private TabPage tabPageSessEditor;
        private SyntaxEdit syntaxEditBES;
        private Parser parser1;
        private ToolStripDropDownButton toolStripDropDownButton1;
        private ToolStripMenuItem toolStripMenuItemWrap;
        private ToolStripMenuItem toolStripMenuItemSplit;
        private AddinExpress.ToolbarControls.ADXExcelControlAdapter adxExcelControlAdapter1;
        private AddinExpress.ToolbarControls.ADXExcelControlAdapter adxExcelControlAdapter2; 
        private System.ComponentModel.IContainer components = null;
 
 	    public ADXExcelTaskPaneSessEditor()
 	     {
            // This call is required by the Windows Form Designer.
            InitializeComponent();
            // TODO: Add any initialization after the InitializeComponent call

            textSourceBES.OpenBraces = ("(").ToCharArray();
            textSourceBES.ClosingBraces = (")").ToCharArray();

            syntaxEditBES.Cursor = System.Windows.Forms.Cursors.IBeam;
            syntaxEditBES.Braces.BracesOptions = BracesOptions.Highlight;
            syntaxEditBES.Braces.BackColor = Color.Orange;
            syntaxEditBES.WordWrap = true;
            syntaxEditBES.IndentOptions = IndentOptions.AutoIndent;
            syntaxEditBES.Gutter.Options = GutterOptions.PaintLinesOnGutter;
            syntaxEditBES.Gutter.Options = GutterOptions.PaintLineNumbers;
            syntaxEditBES.WordWrap = true;

            // ************************************************************************************************************
            // Have the form capture keyboard events first.
            this.KeyPreview = true;

            // Assign the event handler to the form.
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Form1_KeyDown);
            // Assign the event handler to the text box.
            // this.textBoxRelevance.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textboxRelevance_KeyDown);

        }
 
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose( bool disposing )
        {
            if( disposing )
            {
                if(components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose( disposing );
        }
 
        #region Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ADXExcelTaskPaneSessEditor));
            this.parserBES = new QWhale.Syntax.Parser();
            this.textSourceBES = new QWhale.Editor.TextSource.TextSource(this.components);
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageSessEditor = new System.Windows.Forms.TabPage();
            this.syntaxEditBES = new QWhale.Editor.SyntaxEdit(this.components);
            this.toolStripStatusLabelMessage = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripDropDownButtonRun = new System.Windows.Forms.ToolStripDropDownButton();
            this.toolStripStatusLabelEvalTime = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabelConnectedUser = new System.Windows.Forms.ToolStripStatusLabel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripDropDownButton1 = new System.Windows.Forms.ToolStripDropDownButton();
            this.toolStripMenuItemWrap = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemSplit = new System.Windows.Forms.ToolStripMenuItem();
            this.parser1 = new QWhale.Syntax.Parser();
            this.adxExcelControlAdapter1 = new AddinExpress.ToolbarControls.ADXExcelControlAdapter(this.components);
            this.adxExcelControlAdapter2 = new AddinExpress.ToolbarControls.ADXExcelControlAdapter(this.components);
            this.tabControl1.SuspendLayout();
            this.tabPageSessEditor.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // parserBES
            // 
            this.parserBES.DefaultState = 0;
            this.parserBES.XmlScheme = resources.GetString("parserBES.XmlScheme");
            // 
            // textSourceBES
            // 
            this.textSourceBES.Lexer = this.parserBES;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageSessEditor);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(726, 108);
            this.tabControl1.TabIndex = 1;
            // 
            // tabPageSessEditor
            // 
            this.tabPageSessEditor.Controls.Add(this.syntaxEditBES);
            this.tabPageSessEditor.Location = new System.Drawing.Point(4, 22);
            this.tabPageSessEditor.Name = "tabPageSessEditor";
            this.tabPageSessEditor.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageSessEditor.Size = new System.Drawing.Size(718, 82);
            this.tabPageSessEditor.TabIndex = 0;
            this.tabPageSessEditor.Text = "Session Relevance";
            this.tabPageSessEditor.UseVisualStyleBackColor = true;
            // 
            // syntaxEditBES
            // 
            this.syntaxEditBES.BackColor = System.Drawing.SystemColors.Window;
            this.syntaxEditBES.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.syntaxEditBES.Dock = System.Windows.Forms.DockStyle.Fill;
            this.syntaxEditBES.Font = new System.Drawing.Font("Courier New", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.syntaxEditBES.Lexer = this.parserBES;
            this.syntaxEditBES.Location = new System.Drawing.Point(3, 3);
            this.syntaxEditBES.Name = "syntaxEditBES";
            this.syntaxEditBES.Size = new System.Drawing.Size(712, 76);
            this.syntaxEditBES.Source = this.textSourceBES;
            this.syntaxEditBES.TabIndex = 0;
            this.syntaxEditBES.Click += new System.EventHandler(this.syntaxEditBES_Click);
            this.syntaxEditBES.Enter += new System.EventHandler(this.syntaxEditBES_Enter);
            // 
            // toolStripStatusLabelMessage
            // 
            this.toolStripStatusLabelMessage.BorderSides = ((System.Windows.Forms.ToolStripStatusLabelBorderSides)((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left | System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) 
            | System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) 
            | System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom)));
            this.toolStripStatusLabelMessage.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenInner;
            this.toolStripStatusLabelMessage.Name = "toolStripStatusLabelMessage";
            this.toolStripStatusLabelMessage.Size = new System.Drawing.Size(542, 17);
            this.toolStripStatusLabelMessage.Spring = true;
            this.toolStripStatusLabelMessage.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // toolStripDropDownButtonRun
            // 
            this.toolStripDropDownButtonRun.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButtonRun.Image")));
            this.toolStripDropDownButtonRun.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButtonRun.Name = "toolStripDropDownButtonRun";
            this.toolStripDropDownButtonRun.Size = new System.Drawing.Size(78, 20);
            this.toolStripDropDownButtonRun.Text = "Evaluate";
            this.toolStripDropDownButtonRun.ToolTipText = "Evaluate (Ctrl-Enter or Ctrl-E)";
            this.toolStripDropDownButtonRun.Click += new System.EventHandler(this.toolStripDropDownButtonRun_Click);
            // 
            // toolStripStatusLabelEvalTime
            // 
            this.toolStripStatusLabelEvalTime.BorderSides = ((System.Windows.Forms.ToolStripStatusLabelBorderSides)((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left | System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) 
            | System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) 
            | System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom)));
            this.toolStripStatusLabelEvalTime.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenInner;
            this.toolStripStatusLabelEvalTime.Name = "toolStripStatusLabelEvalTime";
            this.toolStripStatusLabelEvalTime.Padding = new System.Windows.Forms.Padding(0, 0, 5, 0);
            this.toolStripStatusLabelEvalTime.Size = new System.Drawing.Size(9, 17);
            // 
            // toolStripStatusLabelConnectedUser
            // 
            this.toolStripStatusLabelConnectedUser.BorderSides = ((System.Windows.Forms.ToolStripStatusLabelBorderSides)((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left | System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) 
            | System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) 
            | System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom)));
            this.toolStripStatusLabelConnectedUser.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenInner;
            this.toolStripStatusLabelConnectedUser.Name = "toolStripStatusLabelConnectedUser";
            this.toolStripStatusLabelConnectedUser.Padding = new System.Windows.Forms.Padding(0, 0, 5, 0);
            this.toolStripStatusLabelConnectedUser.Size = new System.Drawing.Size(9, 17);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.tabControl1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.statusStrip1, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 22F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(726, 130);
            this.tableLayoutPanel1.TabIndex = 5;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabelMessage,
            this.toolStripStatusLabelEvalTime,
            this.toolStripStatusLabelConnectedUser,
            this.toolStripDropDownButton1,
            this.toolStripDropDownButtonRun});
            this.statusStrip1.Location = new System.Drawing.Point(0, 108);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(726, 22);
            this.statusStrip1.TabIndex = 0;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripDropDownButton1
            // 
            this.toolStripDropDownButton1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemWrap,
            this.toolStripMenuItemSplit});
            this.toolStripDropDownButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButton1.Image")));
            this.toolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton1.Name = "toolStripDropDownButton1";
            this.toolStripDropDownButton1.Size = new System.Drawing.Size(73, 20);
            this.toolStripDropDownButton1.Text = "Options";
            // 
            // toolStripMenuItemWrap
            // 
            this.toolStripMenuItemWrap.Checked = true;
            this.toolStripMenuItemWrap.CheckOnClick = true;
            this.toolStripMenuItemWrap.CheckState = System.Windows.Forms.CheckState.Checked;
            this.toolStripMenuItemWrap.Image = ((System.Drawing.Image)(resources.GetObject("toolStripMenuItemWrap.Image")));
            this.toolStripMenuItemWrap.Name = "toolStripMenuItemWrap";
            this.toolStripMenuItemWrap.Size = new System.Drawing.Size(100, 22);
            this.toolStripMenuItemWrap.Text = "Wrap";
            this.toolStripMenuItemWrap.Click += new System.EventHandler(this.toolStripMenuItemWrap_Click);
            // 
            // toolStripMenuItemSplit
            // 
            this.toolStripMenuItemSplit.Checked = true;
            this.toolStripMenuItemSplit.CheckOnClick = true;
            this.toolStripMenuItemSplit.CheckState = System.Windows.Forms.CheckState.Checked;
            this.toolStripMenuItemSplit.Image = ((System.Drawing.Image)(resources.GetObject("toolStripMenuItemSplit.Image")));
            this.toolStripMenuItemSplit.Name = "toolStripMenuItemSplit";
            this.toolStripMenuItemSplit.Size = new System.Drawing.Size(100, 22);
            this.toolStripMenuItemSplit.Text = "Split";
            this.toolStripMenuItemSplit.Click += new System.EventHandler(this.toolStripMenuItemSplit_Click);
            // 
            // parser1
            // 
            this.parser1.DefaultState = 0;
            this.parser1.XmlScheme = "<?xml version=\"1.0\" encoding=\"utf-16\"?>\r\n<LexScheme xmlns:xsi=\"http://www.w3.org/" +
    "2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\">\r\n  <Versi" +
    "on>1.5</Version>\r\n</LexScheme>";
            // 
            // ADXExcelTaskPaneSessEditor
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(726, 130);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "ADXExcelTaskPaneSessEditor";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.Text = "IBM Endpoint Manager Session Relevance Editor";
            this.tabControl1.ResumeLayout(false);
            this.tabPageSessEditor.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);

        }
        #endregion

        private void DeleteCharts()
        {
            // Delete any existing charts
            try
            {
                Excel.ShapeRange xlChartsToBeDeleted = ((Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet).Shapes.get_Range(1);
                if (xlChartsToBeDeleted != null)
                    xlChartsToBeDeleted.Delete();
            }
            catch { }
        }

        private void processEvaluate()
        {
            String[] sarray = { };

            try
            {
                toolStripStatusLabelMessage.Text = "Processing...";
                statusStrip1.Update();

                // Get URL, username and password from registry
                RetrieveSettings();


                // Clear spreadsheet
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.ClearContents();
                DeleteCharts();
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.ClearFormats();

                /*
                Excel.Range range;
                range = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A5", Missing.Value);
                range = range.get_Resize(maxRows, colCount);
                range.Cells.ClearFormats();
                */

                HiResTimer hrt = new HiResTimer();
                // Remove comments
                String queryString = System.Text.RegularExpressions.Regex.Replace(syntaxEditBES.Text, @"//.*", string.Empty);
                queryString = System.Text.RegularExpressions.Regex.Replace(queryString, @"/\*[^/]*/", string.Empty);

                hrt.Start();
                bes.Timeout = Timeout.Infinite;
                sarray = bes.GetRelevanceResult(queryString, userName, password);
                hrt.Stop();

                TimeSpan t = TimeSpan.FromMilliseconds(hrt.ElapsedMicroseconds / 1000);
                // toolStripStatusLabelEvalTime.Text = "Evaluation time: " + (hrt.ElapsedMicroseconds / 1000).ToString() + " ms";
                toolStripStatusLabelEvalTime.Text = "Query time: " + t.ToString().Remove(t.ToString().Length - 4);
                toolStripStatusLabelConnectedUser.Text = "Connected as " + userName;
                toolStripStatusLabelConnectedUser.BackColor = System.Drawing.SystemColors.Control;

                Array.Sort(sarray);


                for (int i = 0; i < sarray.Length; i++)
                {
                    if (toolStripMenuItemSplit.Checked)
                    {
                        oneRow = SplitResults(sarray[i]);
                        for (int j = 0; j < oneRow.Length; j++)
                        {
                            (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[i + 2, j + 1] = oneRow[j];
                        }
                    }
                    else
                    {
                        (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[i + 2, 1] = sarray[i];
                    }
                }

                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Columns.AutoFit();

                for (int i = 0; i < oneRow.Length; i++)
                {
                    Excel.Range rangeCol;
                    rangeCol = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(excelColumnLetter(i) + "1", excelColumnLetter(i) + "1");
                    if (Convert.ToInt32(rangeCol.ColumnWidth) > maxColumnWidth)
                    {
                        rangeCol.ColumnWidth = maxColumnWidth;
                    }
                }

                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", "A1").Select();

                // Just to make the results look good, when only 1 item returned, it will say 1 item, rather than 1 items.
                if (sarray.Length == 0)
                {
                    toolStripStatusLabelMessage.Text = "Results: " + sarray.Length + " items returned";
                    toolStripStatusLabelMessage.BackColor = System.Drawing.Color.LightGreen;
                    return;
                }
                else if (sarray.Length == 1)
                {
                    toolStripStatusLabelMessage.Text = "Results: " + sarray.Length + " item returned";
                    toolStripStatusLabelMessage.BackColor = System.Drawing.SystemColors.Control;
                }
                else if (sarray.Length <= Convert.ToInt32(2000))
                {
                    toolStripStatusLabelMessage.Text = "Results: " + sarray.Length + " items returned";
                    toolStripStatusLabelMessage.BackColor = System.Drawing.SystemColors.Control;
                }
                else
                {
                    toolStripStatusLabelMessage.Text = "Results: " + sarray.Length + " items returned. " + "2000 items displayed";
                    toolStripStatusLabelMessage.BackColor = System.Drawing.SystemColors.Control;
                }

            }
            catch (Exception ex)
            {
                toolStripStatusLabelMessage.Text = "Error: " + ex.Message;
                toolStripStatusLabelMessage.BackColor = System.Drawing.Color.LightCoral;
                if ((ex.Message.ToLower().Contains("login failed")) || (ex.Message.ToLower().Contains("remote name could not be resolved"))
                    || (ex.Message.ToLower().Contains("unable to connect to the remote server"))
                    || (ex.Message.ToLower().Contains("invalid uri"))
                    || (ex.Message.ToLower().Contains("uri prefix"))
                    || (ex.Message.ToLower().Contains("connection was closed"))
                    )
                {
                    toolStripStatusLabelConnectedUser.Text = "Not connected";
                    toolStripStatusLabelConnectedUser.BackColor = System.Drawing.Color.LightCoral;
                }

                if (ex.Message.ToLower().Contains("object reference not set to an instance of an object"))
                {
                    toolStripStatusLabelMessage.Text = "Error: " + ex.Message + " - Note that this BigFix AddIn only works for BES 7.2 or above.";
                    toolStripStatusLabelMessage.BackColor = System.Drawing.Color.LightCoral;
                    toolStripStatusLabelConnectedUser.Text = "Not connected";
                    toolStripStatusLabelConnectedUser.BackColor = System.Drawing.Color.LightCoral;
                }
                toolStripStatusLabelEvalTime.Text = "";
            }

        }

        private void toolStripDropDownButtonRun_Click(object sender, EventArgs e)
        {
            processEvaluate();
        }

        private void syntaxEditBES_Click(object sender, EventArgs e)
        {
            syntaxEditBES.Cursor = System.Windows.Forms.Cursors.IBeam;
            syntaxEditBES.UpdateCaret();
            syntaxEditBES.UpdateView();
        }

        private void syntaxEditBES_Enter(object sender, EventArgs e)
        {
            syntaxEditBES.Cursor = System.Windows.Forms.Cursors.IBeam;
            syntaxEditBES.UpdateCaret();
            syntaxEditBES.UpdateView();
            // MessageBox.Show("Enter");
        }

        #region Saving and Retrieving from Registry
        private String GetSettings(String subkey, String key)
        {
            RegistryKey regKey = Registry.CurrentUser.CreateSubKey("Software\\BigFix\\" + subkey);
            String keyValue = (String)regKey.GetValue(key);
            regKey.Close();
            return keyValue;
        }

        private int GetSubKeyCount(String subkey)
        {
            RegistryKey regKey = Registry.CurrentUser.CreateSubKey("Software\\BigFix\\" + subkey);
            int count = regKey.ValueCount;
            regKey.Close();
            return count;
        }

        private void SetSettings(String subkey, String key, String value)
        {
            RegistryKey regKey = Registry.CurrentUser.CreateSubKey("Software\\BigFix\\" + subkey);
            regKey.SetValue(key, value);
            regKey.Close();
        }

        private void RetrieveSettings()
        {
            try
            {
                webReportsURL = GetSettings("ExcelConnector", "WebReportsServer");
                userName = GetSettings("ExcelConnector", "Username");
                password = Decrypt(GetSettings("ExcelConnector", "Password"));
                comboBoxTaskPaneLocation = GetSettings("ExcelConnector", "TaskPaneLocation");

                if ((webReportsURL == String.Empty || userName == String.Empty || password == String.Empty))
                {
                    MessageBox.Show("Please configure login information to BigFix Web Reports first", "Login error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    bes.Url = webReportsURL.TrimEnd('/') + "/webreports";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error retrieving settings", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Safeguard password with encryption
        public static string Encrypt(string decrypted)
        {
            byte[] data = System.Text.ASCIIEncoding.ASCII.GetBytes(decrypted);
            byte[] rgbKey = System.Text.ASCIIEncoding.ASCII.GetBytes("72612512");
            byte[] rgbIV = System.Text.ASCIIEncoding.ASCII.GetBytes("25627182");

            MemoryStream memoryStream = new MemoryStream(1024);

            DESCryptoServiceProvider desCryptoServiceProvider = new
            DESCryptoServiceProvider();

            CryptoStream cryptoStream = new CryptoStream(memoryStream,
            desCryptoServiceProvider.CreateEncryptor(rgbKey, rgbIV),
            CryptoStreamMode.Write);

            cryptoStream.Write(data, 0, data.Length);

            cryptoStream.FlushFinalBlock();

            byte[] result = new byte[(int)memoryStream.Position];

            memoryStream.Position = 0;

            memoryStream.Read(result, 0, result.Length);

            cryptoStream.Close();

            return System.Convert.ToBase64String(result);
        }

        public static string Decrypt(string encrypted)
        {
            if (encrypted == null)
            {
                return "";
            }
            else
            {
                byte[] data = System.Convert.FromBase64String(encrypted);
                byte[] rgbKey = System.Text.ASCIIEncoding.ASCII.GetBytes("72612512");
                byte[] rgbIV = System.Text.ASCIIEncoding.ASCII.GetBytes("25627182");

                MemoryStream memoryStream = new MemoryStream(data.Length);

                DESCryptoServiceProvider desCryptoServiceProvider = new
                DESCryptoServiceProvider();

                CryptoStream cryptoStream = new CryptoStream(memoryStream,
                desCryptoServiceProvider.CreateDecryptor(rgbKey, rgbIV),
                CryptoStreamMode.Read);

                memoryStream.Write(data, 0, data.Length);

                memoryStream.Position = 0;

                string decrypted = new StreamReader(cryptoStream).ReadToEnd();

                cryptoStream.Close();
                return decrypted;
            }
        }
        #endregion

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.Control && e.KeyCode.ToString() == "Return") || (e.Control && e.KeyCode.ToString() == "E"))
            {
                // When the user presses both the 'Control' key and 'E' key,
                // or the 'Control' key and 'Return' key
                // execute the query
                e.SuppressKeyPress = true;

                processEvaluate();
            }
        }

        private string[] SplitResults(string orig_str)
        {
            ArrayList parsed_str = new ArrayList();
            String trimmedText = "";

            StringTokenizer3 tok = new StringTokenizer3(orig_str);
            tok.SymbolChars = new char[] { ',' };
            Token3 token;
            do
            {
                token = tok.Next();

                if (token.Kind.ToString() == "QuotedString")
                {
                    trimmedText = token.Value.TrimEnd(',');
                    if (trimmedText[0] == '(' && trimmedText[trimmedText.Length - 1] == ')')
                    {
                        trimmedText = trimmedText.Remove(0, 2);
                        trimmedText = trimmedText.Remove(trimmedText.Length - 2, 2);
                    }

                    // parsed_str.Add(token.Value);
                    parsed_str.Add(trimmedText);

                }

            } while (token.Kind != TokenKind3.EOF);

            String[] res = (String[])parsed_str.ToArray(typeof(string));
            return res;

        }

        private string excelColumnLetter(int intCol)
        {
            int intFirstLetter = ((intCol) / 26) + 64;
            int intSecondLetter = (intCol % 26) + 65;
            char letter1 = (intFirstLetter > 64) ? (char)intFirstLetter : ' ';
            return string.Concat(letter1, (char)intSecondLetter).Trim();
        }

        private void toolStripMenuItemSplit_Click(object sender, EventArgs e)
        {
            if (toolStripMenuItemSplit.Checked == true)
                toolStripMenuItemSplit.Text = "Split";
            else
                toolStripMenuItemSplit.Text = "No Split";
        }

        private void toolStripMenuItemWrap_Click(object sender, EventArgs e)
        {
            if (toolStripMenuItemWrap.Checked == true)
            {
                toolStripMenuItemWrap.Text = "Wrap";
                syntaxEditBES.WordWrap = true;
            }
            else
            {
                toolStripMenuItemWrap.Text = "No Wrap";
                syntaxEditBES.WordWrap = false;
            }
        }

    }
}

