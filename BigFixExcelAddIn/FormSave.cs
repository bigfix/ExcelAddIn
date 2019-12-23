using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace BigFixExcelConnector
{
    public partial class FormSave : Form
    {
        public FormSave()
        {
            InitializeComponent();
            LoadSavedQueries();
        }

        private void LoadSavedQueries()
        {
            try
            {
                Registry.CurrentUser.CreateSubKey("Software\\BigFix\\ExcelConnector\\SavedQueries");
                RegistryKey savedKeys = Registry.CurrentUser.OpenSubKey("Software\\BigFix\\ExcelConnector\\SavedQueries");

                dataGridViewSaveList.Rows.Clear();

                foreach (String keyname in savedKeys.GetSubKeyNames())
                {
                    String[] row = new String[] { GetSettings("ExcelConnector\\SavedQueries\\" + keyname, "Name"),
                                            GetSettings("ExcelConnector\\SavedQueries\\" + keyname, "A2"),
                                            GetSettings("ExcelConnector\\SavedQueries\\" + keyname, "A16")};
                    dataGridViewSaveList.Rows.Add(row);
                }

                if (dataGridViewSaveList.Rows.Count > 0)
                {
                    dataGridViewSaveList.Rows[0].Cells[0].Selected = false;
                    dataGridViewSaveList.Sort(dataGridViewSaveList.Columns[0], System.ComponentModel.ListSortDirection.Ascending);
                }
                textBoxQueryName.Focus();
                dataGridViewSaveList.Update();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error retrieving saved definitions", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private String CheckDuplicateName()
        {
            try
            {
                RegistryKey savedKeys = Registry.CurrentUser.OpenSubKey("Software\\BigFix\\ExcelConnector\\SavedQueries");

                foreach (String keyname in savedKeys.GetSubKeyNames())
                {
                    if (textBoxQueryName.Text.Trim().ToLower() == GetSettings("ExcelConnector\\SavedQueries\\" + keyname, "Name").ToLower())
                        return keyname;
                }

                return "No duplicate name";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error checking for duplicate names", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "No duplicate name";
            }

        }

        // This allows the Enter key to invoke the Save button
        private void textBoxQueryName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                buttonSave_Click(sender, e);
        }

        private void ProcessSave()
        {
            try
            {
                String regName = CheckDuplicateName();

                if (regName == "No duplicate name")
                {
                    DateTime now = DateTime.Now;
                    regName = now.ToString("yyyyMMddHHmmss");
                }
                else
                {
                    if (MessageBox.Show("There is already a query with the same name. Do you want to Replace?", "Name found", MessageBoxButtons.YesNo) == DialogResult.No)
                        return;
                }

                Excel.Range oneCell;
                Excel.Worksheet hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                // A17 is the name of the report 
                oneCell = hiddenWorksheet.get_Range("A17", "A17");
                oneCell.Value = textBoxQueryName.Text.Trim();

                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = textBoxQueryName.Text.Trim();

                // Name the sheet to be the same as the report name. Sheet name has to be less than 31 characters
                String sheetName = "";
                if (textBoxQueryName.Text.Trim().Length > 30)
                {
                    sheetName = textBoxQueryName.Text.Trim().Substring(0, 28) + "..";
                }
                else
                {
                    sheetName = textBoxQueryName.Text.Trim();
                }
                ((Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet).Name = sheetName;

                SetSettings("ExcelConnector\\SavedQueries\\" + regName, "Name", textBoxQueryName.Text.Trim());

                for (int i = 1; i <= 19; i++)
                {
                    // i == 1 is the Relevance Statement
                    // Since there are \n linefeeds that will mess up import and export, first turn them textual \n
                    /*
                    if (i == 1)
                    {
                        oneCell = hiddenWorksheet.get_Range("A" + i.ToString(), "A" + i.ToString());
                        oneCell.Value = oneCell.Value.ToString().Replace("\n", "\\n");
                    }
                    else
                    {
                        oneCell = hiddenWorksheet.get_Range("A" + i.ToString(), "A" + i.ToString());
                    }
                    */

                    oneCell = hiddenWorksheet.get_Range("A" + i.ToString(), "A" + i.ToString());

                    if (oneCell.Value == null)
                        SetSettings("ExcelConnector\\SavedQueries\\" + regName, "A" + i.ToString(), "");
                    else
                    {
                        if (i == 1)
                        {
                            SetSettings("ExcelConnector\\SavedQueries\\" + regName, "A" + i.ToString(), oneCell.Value.ToString().Replace("\n", "\\n"));
                        }
                        else
                        {
                            SetSettings("ExcelConnector\\SavedQueries\\" + regName, "A" + i.ToString(), oneCell.Value.ToString());
                        }
                    }
                }

                LoadSavedQueries();

                // Highlight the row that the user just saved
                for (int i = 0; i < dataGridViewSaveList.Rows.Count; i++)
                {
                    if (dataGridViewSaveList.Rows[i].Cells[0].Value.ToString() == textBoxQueryName.Text.Trim())
                    {
                        dataGridViewSaveList.CurrentCell = dataGridViewSaveList.Rows[i].Cells[0];
                    }
                }

                labelStatus.ForeColor = System.Drawing.Color.DarkGreen;
                labelStatus.Text = "Successfully saved";

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error saving query definition to registry", MessageBoxButtons.OK, MessageBoxIcon.Error);
                labelStatus.ForeColor = System.Drawing.Color.DarkRed;
                labelStatus.Text = ex.Message;
            }
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            ProcessSave();
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

        #endregion

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void dataGridViewSaveList_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            CopySelectedRowToTextBox(e.RowIndex);
        }

        private void CopySelectedRowToTextBox(int rowIndex)
        {
            if (dataGridViewSaveList.SelectedRows.Count != 0 || rowIndex != -1)
            {
                textBoxQueryName.Text = dataGridViewSaveList.SelectedRows[0].Cells[0].Value.ToString();
                buttonSave.Enabled = true;
            }
        }

        private void textBoxQueryName_KeyPress(object sender, KeyPressEventArgs e)
        {            
            if (e.KeyChar == '?' || 
                e.KeyChar == '\\' || 
                e.KeyChar == '/' || 
                e.KeyChar == '*' || 
                e.KeyChar == '[' || 
                e.KeyChar == ']')
            {
                SendKeys.Send("{BACKSPACE}");
                labelStatus.ForeColor = System.Drawing.Color.DarkRed;
                labelStatus.Text = "Please don't use these characters: \\ / ? * [ ]";
            } 

            if (textBoxQueryName.Text == "")
                buttonSave.Enabled = false;
            else
                buttonSave.Enabled = true;
        }

        private void dataGridViewSaveList_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            CopySelectedRowToTextBox(e.RowIndex);
            ProcessSave();
        }


    }
}
