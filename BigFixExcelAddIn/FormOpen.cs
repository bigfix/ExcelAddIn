using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Reflection;

namespace BigFixExcelConnector
{
    public partial class FormOpen : Form
    {
        String RunQueryAfterOpen = "close";
        Excel.XlSheetVisibility HiddenSheetVisibility = Excel.XlSheetVisibility.xlSheetHidden;

        public FormOpen()
        {
            InitializeComponent();
            LoadSavedQueries();
        }

        public String RunFormOpen()
        {
            this.ShowDialog();
            return RunQueryAfterOpen;
        }

        private void LoadSavedQueries()
        {
            RegistryKey savedKeys;
            String[] savedList;
            try
            {
                savedKeys = Registry.CurrentUser.OpenSubKey("Software\\BigFix\\ExcelConnector\\SavedQueries");
                savedList = savedKeys.GetSubKeyNames();
                if (savedList == null || savedList.Length == 0)
                {
                    labelStatus.ForeColor = System.Drawing.Color.DarkRed;
                    labelStatus.Text = "No saved query definitions found";
                    return;
                }
            }
            catch
            {
                labelStatus.ForeColor = System.Drawing.Color.DarkRed;
                labelStatus.Text = "No saved query definitions found";
                return;
            }

            try
            {
                dataGridViewSaveList.Rows.Clear();

                savedKeys = Registry.CurrentUser.OpenSubKey("Software\\BigFix\\ExcelConnector\\SavedQueries");
                savedList = savedKeys.GetSubKeyNames();

                foreach (String keyname in savedList)
                {
                    String[] row = new String[] { GetSettings("ExcelConnector\\SavedQueries\\" + keyname, "Name"),
                                                  GetSettings("ExcelConnector\\SavedQueries\\" + keyname, "A2"),
                                                  GetSettings("ExcelConnector\\SavedQueries\\" + keyname, "A16")};
                    dataGridViewSaveList.Rows.Add(row);
                }

                dataGridViewSaveList.Rows[0].Cells[0].Selected = false;
                dataGridViewSaveList.Sort(dataGridViewSaveList.Columns[0], System.ComponentModel.ListSortDirection.Ascending);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error retrieving saved definitions", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenQuery()
        {
            try
            {
                RegistryKey savedKeys = Registry.CurrentUser.OpenSubKey("Software\\BigFix\\ExcelConnector\\SavedQueries");
                String selectedName = "";

                foreach (String keyname in savedKeys.GetSubKeyNames())
                {
                    if (((String)(dataGridViewSaveList.SelectedRows[0].Cells[0].Value)).ToLower() == GetSettings("ExcelConnector\\SavedQueries\\" + keyname, "Name").ToLower())
                        selectedName = keyname;
                }

                CreateHiddenExcelWorksheet();

                Excel.Range oneCell;
                Excel.Worksheet hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");

                for (int i = 1; i <= 19; i++)
                {
                    oneCell = hiddenWorksheet.get_Range("A" + i.ToString(), "A" + i.ToString());
                    oneCell.Value = GetSettings("ExcelConnector\\SavedQueries\\" + selectedName, "A" + i.ToString());

                    if (i == 1)
                    {
                        oneCell.Value = oneCell.Value.ToString().Replace("\\n", "\n");
                    }
                }

                labelStatus.ForeColor = System.Drawing.Color.DarkGreen;
                labelStatus.Text = "Successfully opened";

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error opening the selected query", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateHiddenExcelWorksheet()
        {

            Excel.Worksheet storageWorksheet;

            try
            {
                storageWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
            }
            catch (Exception ex)
            {
                storageWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                storageWorksheet.Name = "BigFixExcelConnector";
                storageWorksheet.Visible = HiddenSheetVisibility;
                String nullMsg = ex.Message;
            }

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
            RunQueryAfterOpen = "close";

            Close();
        }

        private void buttonOpen_Click(object sender, EventArgs e)
        {
            if (dataGridViewSaveList.SelectedRows.Count == 0)
            {
                labelStatus.ForeColor = System.Drawing.Color.DarkRed;
                labelStatus.Text = "Please select a query to open";
                return;
            }

            RunQueryAfterOpen = "false";

            OpenQuery();

            Close();
        }

        private void buttonOpenAndRun_Click(object sender, EventArgs e)
        {
            if (dataGridViewSaveList.SelectedRows.Count == 0)
            {
                labelStatus.ForeColor = System.Drawing.Color.DarkRed;
                labelStatus.Text = "Please select a query to open";
                return;
            }

            RunQueryAfterOpen = "true";

            OpenQuery();

            Close();
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridViewSaveList.SelectedRows.Count == 0)
                {
                    labelStatus.ForeColor = System.Drawing.Color.DarkRed;
                    labelStatus.Text = "Please select a query to delete";
                    return;
                }

                RegistryKey savedKeys = Registry.CurrentUser.OpenSubKey("Software\\BigFix\\ExcelConnector\\SavedQueries");

                foreach (String keyname in savedKeys.GetSubKeyNames())
                {
                    if (((String)(dataGridViewSaveList.SelectedRows[0].Cells[0].Value)).ToLower() == GetSettings("ExcelConnector\\SavedQueries\\" + keyname, "Name").ToLower())
                    {
                        Registry.CurrentUser.DeleteSubKeyTree("Software\\BigFix\\ExcelConnector\\SavedQueries\\" + keyname);
                        break;
                    }
                }

                dataGridViewSaveList.Rows.Remove(dataGridViewSaveList.SelectedRows[0]);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error deleting query", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void dataGridViewSaveList_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            if (dataGridViewSaveList.SelectedRows.Count == 0 || e.RowIndex == -1)
            {
                labelStatus.ForeColor = System.Drawing.Color.DarkRed;
                labelStatus.Text = "Please double-click on a row to open";
                return;
            }

            OpenQuery();

            RunQueryAfterOpen = "true";

            Close();

        }


    }
}
