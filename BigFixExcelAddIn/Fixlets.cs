using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.IO;
using System.Security.Cryptography;

namespace BigFixExcelConnector
{
    public partial class Fixlets : Form
    {
        // RelevanceService bes = new RelevanceService();
        RelevanceBindingEx bes = new RelevanceBindingEx();

        String webReportsURL = "";
        String userName = "";
        String password = "";
        String[] results;

        public Fixlets()
        {
            InitializeComponent();

            // Get URL, username and password from registry
            RetrieveSettings();
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

                bes.Url = webReportsURL + "/webreports";
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

        private void buttonGetSites_Click(object sender, EventArgs e)
        {
            checkedListBoxSites.Items.Clear();
            checkedListBoxSites.Update();

            String queryString = "(name of it & \" (\" & number of fixlets of it as string & \")\") of bes sites";
            results = bes.GetRelevanceResult(queryString, userName, password);
            Array.Sort(results);
            for (int i = 0; i < results.Length; i++)
            {
                checkedListBoxSites.Items.Add(results[i]);
            }
        }

        private void buttonSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxSites.Items.Count; i++)
            {
                checkedListBoxSites.SetItemChecked(i, true);
            }
            calcFixlets();
        }

        private void buttonUnselectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxSites.Items.Count; i++)
            {
                checkedListBoxSites.SetItemChecked(i, false);
            }
            calcFixlets();
        }

        private void checkedListBoxSites_SelectedValueChanged(object sender, EventArgs e)
        {
            calcFixlets();
        }

        private void calcFixlets()
        {
            int numberOfFixletsSelected = 0;
            String checkedItem = "";

            for (int i = 0; i < checkedListBoxSites.CheckedItems.Count; i++)
            {
                checkedItem = checkedListBoxSites.CheckedItems[i].ToString();

                numberOfFixletsSelected = numberOfFixletsSelected + Convert.ToInt32(checkedItem.Substring(checkedItem.LastIndexOf("(") + 1, checkedItem.Length - (checkedItem.LastIndexOf("(") + 2)));
            }

            labelFixletsSelected.Text = "Fixlets selected: " + numberOfFixletsSelected.ToString();
        }
    }
}
