using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
// using System.Linq;
using System.Text;
using System.Windows.Forms;
using Owf.Controls;
using Microsoft.Win32;
using System.IO;
using System.Security.Cryptography;

namespace BigFixExcelConnector
{
    public partial class FormConfig : Form
    {
        // RelevanceService bes = new RelevanceService();
        RelevanceBindingEx bes = new RelevanceBindingEx();

        public FormConfig()
        {
            InitializeComponent();
        }

        private void FormConfig_Load(object sender, EventArgs e)
        {
            try
            {
                textBoxURL.Text = GetSettings("ExcelConnector", "WebReportsServer");
                textBoxUsername.Text = GetSettings("ExcelConnector", "Username");
                textBoxPassword.Text = Decrypt(GetSettings("ExcelConnector", "Password"));
                comboBoxTaskPaneLocation.Text = GetSettings("ExcelConnector", "TaskPaneLocation");
                if (comboBoxTaskPaneLocation.Text == "")
                    comboBoxTaskPaneLocation.Text = "Top";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error retrieving settings", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void buttonConfigSave_Click(object sender, EventArgs e)
        {
            try
            {
                SaveAll();
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void buttonTest_Click(object sender, EventArgs e)
        {
            try
            {
                bes.Url = textBoxURL.Text;
                String[] results = bes.GetRelevanceResult("organization of bes license", textBoxUsername.Text, textBoxPassword.Text);
                MessageBox.Show("Connected to BigFix WebReports owned by " + results[0], "Connected successfully", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if ((ex.Message.ToLower().Contains("login failed")) || (ex.Message.ToLower().Contains("remote name could not be resolved"))
                    || (ex.Message.ToLower().Contains("unable to connect to the remote server"))
                    || (ex.Message.ToLower().Contains("invalid uri"))
                    || (ex.Message.ToLower().Contains("uri prefix"))
                    || (ex.Message.ToLower().Contains("connection was closed"))
                    )
                {
                    MessageBox.Show(ex.Message, "Not connected", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else if (ex.Message.ToLower().Contains("object reference not set to an instance of an object"))
                {
                    MessageBox.Show("Error: " + ex.Message + " - Note that this BigFix AddIn only works for BES 7.2 or above.", "Check BigFix version", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
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

        private void SaveAll()
        {
            try
            {
                SetSettings("ExcelConnector", "WebReportsServer", textBoxURL.Text);
                SetSettings("ExcelConnector", "Username", textBoxUsername.Text);
                SetSettings("ExcelConnector", "Password", Encrypt(textBoxPassword.Text));
                SetSettings("ExcelConnector", "RefreshCache", "True");

                SetSettings("ExcelConnector", "TaskPaneLocation", comboBoxTaskPaneLocation.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error saving settings to registry", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

    }
}
