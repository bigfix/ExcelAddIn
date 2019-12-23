using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using AddinExpress.MSO;
using Microsoft.Win32;
using TreeViewSerialization;

namespace BigFixExcelConnector
{
    public partial class FormWizard : Form
    {
        #region Initialization

        // RelevanceService bes = new RelevanceService();
        RelevanceBindingEx bes = new RelevanceBindingEx();

        public Boolean userStarted = false;

        String webReportsURL = "";
        String userName = "";
        String password = "";
        String refreshContentCache = "";
        String concatenationSeparator = "%0A";
        String nullSubstitution = "<none>";
        String autofitRowHeightMax = "100";
        String timeOutSecs = "100";

        int maxCellLengthExcel14 = 32767; // Actual max seems to be 32767 Excel 2010
        int maxCellLengthExcel12 = 8200; // Actual max seems to be 8203 Excel 2007
        int maxCellLengthExcel11 = 900; // Actual max seems to be 911 for Excel 2003
        int maxCellLength = 900; // Default max

        int maxRowsExcel14 = 1048575; //It was 130000 up to version 3.3 of the Excel Connector
        int maxRowsExcel12 = 1048575; // It was 130000 up to version 3.3
        int maxRowsExcel11 = 65000;
        int maxRowsExcel = 65000;

        int maxColumnsExcel14 = 2048; // Actual limit is ???
        int maxColumnsExcel12 = 2048; // Actual limit is 16K
        int maxColumnsExcel11 = 255;
        int maxColumnsExcel = 255;

        int maxColumnWidth = 80;

        Excel.XlSheetVisibility HiddenSheetVisibility = Excel.XlSheetVisibility.xlSheetHidden;

        String[] results;
        String[] resultsAllSites;
        String[] resultsSitesNonOperators;
        String[] resultsSitesTemp;
        String[] resultsComputerGroups;

        String[] resultsGlobalProperties;
        String[] resultsAnalysisProperties;
        String[] resultsDuplicateProperties;
        String queryStringForGlobalProperties = "(name of it & \"||\" & id of it as string) of bes properties whose (analysis flag of it = false and name of it does not start with \"_BESClient\")";
        String queryStringForAnalysisProperties = "(name of source analysis of it & \"||\" & name of it & \"!!\" & id of it as string) of bes properties whose (analysis flag of it = true and active flag of best activations of source analysis of it = true)";
        String queryStringForDuplicateProperties = "unique values whose (multiplicity of it > 1) of (it as lowercase) of names of bes properties";

        Boolean bChildTrigger = true;
        Boolean bParentTrigger = true;

        public String QueryTime = "";
        TimeSpan t;
        TimeSpan t2;
        HiResTimer totalTime = new HiResTimer();

        String reportName = "";

        public Boolean NeedToRefreshWizard = false;

        // String selectedAttributes = "";

        // Xml tag for node, e.g. 'node' in case of <node></node>
        private const string XmlNodeTag = "fixlet";

        // Xml attributes for node e.g. <node text="Name" tag="" imageindex="1"></node>
        private const string XmlNodeTextAtt = "text";
        private const string XmlNodeTagAtt = "tag";
        private const string XmlNodeImageIndexAtt = "imageindex";


        public FormWizard()
        {
            InitializeComponent();
            this.DoubleBuffered = true;

            // The application will support connection via HTTPS using a certificate signed by a trusted authority
            // However, for testing purposes, some Web Reports server might have a local cert that will cause error
            // The following is used to circumvent the error
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(delegate { return true; });

            System.Net.ServicePointManager.SecurityProtocol =
                System.Net.SecurityProtocolType.Tls |
                System.Net.SecurityProtocolType.Tls11 |
                System.Net.SecurityProtocolType.Tls12 |
                System.Net.SecurityProtocolType.Ssl3;

            // Get URL, username and password from registry
            RetrieveSettings();

            Thread firstThread = new Thread(new ThreadStart(cacheSites));
            firstThread.Priority = ThreadPriority.Highest;
            firstThread.Start();

            String excelVersion = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Version;

            if (excelVersion == "16.0")
            {
                excelVersion = excelVersion + " - " + "2016";
                textBoxExcelVersion.Text = excelVersion;
                maxCellLength = maxCellLengthExcel14;
                maxRowsExcel = maxRowsExcel14;
                maxColumnsExcel = maxColumnsExcel14;
            }
            else if (excelVersion == "15.0")
            {
                excelVersion = excelVersion + " - " + "2013";
                textBoxExcelVersion.Text = excelVersion;
                maxCellLength = maxCellLengthExcel14;
                maxRowsExcel = maxRowsExcel14;
                maxColumnsExcel = maxColumnsExcel14;
            }
            else if (excelVersion == "14.0")
            {
                excelVersion = excelVersion + " - " + "2010";
                textBoxExcelVersion.Text = excelVersion;
                maxCellLength = maxCellLengthExcel14;
                maxRowsExcel = maxRowsExcel14;
                maxColumnsExcel = maxColumnsExcel14;
            }
            else if (excelVersion == "12.0")
            {
                excelVersion = excelVersion + " - " + "2007";
                textBoxExcelVersion.Text = excelVersion;
                maxCellLength = maxCellLengthExcel12;
                maxRowsExcel = maxRowsExcel12;
                maxColumnsExcel = maxColumnsExcel12;
            }
            else if (excelVersion == "11.0")
            {
                excelVersion = excelVersion + " - " + "2003";
                textBoxExcelVersion.Text = excelVersion;
                maxCellLength = maxCellLengthExcel11;
                maxRowsExcel = maxRowsExcel11;
                maxColumnsExcel = maxColumnsExcel11;
            }
            else
            {
                textBoxExcelVersion.Text = excelVersion;
            }
            
            comboBoxANDsORs.Text = "AND";
            comboBoxANDsORsForComputerGroup.Text = "OR";
            comboBoxANDsORsForComputerGroup.BackColor = System.Drawing.Color.FromArgb(231, 243, 241); 

            wizard1.NextEnabled = false;

            // Extracts the version number from the Assembly and display in Form title
            Assembly ass = Assembly.GetExecutingAssembly();
            string version;
            if (ass != null)
            {
                FileVersionInfo FVI = FileVersionInfo.GetVersionInfo(ass.Location);
                version = String.Format("{0} v{1:0}.{2:0}.{3:0}",
                              FVI.ProductName,
                              FVI.FileMajorPart.ToString(),
                              FVI.FileMinorPart.ToString(),
                              FVI.FileBuildPart.ToString());
            }
            else
            {
                version = "Unknown";
            }
            this.Text = "Query Wizard - " + version;
        }

        #endregion
                
        #region Wizard Page 1 Select Objects

        private void wizardPageChooseObjects_CloseFromNext(object sender, Gui.Wizard.PageEventArgs e)
        {
            if (    listBoxObjects.SelectedItem.ToString() == "BES Fixlets" || 
                    listBoxObjects.SelectedItem.ToString() == "Results of BES Fixlets" )
            {
                e.Page = wizardPageFixletSites;
            }
            else 
            {
                e.Page = wizardPageAttributes;
            }
        }

        private void wizard1_Load(object sender, EventArgs e)
        {
            listBoxObjects.SelectedItem = LoadFromExcel("A2");

            if (listBoxObjects.SelectedItem == null || listBoxObjects.SelectedItem.ToString() == "")
                wizard1.NextEnabled = false;
            else
                wizard1.NextEnabled = true;
        }

        private void listBoxObjects_SelectedIndexChanged(object sender, EventArgs e)
        {

            String selectedObj = listBoxObjects.SelectedItem.ToString();

            if (selectedObj == "")
                wizard1.NextEnabled = false;
            else
                wizard1.NextEnabled = true;

            switch (selectedObj)
            {
                case "BES Actions":
                     textBoxObjDesc.Text = "BES Actions inspectors are used to access information about the actions which have been issued by the BES Operators. You can iterate over the actions to create lists. Each action may have several properties that can be examined.";
                    break;
                case "BES Computer Groups":
                    textBoxObjDesc.Text = "BES Computer Groups inspectors return an iterated list of computer groups, as defined in the BES Console.";
                    break;
                case "BES Computers":
                    textBoxObjDesc.Text = "BES Computers inspectors return lists of the computers currently visible through the BES Console and Web Reports. Use these inspectors to retrieve computer properties.";
                    break;
                case "BES Custom Sites":
                    textBoxObjDesc.Text = "BES Custom Sites inspectors return the names and IDs of the custom site objects. Custom sites are those created locally within an IBM BigFix installation, rather than those subscribed from IBM BigFix.";
                    break;
                case "BES Fixlets":
                    textBoxObjDesc.Text = "BES Fixlets inspectors allow you to iterate over the BES Fixlet messages to create lists of various Fixlet properties such as name, ID, source severity, source release dates, etc.";
                    break;
                case "BES Properties":
                    textBoxObjDesc.Text = "BES Properties inspectors allow you to select one specific Retrieve Property, then summaries the count of the values reported by managed computers";
                    break;
                case "BES Sites":
                    textBoxObjDesc.Text = "BES Sites inspectors return the names and IDs of the specified site objects. BES Sites represent all supported types, including external sites, master action sites and operator sites. For custom sites, use the BES Custom Sites inspector.";
                    break;
                case "BES Users":
                    textBoxObjDesc.Text = "BES Users inspectors let you keep track of the users authorized to use the BES Console. You can iterate over the users, producing lists containing information such as the name and authorization level.";
                    break;
                case "BES UnmanagedAssets":
                    textBoxObjDesc.Text = "BES UnamanagedAssets inspectors provide access to externally sourced data, such as that derived from Nmap scans on client computers. The results, such as OS, Device Type, Network Card Vendor, and Open Ports, are uploaded to the BES Server for storage and analysis. These Inspectors provide a way to monitor and report on mobile or hand-held devices that are not traditional BES Clients, but instead use \"microAgents\" to report their status. For more information on currently supported devices, consult the IBM BigFix support pages.";
                    break;
                case "BES UnmanagedAsset Fields":
                    textBoxObjDesc.Text = "BES UnmanagedAsset Fields inspectors provide access to the individual fields of various unmanaged assets. Each field consists of a name / value pair, analogous to BES properties. There are three types of fields:\r\n• IdentifyingField: Each asset must have one IdentifyingField, such as a MAC Address, which is used to identify and correlate different reports from the same asset.\r\n• FilterableField: These are displayed in the Console in both the Unmanaged Asset list and the unmanaged asset document, allowing sorting and filtering.\r\n• NonFilterable: These are only displayed in the Unmanaged Assets document, and typically return a large amount of data, such as a list of vulnerabilities.";
                    break;
                case "Results of BES Fixlets":
                    textBoxObjDesc.Text = "These Inspectors allow you to inspect the results of BES Fixlet messages.";
                    break;
                default:
                    textBoxObjDesc.Text = "";
                    break;
            }

        }

        #endregion 

        #region Wizard Page 2 Fixlet Sites

        private void wizardPageFixletSites_ShowFromNext(object sender, EventArgs e)
        {
            wizard1.NextEnabled = false;

            if (checkedListBoxSites.Items.Count == 0)
            {
                getSites();
                loadFixletSitesFromExcel();
                calcFixlets();
            }

            if (checkedListBoxSites.CheckedItems.Count > 0)
                wizard1.NextEnabled = true;
            else
                wizard1.NextEnabled = false;
        }

        private void buttonSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxSites.Items.Count; i++)
            {
                if (checkedListBoxSites.GetItemChecked(i) == false)
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

        public void cacheSites()
        {
            try
            {
                // System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo((AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.LanguageSettings.get_LanguageID(Office.MsoAppLanguageID.msoLanguageIDUI)); 

                String queryString = "";

                // If data already in cache, do not have to query Web Reports again
                // This section combines the original 2 queries into one for efficiency

                if (resultsSitesTemp == null || refreshContentCache == "True")
                {
                    queryString = "(operator site flag of it as string & \"|X|\" & (if (name of it = \"Enterprise Security\") then (\"Patches for Windows\") else (name of it)) & \" (\" & number of fixlets of it as string & \")\") of all bes sites";
                    resultsSitesTemp = bes.GetRelevanceResult(queryString, userName, password);

                    resultsAllSites = new String[resultsSitesTemp.Length];
                    resultsSitesNonOperators = new String[resultsSitesTemp.Length];

                    for (int i = 0; i < resultsSitesTemp.Length; i++)
                    {
                        resultsAllSites[i] = resultsSitesTemp[i].Substring(resultsSitesTemp[i].IndexOf("|X|") + 3);
                        if (resultsSitesTemp[i].StartsWith("False"))
                            resultsSitesNonOperators[i] = resultsSitesTemp[i].Substring(resultsSitesTemp[i].IndexOf("|X|") + 3);
                    }
                }

                /*
                // If data already in cache, do not have to query Web Reports again
                if (resultsSitesNonOperators == null)
                {
                    queryString = "((if (name of it = \"Enterprise Security\") then (\"Patches for Windows (English)\") else (name of it)) & \" (\" & number of fixlets of it as string & \")\") of all bes sites whose (operator site flag of it = false)";
                    resultsSitesNonOperators = bes.GetRelevanceResult(queryString, userName, password);
                }

                if (resultsAllSites == null)
                {
                    queryString = "((if (name of it = \"Enterprise Security\") then (\"Patches for Windows (English)\") else (name of it)) & \" (\" & number of fixlets of it as string & \")\") of all bes sites";
                    resultsAllSites = bes.GetRelevanceResult(queryString, userName, password);
                } */

                if (resultsGlobalProperties == null || refreshContentCache == "True")
                {
                    // queryString = "unique values of names of bes properties whose (analysis flag of it = false and name of it does not start with \"_BESClient\")";
                    queryString = queryStringForGlobalProperties;
                    resultsGlobalProperties = bes.GetRelevanceResult(queryString, userName, password);
                }

                if (resultsAnalysisProperties == null || refreshContentCache == "True")
                {
                    // queryString = "unique values of names of bes properties whose (analysis flag of it = true and active flag of activation of source analysis of it = true)";
                    // queryString = "(name of source analysis of it & \"||\" & name of it) of bes properties whose (analysis flag of it = true and active flag of best activations of source analysis of it = true)";
                    queryString = queryStringForAnalysisProperties;
                    resultsAnalysisProperties = bes.GetRelevanceResult(queryString, userName, password);
                }


            }
            catch (Exception ex)
            {
                String nullMsg = ex.Message;
            }
        }

        private void getSites()
        {
            try
            {
                checkBoxHideOperatorSites.Enabled = false;
                checkBoxHideOperatorSites.Update();

                checkedListBoxSites.Items.Clear();
                checkedListBoxSites.Items.Add("Retrieving Fixlet sites...");
                checkedListBoxSites.Update();

                String queryString = "";

                // If data already in cache, do not have to query Web Reports again
                // This section combines the original 2 queries into one for efficiency

                if (resultsAllSites == null || resultsSitesNonOperators == null)
                {
                    queryString = "(operator site flag of it as string & \"|X|\" & (if (name of it = \"Enterprise Security\") then (\"Patches for Windows\") else (name of it)) & \" (\" & number of fixlets of it as string & \")\") of all bes sites";

                    resultsSitesTemp = bes.GetRelevanceResult(queryString, userName, password);

                    resultsAllSites = new String[resultsSitesTemp.Length];
                    resultsSitesNonOperators = new String[resultsSitesTemp.Length];

                    for (int i = 0; i < resultsSitesTemp.Length; i++)
                    {
                        resultsAllSites[i] = resultsSitesTemp[i].Substring(resultsSitesTemp[i].IndexOf("|X|") + 3);
                        if (resultsSitesTemp[i].StartsWith("False"))
                            resultsSitesNonOperators[i] = resultsSitesTemp[i].Substring(resultsSitesTemp[i].IndexOf("|X|") + 3);
                    }
                }

                if (checkBoxHideOperatorSites.Checked == true)
                {
                    results = resultsSitesNonOperators;
                }
                else
                {
                    results = resultsAllSites;
                }

                Array.Sort(results);

                checkedListBoxSites.Items.Clear();
                for (int i = 0; i < results.Length; i++)
                {
                    if (results[i] != null)
                        checkedListBoxSites.Items.Add(results[i]);
                }

                textBoxFixletNumber.Text = "";
                buttonSelectAll.Focus();

                checkBoxHideOperatorSites.Enabled = true;

            }
            catch (Exception ex)
            {
                processError(ex);
                checkedListBoxSites.Items.Clear();
            }

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
            textBoxFixletNumber.Text = numberOfFixletsSelected.ToString();

            if (numberOfFixletsSelected > 0)
                wizard1.NextEnabled = true;
            else
                wizard1.NextEnabled = false;
        }

        private void checkBoxHideOperatorSites_CheckedChanged(object sender, EventArgs e)
        {
            getSites();
        }
        
        private void checkedListBoxSites_SelectedValueChanged(object sender, EventArgs e)
        {
            calcFixlets();
        }

        #endregion

        #region Wizard Page 3 Select Attributes

        private void WizardPageAttributes_ShowFromBack(object sender, EventArgs e)
        {
            wizard1.NextEnabled = true;

            // These are the columns for BES Properties that are pseudo.
            // They are Count, Percent and Graph.
            // Remove the rows when the user tracks back in the Wizard
            if (listBoxObjects.SelectedItem.ToString() == "BES Properties")
            {
                listBoxPropertiesSelected.Items.RemoveAt(3);
                listBoxPropertiesSelected.Items.RemoveAt(2);
                listBoxPropertiesSelected.Items.RemoveAt(1);
            }
        }

        private void WizardPageAttributes_ShowFromNext(object sender, EventArgs e)
        {
            labelAttributes.Text = "Attributes Available for " + listBoxObjects.SelectedItem.ToString() + ":";
            textBoxAttributesSelected.Text = "0";
            
            // If BES Fixlets selected, and the treeview is empty or if it contains info from other objects...
            if (listBoxObjects.SelectedItem.ToString() == "BES Fixlets" &&
                (listBoxPropertiesSelected.Items.Count == 0 || treeView1.Tag.ToString() != "BES Fixlets" || NeedToRefreshWizard ))
            {
                listBoxPropertiesSelected.Items.Clear();
                dataGridView1.Rows.Clear();
                dataGridViewComputerGroup.Rows.Clear();
                dataGridViewFilters.Rows.Clear();
                textBoxValue.Text = "";
                comboBoxOperator.SelectedValue = "";
                comboBoxBoolean.SelectedValue = "";
                checkBoxGroupAnalysis.Visible = false;
                processAttrXML("fixlets");
                treeView1.Nodes[0].Expand();
                treeView1.Tag = "BES Fixlets";
                WizardPage3AttributesWidgetState(true);
                wizard1.NextEnabled = false;
                restoreProperties();
                loadFiltersFromExcel(listBoxObjects.SelectedItem.ToString());
                comboBoxANDsORsForComputerGroup.Visible = false;
                labelComputerGroupsAndOr.Visible = false;
            }
            else if (listBoxObjects.SelectedItem.ToString() == "Results of BES Fixlets" && 
                (listBoxPropertiesSelected.Items.Count == 0 || treeView1.Tag.ToString() != "Results of BES Fixlets" || NeedToRefreshWizard ))
            {
                String selectedObject = listBoxObjects.SelectedItem.ToString();
                selectedObject = selectedObject.Replace(" ", "_").ToLower();
                listBoxPropertiesSelected.Items.Clear();
                dataGridView1.Rows.Clear();
                dataGridViewComputerGroup.Rows.Clear();
                dataGridViewFilters.Rows.Clear();
                textBoxValue.Text = "";
                comboBoxOperator.SelectedValue = "";
                comboBoxBoolean.SelectedValue = "";
                checkBoxGroupAnalysis.Visible = false;
                processAttrXML(selectedObject);
                treeView1.Tag = listBoxObjects.SelectedItem.ToString();
                WizardPage3AttributesWidgetState(true);
                restoreProperties();
                loadFiltersFromExcel(listBoxObjects.SelectedItem.ToString());
                comboBoxANDsORsForComputerGroup.Visible = false;
                labelComputerGroupsAndOr.Visible = false;
            }
            else if (listBoxObjects.SelectedItem.ToString() == "Results of BES Actions" &&
                (listBoxPropertiesSelected.Items.Count == 0 || treeView1.Tag.ToString() != "Results of BES Actions" || NeedToRefreshWizard ))
            {
                String selectedObject = listBoxObjects.SelectedItem.ToString();
                selectedObject = selectedObject.Replace(" ", "_").ToLower();
                listBoxPropertiesSelected.Items.Clear();
                dataGridView1.Rows.Clear();
                dataGridViewComputerGroup.Rows.Clear();
                dataGridViewFilters.Rows.Clear();
                textBoxValue.Text = "";
                comboBoxOperator.SelectedValue = "";
                comboBoxBoolean.SelectedValue = "";
                checkBoxGroupAnalysis.Visible = false;
                processAttrXML(selectedObject);
                treeView1.Tag = listBoxObjects.SelectedItem.ToString();
                WizardPage3AttributesWidgetState(true);
                restoreProperties();
                loadFiltersFromExcel(listBoxObjects.SelectedItem.ToString());
                comboBoxANDsORsForComputerGroup.Visible = false;
                labelComputerGroupsAndOr.Visible = false;
            }
            else if (listBoxObjects.SelectedItem.ToString() == "BES Computers" && 
                    (listBoxPropertiesSelected.Items.Count == 0 || treeView1.Tag.ToString() != "BES Computers" || NeedToRefreshWizard ))
            {
                listBoxPropertiesSelected.Items.Clear();
                dataGridView1.Rows.Clear();
                dataGridViewComputerGroup.Rows.Clear();
                dataGridViewFilters.Rows.Clear();
                textBoxValue.Text = "";
                comboBoxOperator.SelectedValue = "";
                comboBoxBoolean.SelectedValue = "";
                checkBoxGroupAnalysis.Visible = true;
                getProperties();
                treeView1.Tag = "BES Computers";
                WizardPage3AttributesWidgetState(true);
                restoreProperties();
                loadFiltersFromExcel(listBoxObjects.SelectedItem.ToString());
                NeedToRefreshWizard = false;
                comboBoxANDsORsForComputerGroup.Visible = true;
                labelComputerGroupsAndOr.Visible = true;
            }
            else if (listBoxObjects.SelectedItem.ToString() == "BES Properties" &&
                    (listBoxPropertiesSelected.Items.Count == 0 || treeView1.Tag.ToString() != "BES Properties" || NeedToRefreshWizard))
            {
                listBoxPropertiesSelected.Items.Clear();
                dataGridView1.Rows.Clear();
                dataGridViewComputerGroup.Rows.Clear();
                dataGridViewFilters.Rows.Clear();
                textBoxValue.Text = "";
                comboBoxOperator.SelectedValue = "";
                comboBoxBoolean.SelectedValue = "";
                checkBoxGroupAnalysis.Visible = true;
                treeView1.Tag = "BES Properties";
                WizardPage3AttributesWidgetState(false);
                getProperties();
                restoreProperties();
                loadFiltersFromExcel(listBoxObjects.SelectedItem.ToString());
                NeedToRefreshWizard = false;
                comboBoxANDsORsForComputerGroup.Visible = true;
                labelComputerGroupsAndOr.Visible = true;
            }
            else if (listBoxPropertiesSelected.Items.Count == 0 || listBoxObjects.SelectedItem.ToString() != treeView1.Tag.ToString() || NeedToRefreshWizard)
            {
                String selectedObject = listBoxObjects.SelectedItem.ToString();
                selectedObject = selectedObject.Substring(selectedObject.IndexOf(" ") + 1).ToLower();
                selectedObject = selectedObject.Replace(" ", "_");
                listBoxPropertiesSelected.Items.Clear();
                dataGridView1.Rows.Clear();
                dataGridViewComputerGroup.Rows.Clear();
                dataGridViewFilters.Rows.Clear();
                textBoxValue.Text = "";
                comboBoxOperator.SelectedValue = "";
                comboBoxBoolean.SelectedValue = "";
                checkBoxGroupAnalysis.Visible = false;
                processAttrXML(selectedObject);
                treeView1.Nodes[0].Expand();
                treeView1.Tag = listBoxObjects.SelectedItem.ToString();
                WizardPage3AttributesWidgetState(true);
                restoreProperties();
                loadFiltersFromExcel(listBoxObjects.SelectedItem.ToString());
                comboBoxANDsORsForComputerGroup.Visible = false;
                labelComputerGroupsAndOr.Visible = false;
            }
            else
            {
                wizard1.NextEnabled = true;
            }

            if (listBoxPropertiesSelected.Items.Count == 0)
                wizard1.NextEnabled = false;
            else
                wizard1.NextEnabled = true;

            textBoxAttributesSelected.Text = listBoxPropertiesSelected.Items.Count.ToString();

        }

        private void WizardPage3AttributesWidgetState(Boolean ShowOrNot)
        {
            buttonUnselectAllAttributes.Enabled = ShowOrNot;
            buttonSelectAllAttributes.Enabled = ShowOrNot;
            checkBoxRowHeightAutoFit.Enabled = ShowOrNot;
            numericUpDownRowHeightMaximum.Enabled = ShowOrNot;
            checkBoxConcatenation.Enabled = ShowOrNot;
            textBoxConcatenationSeparator.Enabled = ShowOrNot;
            labelNullSubstitution.Enabled = ShowOrNot;
            textBoxNull.Enabled = ShowOrNot;
            pictureBoxNullSub.Enabled = ShowOrNot;
            pictureBoxConcat.Enabled = ShowOrNot;
            pictureBoxRowHeightAutoFit.Enabled = ShowOrNot;
            buttonTop.Enabled = ShowOrNot;
            buttonMoveUp.Enabled = ShowOrNot;
            buttonMoveDown.Enabled = ShowOrNot;
            buttonBottom.Enabled = ShowOrNot;
            treeView1.CheckBoxes = ShowOrNot;
        }

        private void wizardPageAttributes_CloseFromBack(object sender, Gui.Wizard.PageEventArgs e)
        {
            if (listBoxObjects.SelectedItem.ToString() == "BES Fixlets")
            {
                e.Page = wizardPageFixletSites;
            }
            else
            {
                e.Page = wizardPageChooseObjects;
            }
        }

        private void restoreProperties31()
        {
            Excel.Worksheet storageWorksheet;
            try
            {
                storageWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                Excel.Range rangeStorage;
                Excel.Range previouslyStoredAttrs;

                rangeStorage = storageWorksheet.get_Range("A2", "A2");
                String[] previouslySelectedAttrs;
                char[] delimiters = new char[] { '|' };

                if (rangeStorage.Value.ToString() == treeView1.Tag.ToString())
                {
                    previouslyStoredAttrs = storageWorksheet.get_Range("A4", "A4");
                    previouslySelectedAttrs = previouslyStoredAttrs.Value.ToString().Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    foreach (String s in previouslySelectedAttrs)
                    {
                        TraverseTreeViewToCheck(treeView1, s);
                    }
                }

                // Restoring the previously selected BES Property
                if (rangeStorage.Value.ToString() == treeView1.Tag.ToString() && rangeStorage.Value.ToString() == "BES Properties")
                {
                    previouslyStoredAttrs = storageWorksheet.get_Range("A4", "A4");
                    previouslySelectedAttrs = previouslyStoredAttrs.Value.ToString().Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    String selectedProperty = previouslySelectedAttrs[0].Substring(previouslySelectedAttrs[0].IndexOf("!!") + 2);
                    listBoxPropertiesSelected.Items.Add(new FixletProperty(selectedProperty, selectedProperty, "String", "")); 
                }

                if (listBoxPropertiesSelected.Items.Count > 0)
                    wizard1.NextEnabled = true;
                else
                    wizard1.NextEnabled = false;

            }
            catch (Exception ex)
            {
                String nullMsg = ex.Message;
            }
        }

        private void restoreProperties()
        {
            Excel.Worksheet storageWorksheet;
            try
            {
                storageWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                Excel.Range rangeStorage;
                Excel.Range previouslyStoredAttrs;
                String[] previouslySelectedAttrs;
                rangeStorage = storageWorksheet.get_Range("A2", "A2");
                char[] delimiters = new char[] { '|' };

                if (    rangeStorage.Value.ToString() == treeView1.Tag.ToString()    && 
                        rangeStorage.Value.ToString() == "BES Computers" )
                {
                    if (rangeStorage.Value.ToString() == treeView1.Tag.ToString())
                    {
                        previouslyStoredAttrs = storageWorksheet.get_Range("A7", "A7");
                        previouslySelectedAttrs = previouslyStoredAttrs.Value.ToString().Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                        foreach (String s in previouslySelectedAttrs)
                        {
                            TraverseTreeViewToCheckForBESComputers(treeView1, s);
                        }
                    }
                }
                else
                {
                    if (rangeStorage.Value.ToString() == treeView1.Tag.ToString())
                    {
                        previouslyStoredAttrs = storageWorksheet.get_Range("A4", "A4");
                        previouslySelectedAttrs = previouslyStoredAttrs.Value.ToString().Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                        foreach (String s in previouslySelectedAttrs)
                        {
                            TraverseTreeViewToCheck(treeView1, s);
                        }
                    }
                }

                // Restoring the previously selected BES Property
                if (rangeStorage.Value.ToString() == treeView1.Tag.ToString() && rangeStorage.Value.ToString() == "BES Properties")
                {
                    previouslyStoredAttrs = storageWorksheet.get_Range("A7", "A7");
                    previouslySelectedAttrs = previouslyStoredAttrs.Value.ToString().Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    String selectedProperty = previouslySelectedAttrs[0].Substring(0, previouslySelectedAttrs[0].IndexOf("^"));
                    listBoxPropertiesSelected.Items.Add(new FixletProperty(
                                            selectedProperty, 
                                            selectedProperty, 
                                            previouslySelectedAttrs[0].Substring(previouslySelectedAttrs[0].IndexOf("^") +1), 
                                            ""));

                    /*
                    previouslyStoredAttrs = storageWorksheet.get_Range("A4", "A4");
                    previouslySelectedAttrs = previouslyStoredAttrs.Value.ToString().Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    String selectedProperty = previouslySelectedAttrs[0].Substring(previouslySelectedAttrs[0].IndexOf("!!") + 2);
                    listBoxPropertiesSelected.Items.Add(new FixletProperty(selectedProperty, selectedProperty, "String", ""));
                    */
                }

                if (listBoxPropertiesSelected.Items.Count > 0)
                    wizard1.NextEnabled = true;
                else
                    wizard1.NextEnabled = false;

            }
            catch (Exception ex)
            {
                String nullMsg = ex.Message;
            }
        }

        private void getProperties31()
        {
            try
            {
                treeView1.Nodes.Clear();
                treeView1.Nodes.Add("Retrieving properties...");
                treeView1.Update();
                
                // Populate treeview with global properites (those not part of analyses)
                if (resultsGlobalProperties == null)
                {
                    String queryString = "unique values of names of bes properties whose (analysis flag of it = false and name of it does not start with \"_BESClient\")";
                    // String queryString = "(name of it & \"||\" & id of it as string) of bes properties whose (analysis flag of it = false and name of it does not start with \"_BESClient\")";
                    resultsGlobalProperties = bes.GetRelevanceResult(queryString, userName, password);
                    results = resultsGlobalProperties;
                }
                else
                {
                    results = resultsGlobalProperties;
                }

                treeView1.Nodes.Clear();

                TreeNode nd = new TreeNode();
                nd.Text = "Global Properties";
                nd.ImageIndex = 5;
                nd.Tag = "Object";

                // treeView1.Nodes.Add("Global Properties");
                treeView1.Nodes.Add(nd);

                for (int i = 0; i < results.Length; i++)
                {
                    TreeNode childNode = new TreeNode();
                    childNode.Text = results[i];
                    if (results[i].ToLower() == "last report time")
                    {
                        childNode.Tag = "Time";
                        childNode.ImageIndex = 7;
                    }
                    else
                    {
                        childNode.Tag = "String";
                        childNode.ImageIndex = 6;
                    }
                    nd.Nodes.Add(childNode); 
                }

                nd.ExpandAll();
                treeView1.Update();

                // Populate treeview with analysis properites, but activated
                if (resultsAnalysisProperties == null)
                {
                    // String queryString2 = "unique values of names of bes properties whose (analysis flag of it = true and active flag of activation of source analysis of it = true)";
                    String queryString2 = "(name of source analysis of it & \"||\" & name of it) of bes properties whose (analysis flag of it = true and active flag of best activations of source analysis of it = true)";
                    resultsAnalysisProperties = bes.GetRelevanceResult(queryString2, userName, password);
                    results = resultsAnalysisProperties;
                }
                else
                {
                    results = resultsAnalysisProperties;
                }

                TreeNode nd2 = new TreeNode();
                nd2.Text = "Analysis Properties";
                nd2.ImageIndex = 5;
                nd2.Tag = "Object";

                // treeView1.Nodes.Add("Global Properties");
                treeView1.Nodes.Add(nd2);

                if (checkBoxGroupAnalysis.Checked == true)
                {
                    // This section for analysis properties grouped by Analysis for easier reading
                    String currentAnalysisName = "";
                    String previousAnalysisName = "";
                    String propertyName = "";
                    Array.Sort(results);

                    TreeNode analysisNameNode = null ;

                    for (int j = 0; j < results.Length; j++)
                    {
                        currentAnalysisName = results[j].Substring(0, results[j].IndexOf("||"));
                        
                        if (previousAnalysisName == currentAnalysisName)
                        {
                            propertyName = results[j].Substring(results[j].IndexOf("||") + 2);

                            TreeNode analysisPropertyNode = new TreeNode();
                            analysisPropertyNode.Text = propertyName;
                            analysisPropertyNode.Tag = "String";
                            analysisPropertyNode.ImageIndex = 6;
                            analysisNameNode.Nodes.Add(analysisPropertyNode);
                        }
                        else
                        {
                            propertyName = results[j].Substring(results[j].IndexOf("||") + 2);
                            // MessageBox.Show(currentAnalysisId + " " + currentAnalysisName);
                            analysisNameNode = new TreeNode();
                            analysisNameNode.Text = currentAnalysisName;
                            analysisNameNode.Tag = "Object";
                            analysisNameNode.ImageIndex = 5;
                            nd2.Nodes.Add(analysisNameNode);

                            TreeNode analysisPropertyNode = new TreeNode();
                            analysisPropertyNode.Text = propertyName;
                            analysisPropertyNode.Tag = "String";
                            analysisPropertyNode.ImageIndex = 6;
                            analysisNameNode.Nodes.Add(analysisPropertyNode);
                        }

                        previousAnalysisName = currentAnalysisName;
                    }
                }
                else
                {
                    // Alphabetical listing of analysis properties
                    String[] justAnalysisProps = new String[results.Length];

                    for (int j = 0; j < results.Length; j++)
                    {
                        justAnalysisProps[j] = results[j].Substring(results[j].LastIndexOf("||") + 2);
                    }
                    Array.Sort(justAnalysisProps);

                    for (int i = 0; i < justAnalysisProps.Length; i++)
                    {
                        TreeNode analysisNameNode = new TreeNode();
                        analysisNameNode.Text = justAnalysisProps[i];
                        analysisNameNode.Tag = "String";
                        analysisNameNode.ImageIndex = 6;
                        nd2.Nodes.Add(analysisNameNode);

                    }
                }
            }
            catch (Exception ex)
            {
                processError(ex);
                treeView1.Nodes.Clear();
            }

        }

        private void getProperties()
        {
            try
            {
                treeView1.Nodes.Clear();
                treeView1.Nodes.Add("Retrieving properties...");
                treeView1.Update();

                // Detect properties with duplicate names
                if (resultsDuplicateProperties == null)
                {
                    String queryString = queryStringForDuplicateProperties;
                    resultsDuplicateProperties = bes.GetRelevanceResult(queryString, userName, password);
                }

                // Populate treeview with global properites (those not part of analyses)
                if (resultsGlobalProperties == null)
                {
                    // String queryString = "unique values of names of bes properties whose (analysis flag of it = false and name of it does not start with \"_BESClient\")";
                    // String queryString = "(name of it & \"||\" & id of it as string) of bes properties whose (analysis flag of it = false and name of it does not start with \"_BESClient\")";
                    String queryString = queryStringForGlobalProperties;
                    resultsGlobalProperties = bes.GetRelevanceResult(queryString, userName, password);
                    results = resultsGlobalProperties;
                }
                else
                {
                    results = resultsGlobalProperties;
                }

                treeView1.Nodes.Clear();

                TreeNode nd = new TreeNode();
                nd.Text = "Global Properties";
                nd.ImageIndex = 5;
                nd.Tag = "Object";

                // treeView1.Nodes.Add("Global Properties");
                treeView1.Nodes.Add(nd);

                Array.Sort(results);

                for (int i = 0; i < results.Length; i++)
                {
                    TreeNode childNode = new TreeNode();
                    childNode.Text = results[i].Substring(0, results[i].IndexOf("||"));
                    if (childNode.Text.ToLower() == "last report time")
                    {
                        childNode.Tag = "Time";
                        childNode.ImageIndex = 7;
                    }
                    else
                    {
                        childNode.Tag = "String";
                        childNode.ImageIndex = 6;
                    }

                    Boolean DuplicateProp = false;

                    for (int j = 0; j < resultsDuplicateProperties.Length; j++)
                    {
                        if (childNode.Text.ToLower() == resultsDuplicateProperties[j])
                        {
                            DuplicateProp = true;
                            break;
                        }
                    }

                    if (DuplicateProp)
                    {
                        childNode.Tag = results[i].Substring(results[i].IndexOf("||") + 2) + "*";
                        childNode.ToolTipText = "ID: " + results[i].Substring(results[i].IndexOf("||") + 2) + " (Diplicate Property Names Found)";
                    }
                    else
                    {
                        childNode.Tag = results[i].Substring(results[i].IndexOf("||") + 2);
                        childNode.ToolTipText = "ID: " + results[i].Substring(results[i].IndexOf("||") + 2);
                    }

                    nd.Nodes.Add(childNode);
                }

                nd.ExpandAll();
                treeView1.Update();

                // Populate treeview with analysis properites, but activated
                if (resultsAnalysisProperties == null)
                {
                    // String queryString2 = "unique values of names of bes properties whose (analysis flag of it = true and active flag of activation of source analysis of it = true)";
                    // String queryString2 = "(name of source analysis of it & \"||\" & name of it) of bes properties whose (analysis flag of it = true and active flag of best activations of source analysis of it = true)";
                    String queryString2 = queryStringForAnalysisProperties;
                    resultsAnalysisProperties = bes.GetRelevanceResult(queryString2, userName, password);
                    results = resultsAnalysisProperties;
                }
                else
                {
                    results = resultsAnalysisProperties;
                }


                TreeNode nd2 = new TreeNode();
                nd2.Text = "Analysis Properties";
                nd2.ImageIndex = 5;
                nd2.Tag = "Object";

                // treeView1.Nodes.Add("Global Properties");
                treeView1.Nodes.Add(nd2);

                if (checkBoxGroupAnalysis.Checked == true)
                {
                    // This section for analysis properties grouped by Analysis for easier reading
                    String currentAnalysisName = "";
                    String previousAnalysisName = "";
                    String propertyName = "";
                    Array.Sort(results);

                    TreeNode analysisNameNode = null;

                    for (int j = 0; j < results.Length; j++)
                    {
                        currentAnalysisName = results[j].Substring(0, results[j].IndexOf("||"));
                        propertyName = results[j].Substring(results[j].IndexOf("||") + 2, results[j].IndexOf("!!") - results[j].IndexOf("||") - 2);

                        Boolean DuplicateProp = false;

                        for (int k = 0; k < resultsDuplicateProperties.Length; k++)
                        {
                            if (propertyName.ToLower() == resultsDuplicateProperties[k])
                            {
                                DuplicateProp = true;
                                break;
                            }
                        }

                        if (previousAnalysisName == currentAnalysisName)
                        {

                            TreeNode analysisPropertyNode = new TreeNode();
                            analysisPropertyNode.Text = propertyName;
                            // analysisPropertyNode.Tag = "String";
                            if (DuplicateProp)
                            {
                                analysisPropertyNode.Tag = results[j].Substring(results[j].IndexOf("!!") + 2) + "*";
                                analysisPropertyNode.ToolTipText = "ID: " + results[j].Substring(results[j].IndexOf("!!") + 2) + " (Diplicate Property Names Found)";
                            }
                            else
                            {
                                analysisPropertyNode.Tag = results[j].Substring(results[j].IndexOf("!!") + 2);
                                analysisPropertyNode.ToolTipText = "ID: " + results[j].Substring(results[j].IndexOf("!!") + 2);
                            }
                            analysisPropertyNode.ImageIndex = 6;
                            analysisNameNode.Nodes.Add(analysisPropertyNode);
                        }
                        else
                        {
                            analysisNameNode = new TreeNode();
                            analysisNameNode.Text = currentAnalysisName;
                            analysisNameNode.Tag = "Object";
                            analysisNameNode.ImageIndex = 5;
                            nd2.Nodes.Add(analysisNameNode);

                            TreeNode analysisPropertyNode = new TreeNode();
                            analysisPropertyNode.Text = propertyName;
                            // analysisPropertyNode.Tag = "String";
                            if (DuplicateProp)
                            {
                                analysisPropertyNode.Tag = results[j].Substring(results[j].IndexOf("!!") + 2) + "*";
                                analysisPropertyNode.ToolTipText = "ID: " + results[j].Substring(results[j].IndexOf("!!") + 2) + " (Diplicate Property Names Found)";
                            }
                            else
                            {
                                analysisPropertyNode.Tag = results[j].Substring(results[j].IndexOf("!!") + 2);
                                analysisPropertyNode.ToolTipText = "ID: " + results[j].Substring(results[j].IndexOf("!!") + 2);
                            }
                            analysisPropertyNode.ImageIndex = 6;
                            analysisNameNode.Nodes.Add(analysisPropertyNode);
                        }

                        previousAnalysisName = currentAnalysisName;
                    }
                }
                else
                {
                    // Alphabetical listing of analysis properties
                    String[] ArrayWithPropertiesInFront = new String[results.Length];
                    String PropName = "";
                    String AnalysisName = "";
                    String PropID = "";

                    for (int j = 0; j < results.Length; j++)
                    {
                        // results data looking this this. We want to put the Property Name first for sorting
                        // BES Component Versions||BES Web Reports Version!!1, 204, 6
                        // BES Relay Status||BES Relay Installed Status!!1, 205, 1
                        PropName = results[j].Substring(results[j].IndexOf("||") + 2, results[j].IndexOf("!!") - results[j].IndexOf("||") - 2);
                        AnalysisName = results[j].Substring(0, results[j].IndexOf("||"));
                        PropID = results[j].Substring(results[j].IndexOf("!!") + 2);

                        ArrayWithPropertiesInFront[j] = PropName + "||" + AnalysisName + "!!" + PropID;
                    }

                    Array.Sort(ArrayWithPropertiesInFront);

                    Boolean DuplicateProp = false;

                    for (int i = 0; i < ArrayWithPropertiesInFront.Length; i++)
                    {
                        PropName = ArrayWithPropertiesInFront[i].Substring(0, ArrayWithPropertiesInFront[i].IndexOf("||"));
                        AnalysisName = ArrayWithPropertiesInFront[i].Substring(ArrayWithPropertiesInFront[i].IndexOf("||") + 2, ArrayWithPropertiesInFront[i].IndexOf("!!") - ArrayWithPropertiesInFront[i].IndexOf("||") - 2);
                        PropID = ArrayWithPropertiesInFront[i].Substring(ArrayWithPropertiesInFront[i].IndexOf("!!") + 2);

                        // Check to see if the property name has duplicates
                        for (int k = 0; k < resultsDuplicateProperties.Length; k++)
                        {
                            if (PropName.ToLower() == resultsDuplicateProperties[k])
                            {
                                DuplicateProp = true;
                                break;
                            }
                        }

                        TreeNode analysisNameNode = new TreeNode();
                        analysisNameNode.Text = PropName;
                        // analysisNameNode.Tag = "String";
                        if (DuplicateProp)
                        {
                            analysisNameNode.Tag = PropID + "*";
                            analysisNameNode.ToolTipText = "Analysis Name: " + AnalysisName + "\n" +
                                                           "ID: " + PropID + " (Diplicate Property Names Found)";
                        }
                        else
                        {
                            analysisNameNode.Tag = PropID;
                            analysisNameNode.ToolTipText = "Analysis Name: " + AnalysisName + "\n" +
                                                           "ID: " + PropID;
                        }
                        analysisNameNode.ImageIndex = 6;
                        nd2.Nodes.Add(analysisNameNode);

                        DuplicateProp = false;
                    }
                }
            }
            catch (Exception ex)
            {
                processError(ex);
                treeView1.Nodes.Clear();
            }

        }

        private void processAttrXML(string besTag)
        {
            try
            {
                listBoxPropertiesSelected.Items.Clear();

                this.treeView1.Nodes.Clear();
                TreeViewSerializer serializer = new TreeViewSerializer();
                serializer.DeserializeTreeView(this.treeView1, "bes.xml", besTag);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error reading XML file bes.xml with attribute information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void visitChildNodesToCheck(TreeNode node, String Attr)
        {
            if (node.Text == Attr.Substring(Attr.IndexOf("!!") + 2) && node.Parent.Text == Attr.Substring(0, Attr.IndexOf("!!")))
            {
                node.Checked = true;
            }
            else if (node.Text == Attr.Substring(Attr.IndexOf("!!") + 2) && listBoxObjects.SelectedItem.ToString() == "BES Computers")
            {
                node.Checked = true;
            }

            //Loop Through this node and its childs recursively
            for (int j = 0; j < node.Nodes.Count; j++)
                visitChildNodesToCheck(node.Nodes[j], Attr);
        }

        private void visitChildNodesToCheckForBESComputers(TreeNode node, String Attr)
        {
            if (node.Text == Attr.Substring(Attr.IndexOf("!!") + 2) && node.Parent.Text == Attr.Substring(0, Attr.IndexOf("!!")))
            {
                node.Checked = true;
            }
            // else if (node.Text == Attr.Substring(Attr.IndexOf("!!") + 2) && listBoxObjects.SelectedItem.ToString() == "BES Computers")
            else if (
                     node.Text == Attr.Substring(0, Attr.IndexOf("^")) && 
                     node.Tag.ToString() == Attr.Substring(Attr.IndexOf("^") + 1)  &&
                     (listBoxObjects.SelectedItem.ToString() == "BES Computers")
                    )
            {
                node.Checked = true;
            }

            // Loop Through this node and its childs recursively
            // Fixed in version 3.3. visitChildNodesToCheckForBESComputers() was incorrectly left as visitChildNodesToCheck() 
            // Causing Analysis properties not being restored 
            for (int j = 0; j < node.Nodes.Count; j++)
                visitChildNodesToCheckForBESComputers(node.Nodes[j], Attr);
        }

        private void TraverseTreeViewToCheck(TreeView tview, String Attr)
        {
            //Create a TreeNode to hold the Parent Node
            TreeNode temp = new TreeNode();

            //Loop through the Parent Nodes
            for (int k = 0; k < tview.Nodes.Count; k++)
            {
                //Store the Parent Node in temp
                temp = tview.Nodes[k];

                //Display the Text of the Parent Node i.e. temp
                // MessageBox.Show(temp.Text);
                if (temp.Text == Attr.Substring(Attr.IndexOf("!!") + 2) && temp.Tag.ToString() != "Object")
                    temp.Checked = true;

                //Now Loop through each of the child nodes in this parent node i.e.temp
                for (int i = 0; i < temp.Nodes.Count; i++)
                    visitChildNodesToCheck(temp.Nodes[i], Attr); //send every child to the function for further traversal
            }
        }

        private void TraverseTreeViewToCheckForBESComputers(TreeView tview, String Attr)
        {
            //Create a TreeNode to hold the Parent Node
            TreeNode temp = new TreeNode();

            //Loop through the Parent Nodes
            for (int k = 0; k < tview.Nodes.Count; k++)
            {
                //Store the Parent Node in temp
                temp = tview.Nodes[k];

                //Display the Text of the Parent Node i.e. temp
                // MessageBox.Show(temp.Tag.ToString() + " - " + Attr.Substring(Attr.IndexOf("^") + 1).TrimEnd('*'));
                if (temp.Text == Attr.Substring(0, Attr.IndexOf("^")) && temp.Tag.ToString() != "Object" && temp.Tag.ToString() == Attr.Substring(Attr.IndexOf("^")+1))
                    temp.Checked = true;

                //Now Loop through each of the child nodes in this parent node i.e.temp
                for (int i = 0; i < temp.Nodes.Count; i++)
                    visitChildNodesToCheckForBESComputers(temp.Nodes[i], Attr); //send every child to the function for further traversal
            }
        }

        private void checkBoxConcatenation_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxConcatenation.Checked)
            {
                textBoxConcatenationSeparator.Enabled = true;
            }
            else
            {
                textBoxConcatenationSeparator.Enabled = false;
            }
        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {

            #region Everything non-BES Properties
            if (bChildTrigger)
            {
                CheckAllChildren(e.Node, e.Node.Checked);
            }
            if (bParentTrigger)
            {
                CheckMyParent(e.Node, e.Node.Checked);
            }

            if (e.Node.Checked == true && e.Node.Tag.ToString() != "Object" && e.Node.Tag.ToString() != "SearchableObject")
            {
                if (e.Node.Parent != null &&
                    (e.Node.Parent.Tag.ToString() == "Object" || e.Node.Parent.Tag.ToString() == "SearchableObject") &&
                    listBoxObjects.SelectedItem.ToString() != "BES Computers" &&
                    e.Node.Parent.Text != "Common Properties" &&
                    e.Node.Parent.Text != "Extended Properties")
                {
                    listBoxPropertiesSelected.Items.Add(new FixletProperty(e.Node.Text, e.Node.Text + " of " + e.Node.Parent.Text, e.Node.Tag.ToString(), e.Node.Parent.Text));
                }
                else
                {
                    if (e.Node.Parent == null)
                    {
                        listBoxPropertiesSelected.Items.Add(new FixletProperty(e.Node.Text, e.Node.Text, e.Node.Tag.ToString(), ""));
                    }
                    else if (e.Node.Parent.Text == "Common Properties" || e.Node.Parent.Text == "Extended Properties")
                    {
                        listBoxPropertiesSelected.Items.Add(new FixletProperty(e.Node.Text, e.Node.Text, e.Node.Tag.ToString(), e.Node.Parent.Text));
                    }
                    else
                    {
                        listBoxPropertiesSelected.Items.Add(new FixletProperty(e.Node.Text, e.Node.Text, e.Node.Tag.ToString(), ""));
                    }
                }

                wizard1.NextEnabled = true;
            }
            else
            {
                for (int i = 0; i < listBoxPropertiesSelected.Items.Count; i++)
                {
                    FixletProperty fp = (FixletProperty)listBoxPropertiesSelected.Items[i];

                    if (e.Node.Parent == null && fp.Name == e.Node.Text)
                    {
                        listBoxPropertiesSelected.Items.RemoveAt(i);
                        break;
                    }

                    if (fp.Name == e.Node.Text && fp.ParentName == e.Node.Parent.Text)
                    {
                        listBoxPropertiesSelected.Items.RemoveAt(i);
                        break;
                    }

                    // This section for BES Computers
                    if (fp.Name == e.Node.Text && listBoxObjects.SelectedItem.ToString() == "BES Computers")
                    {
                        listBoxPropertiesSelected.Items.RemoveAt(i);
                        break;
                    }

                }
            }

            if (listBoxPropertiesSelected.Items.Count == 0)
            {
                wizard1.NextEnabled = false;
            }

            textBoxAttributesSelected.Text = listBoxPropertiesSelected.Items.Count.ToString();
            #endregion
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            treeView1.SelectedImageIndex = treeView1.SelectedNode.ImageIndex;

            if (listBoxObjects.SelectedItem.ToString() == "BES Properties" && e.Node.Tag.ToString() != "Object" && e.Node.Tag.ToString() != "SearchableObject")
            {
                listBoxPropertiesSelected.Items.Clear();
                listBoxPropertiesSelected.Items.Add(new FixletProperty(e.Node.Text, e.Node.Text, e.Node.Tag.ToString(), ""));
            }

            if (listBoxPropertiesSelected.Items.Count == 0)
            {
                wizard1.NextEnabled = false;
            }
            else
            {
                wizard1.NextEnabled = true;
            }

            textBoxAttributesSelected.Text = listBoxPropertiesSelected.Items.Count.ToString();

        }

        private void CheckMyParent(TreeNode tn, Boolean bCheck)
        {
            if (tn == null) return;
            if (tn.Parent == null) return;

            bChildTrigger = false;
            bParentTrigger = false;
            tn.Parent.Checked = bCheck;
            CheckMyParent(tn.Parent, bCheck);
            bParentTrigger = true;
            bChildTrigger = true;
        }

        private void CheckAllChildren(TreeNode tn, Boolean bCheck)
        {
            bParentTrigger = false;
            foreach (TreeNode ctn in tn.Nodes)
            {
                bChildTrigger = false;
                if (ctn.Checked && bCheck)
                {
                    // Do nothing because alreadyc checked
                }
                else
                {
                    ctn.Checked = bCheck;
                }
                bChildTrigger = true;

                CheckAllChildren(ctn, bCheck);
            }
            bParentTrigger = true;
        }

        #region Top-Bottom-Up-Down Buttons
        private void buttonMoveUp_Click(object sender, EventArgs e)
        {
            if (listBoxPropertiesSelected.SelectedItem != null && listBoxPropertiesSelected.SelectedIndex != 0)
            {
                listBoxPropertiesSelected.Items.Insert(listBoxPropertiesSelected.SelectedIndex + 1, listBoxPropertiesSelected.Items[listBoxPropertiesSelected.SelectedIndex - 1]);
                listBoxPropertiesSelected.Items.RemoveAt(listBoxPropertiesSelected.SelectedIndex - 1);
            }
        }

        private void buttonMoveDown_Click(object sender, EventArgs e)
        {
            if (listBoxPropertiesSelected.SelectedItem != null && listBoxPropertiesSelected.SelectedIndex != listBoxPropertiesSelected.Items.Count-1)
            {
                listBoxPropertiesSelected.Items.Insert(listBoxPropertiesSelected.SelectedIndex, listBoxPropertiesSelected.Items[listBoxPropertiesSelected.SelectedIndex + 1]);
                listBoxPropertiesSelected.Items.RemoveAt(listBoxPropertiesSelected.SelectedIndex + 1);
            }
        }

        private void buttonTop_Click(object sender, EventArgs e)
        {
            if (listBoxPropertiesSelected.SelectedItem != null && listBoxPropertiesSelected.SelectedIndex != 0)
            {
                listBoxPropertiesSelected.Items.Insert(0, listBoxPropertiesSelected.SelectedItem);
                listBoxPropertiesSelected.Items.RemoveAt(listBoxPropertiesSelected.SelectedIndex);
                listBoxPropertiesSelected.SelectedIndex = 0;
            }
        }

        private void buttonBottom_Click(object sender, EventArgs e)
        {
            if (listBoxPropertiesSelected.SelectedItem != null && listBoxPropertiesSelected.SelectedIndex != listBoxPropertiesSelected.Items.Count-1)
            {
                listBoxPropertiesSelected.Items.Insert((listBoxPropertiesSelected.Items.Count), listBoxPropertiesSelected.SelectedItem);
                listBoxPropertiesSelected.Items.RemoveAt(listBoxPropertiesSelected.SelectedIndex);
                listBoxPropertiesSelected.SelectedIndex = listBoxPropertiesSelected.Items.Count-1;
            }
        }

        #endregion Top-Bottom-Up-Down Buttons

        private void buttonSelectAllAttributes_Click(object sender, EventArgs e)
        {
            // Unselect all before reselecting again to rid of existing checks
            /*
            foreach (TreeNode childNode in treeView1.Nodes)
            {
                childNode.Checked = false;
            } */

            foreach (TreeNode mainNode in treeView1.Nodes)
            {
                if (mainNode.Checked == false)
                    mainNode.Checked = true;

                foreach (TreeNode childNode in mainNode.Nodes)
                {
                    if (childNode.Checked == false)
                        childNode.Checked = true;

                    foreach (TreeNode grandChildNode in childNode.Nodes)
                    {
                        if (grandChildNode.Checked == false)
                            grandChildNode.Checked = true;
                    }
                }
            }
            wizard1.NextEnabled = true;
        }

        private void buttonUnselectAllAttributes_Click(object sender, EventArgs e)
        {
            foreach (TreeNode childNode in treeView1.Nodes)
            {
                childNode.Checked = false;
            }
            wizard1.NextEnabled = false;
        }

        private void buttonExpandAll_Click(object sender, EventArgs e)
        {
            if (this.buttonExpandAll.Text == "Expand All")
            {
                this.treeView1.ExpandAll();
                this.buttonExpandAll.Text = "Collapse All";
            }
            else
            {
                this.treeView1.CollapseAll();
                // this.treeView1.Nodes[0].Expand();
                this.buttonExpandAll.Text = "Expand All";
            }
        }

        private void checkBoxGroupAnalysis_CheckedChanged(object sender, EventArgs e)
        {
            listBoxPropertiesSelected.Items.Clear();
            getProperties();
            treeView1.Nodes[1].Expand();

            if (listBoxPropertiesSelected.Items.Count == 0)
            {
                wizard1.NextEnabled = false;
            }
            else
            {
                wizard1.NextEnabled = true;
            }

            textBoxAttributesSelected.Text = listBoxPropertiesSelected.Items.Count.ToString();

        }

        #region The 3 Help Buttons
        private void pictureBoxNullSub_Click(object sender, EventArgs e)
        {
            MessageBox.Show("If a property does not have any data, what should the substitution string be?", "What is Null Substitution?", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void pictureBoxConcat_Click(object sender, EventArgs e)
        {
            FormHelpConcat frmHelpConcat = new FormHelpConcat();
            frmHelpConcat.ShowDialog();
        }

        private void pictureBoxRowHeightAutoFit_Click(object sender, EventArgs e)
        {
            FormHelpAutofit frmHelpAutofit = new FormHelpAutofit();
            frmHelpAutofit.ShowDialog();
        }
        #endregion The 3 Help Buttons
        #endregion Wizard Page 3

        #region Wizard Page 4 Run Query

        private void wizardPageExecuteQuery_ShowFromNext(object sender, EventArgs e)
        {
            // textBoxRelevance.Text = "";
            // textBoxRelevance.Update();
            statusMessage("", true, "Normal");
            labelFilterCriteria.Text = "Add Filter Criteria to " + listBoxObjects.SelectedItem.ToString() + ":";

            comboBoxProperty.Items.Clear();
            comboBoxProperty.Text = "";
            comboBoxOperator.Text = "";
            textBoxValue.Text = "";
            comboBoxBoolean.Text = "";

            if (listBoxObjects.SelectedItem.ToString() == "BES Computers")
            {
                ArrayList al = new ArrayList();

                foreach (TreeNode mainNode in treeView1.Nodes)
                {
                    foreach (TreeNode midNode in mainNode.Nodes)
                    {
                        foreach (TreeNode childNode in midNode.Nodes)
                        {
                            al.Add(childNode);
                        }
                        if (midNode.Tag.ToString() != "Object")
                            al.Add(midNode);
                    }
                }

                al.Sort(new NodeSorter());

                TreeNode computerGroupNode = new TreeNode();
                computerGroupNode.Text = "Computer Groups";
                computerGroupNode.Tag = "Computer Groups";

                comboBoxProperty.Items.Add(computerGroupNode);
                comboBoxProperty.DisplayMember = "Text";

                foreach (TreeNode prop in al)
                {
                    comboBoxProperty.Items.Add(prop);
                    comboBoxProperty.DisplayMember = "Text";
                }
            }
            else if (listBoxObjects.SelectedItem.ToString() == "BES Properties")
            {
                TreeNode computerGroupNode = new TreeNode();
                computerGroupNode.Text = "Computer Groups";
                computerGroupNode.Tag = "Computer Groups";
                comboBoxProperty.Items.Add(computerGroupNode);
                comboBoxProperty.DisplayMember = "Text";

                listBoxPropertiesSelected.Items.Add(new FixletProperty("Count", "Count", "Integer", ""));
                listBoxPropertiesSelected.Items.Add(new FixletProperty("Percent", "Percent", "Integer", ""));
                listBoxPropertiesSelected.Items.Add(new FixletProperty("Graph", "Graph", "String", ""));

            }
            else
            {
                foreach (TreeNode mainNode in treeView1.Nodes)
                {
                    if ((String)mainNode.Tag != "Object" && (String)mainNode.Tag != "SearchableObject")
                    {
                        comboBoxProperty.Items.Add(mainNode);
                        comboBoxProperty.DisplayMember = "Text";
                    }

                    if (mainNode.Text == "Extended Properties")
                    {
                        TreeNode newNode = new TreeNode();
                        newNode.Text = "---------------";
                        newNode.Name = "---------------"; 
                        newNode.Tag = "Object";
                        // comboBoxProperty.Items.Add("---------------");
                        comboBoxProperty.Items.Add(newNode);
                    }


                    foreach (TreeNode childNode in mainNode.Nodes)
                    {
                        if (
                            ((String)childNode.Parent.Tag != "Object" &&
                            (String)childNode.Parent.Tag != "SearchableObject" &&
                            (String)childNode.Tag == "Object") ||
                            ((String)childNode.Parent.Text == "Common Properties" && (String)childNode.Tag != "Object") ||
                            ((String)childNode.Parent.Text == "Extended Properties" && (String)childNode.Tag != "Object")
                            )
                        {
                            comboBoxProperty.Items.Add(childNode);
                            comboBoxProperty.DisplayMember = "Text";
                        }
                        else if ((String)childNode.Parent.Tag == "SearchableObject")
                        {
                            TreeNode newNode = new TreeNode();
                            newNode = (TreeNode)childNode.Clone();
                            newNode.Text = childNode.Text + " of " + childNode.Parent.Text;
                            comboBoxProperty.Items.Add(newNode);
                            comboBoxProperty.DisplayMember = "Text";
                        }
                    }
                }
            }
        }

        private void wizardPageExecuteQuery_CloseFromNext(object sender, Gui.Wizard.PageEventArgs e)
        {
            // saveToExcel();

            try
            {
                SetSettings("ExcelConnector", "ConcatenationSeparator", textBoxConcatenationSeparator.Text);
                SetSettings("ExcelConnector", "NullSubstitution", textBoxNull.Text);
                SetSettings("ExcelConnector", "AutofitRowHeightMax", numericUpDownRowHeightMaximum.Value.ToString());
                SetSettings("ExcelConnector", "TimeoutSecs", numericUpDownTimeOut.Value.ToString());
                SetSettings("ExcelConnector", "CheckConcat", checkBoxConcatenation.Checked.ToString());
                SetSettings("ExcelConnector", "CheckAutofit", checkBoxRowHeightAutoFit.Checked.ToString());
                SetSettings("ExcelConnector", "CheckSort", checkBoxSortResults.Checked.ToString());
                SetSettings("ExcelConnector", "CheckTimeout", checkBoxTimeout.Checked.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error saving settings to registry", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void buttonRun_Click(object sender, EventArgs e)
        {
            try
            {
                totalTime.Start();

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

                // Save the name of the report for reuse
                Excel.Range CellA1 = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.get_Range("A1", "A1");
                if (CellA1.Value == null)
                    reportName = "";
                else
                    reportName = CellA1.Value.ToString();

                if (listBoxObjects.SelectedItem.ToString() == "BES Fixlets" ||
                    listBoxObjects.SelectedItem.ToString() == "Results of BES Fixlets" ||
                    listBoxObjects.SelectedItem.ToString() == "Results of BES Actions")
                {
                    ExecuteFixlet();
                }
                else if (listBoxObjects.SelectedItem.ToString() == "BES Computers")
                {
                    ExecuteComputer();
                }
                else if (listBoxObjects.SelectedItem.ToString() == "BES Properties")
                {
                    ExecuteProperty();
                }
                else
                {
                    String selectedObj = listBoxObjects.SelectedItem.ToString();
                    ExecuteObjectQuery(selectedObj);
                }

                saveToExcel();

                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Visible = true;
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.UserControl = true;
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;

                totalTime.Stop();
                t2 = TimeSpan.FromMilliseconds(totalTime.ElapsedMicroseconds / 1000);
                toolStripStatusLabelEvalTime.Text = "Total: " + t2.ToString().Remove(t2.ToString().Length - 4) + " / " + toolStripStatusLabelEvalTime.Text;

                Excel.Worksheet storageWorksheet;
                Excel.Range rangeStorage;
                storageWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                rangeStorage = storageWorksheet.get_Range("A16", "A16");
                rangeStorage.Value = t2.ToString().Remove(t2.ToString().Length - 4) + " / " + t.ToString().Remove(t.ToString().Length - 4);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error in Execute Query", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

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

        private void ProcessExcel(Boolean refreshQuery)
        {
            try
            {
                int colCount = 0;
                Excel.Worksheet hiddenWorksheet;
                String selectedBESObjectName = "";

                if (refreshQuery)
                {
                    hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                    colCount = Convert.ToInt32(hiddenWorksheet.get_Range("A6", "A6").Value.ToString());
                    selectedBESObjectName = hiddenWorksheet.get_Range("A2", "A2").Value.ToString();
                    if (hiddenWorksheet.get_Range("A17", "A17").Value != null)
                    {
                        reportName = hiddenWorksheet.get_Range("A17", "A17").Value.ToString();
                    }
                    else
                    {
                        reportName = selectedBESObjectName;
                    }
                }
                else
                {
                    colCount = listBoxPropertiesSelected.Items.Count;
                    selectedBESObjectName = listBoxObjects.SelectedItem.ToString();

                    String SavedBESObject = LoadFromExcel("A2");

                    if (reportName == "" || SavedBESObject != selectedBESObjectName)
                        reportName = selectedBESObjectName;
                }

                if (colCount > maxColumnsExcel)
                {
                    colCount = maxColumnsExcel;
                }

                // Report info at the first line
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[2, 1] = "Generated by: " + userName;
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[3, 1] = DateTime.Now.ToString("f");
                // (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[3, 1] = selectedBESObjectName;
                
                Excel.Range rangeToRightAlign;
                rangeToRightAlign = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("B2", "C3");
                rangeToRightAlign.Cells.HorizontalAlignment = -4152; // Right align the cells
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[2, 2] = colCount + " columns";
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[3, 2] = results.Length + " rows";

                // (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "C3").Font.Bold = true;
                // (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "C3").Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "C3").Font.Size = "10";

                Excel.Range titleRow = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", "CA1");
                titleRow.Select();
                titleRow.RowHeight = 33;
                titleRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(184, 204, 228));

                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = reportName;
                // Name the sheet to be the same as the report name. Sheet name has to be less than 31 characters
                String sheetName = "";
                if (reportName.Length > 30)
                {
                    sheetName = reportName.Substring(0, 28) + "..";
                }
                else
                {
                    sheetName = reportName;
                }
                ((Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet).Name = sheetName;

                Excel.Range title = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", "A1");
                title.ClearFormats();
                title.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(184, 204, 228));
                title.InsertIndent(1);
                title.Font.Size = 18;
                title.VerticalAlignment = -4108; // xlCenter

                Excel.Range rowSeparator = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A3", "CA3");
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(184, 204, 228)); // 
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1; // xlContinuous
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = 4; // xlThick

                // This block writes the column headers

                if (refreshQuery)
                {
                    Excel.Range rangeStorage;
                    hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                    rangeStorage = hiddenWorksheet.get_Range("A7", "A7"); ;
                    String[] previouslySelectedAttrs;
                    char[] delimiters = new char[] { '|' };

                    previouslySelectedAttrs = rangeStorage.Value.ToString().Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < colCount; j++)
                    {
                        (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[4, j + 1] = previouslySelectedAttrs[j].Substring(0, previouslySelectedAttrs[j].IndexOf('^'));
                    }
                }
                else
                {
                    FixletProperty fp2;
                    for (int j = 0; j < colCount; j++)
                    {
                        fp2 = (FixletProperty)listBoxPropertiesSelected.Items[j];
                        (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[4, j + 1] = fp2.DisplayName;
                    }
                }

                // Make the header row Bold
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A4", "A4").EntireRow.Font.Bold = true;

                int maxRows = maxRowsExcel;

                String[] singleRow;
                String[] splitter = new String[1] { "$x$" };

                if (results.Length > maxRows)
                {
                    // MessageBox.Show("Warning: the query returns " + results.Length + " rows, but only " + maxRows + " will be displayed.", "Too much data", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    maxRows = results.Length;
                }

                String[,] dataBlock = new string[maxRows, colCount];

                Boolean truncatedCells = false;

                for (int i = 0; i < maxRows; i++)
                {
                    singleRow = results[i].Split(splitter, System.StringSplitOptions.None);
                    for (int j = 0; j < colCount; j++)
                    {
                        if (singleRow[j].Length > maxCellLength)
                        {
                            dataBlock[i, j] = singleRow[j].Substring(0, maxCellLength);
                            truncatedCells = true;
                        }
                        else
                            dataBlock[i, j] = singleRow[j];
                    }
                }

                // Do not do this if no data returned
                Excel.Range range;
                range = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A5", Missing.Value);

                if (maxRows > 0)
                {
                    range = range.get_Resize(maxRows, colCount);
                    range.Cells.ClearFormats();
                    // Writes the array into Excel
                    // This is probably the single thing that sped up the report the most, but writing array
                    range.Value = dataBlock;
                    range.Font.Size = "10";
                }

                Excel.Range range2;
                range2 = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A4", Missing.Value);
                range2 = range2.get_Resize(maxRows + 1, colCount);
                range2.Select();
                range2.AutoFilter("1", "<>", Excel.XlAutoFilterOperator.xlOr, "", true);

                // Since version 8.1, Web Reports escapes the control characters, so line feeds are not embedded properly. This works around the problem
                if (checkBoxConcatenation.Checked == true && textBoxConcatenationSeparator.Text.ToLower().Contains("%0a"))
                {
                    (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.DisplayAlerts = false; // Surpresses the dialog box if nothing to replace
                    string LF = "\n";
                    range2.Replace("%0a", LF, Missing.Value, Missing.Value, false);
                    (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.DisplayAlerts = true; // reset
                }

                if (refreshQuery)
                {
                    Excel.Range rangeStorage;
                    hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                    rangeStorage = hiddenWorksheet.get_Range("A7", "A7"); ;
                    String[] previouslySelectedAttrs;
                    char[] delimiters = new char[] { '|' };

                    previouslySelectedAttrs = rangeStorage.Value.ToString().Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < colCount; j++)
                    {
                        if (previouslySelectedAttrs[j].Substring(previouslySelectedAttrs[j].IndexOf('^') + 1) == "Integer")
                        {
                            Excel.Range rangeInt;
                            rangeInt = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(ExcelColumnLetter(j) + "5", ExcelColumnLetter(j) + (maxRows + 5).ToString());
                            rangeInt.NumberFormat = "0";
                            rangeInt.Cells.HorizontalAlignment = -4152; // Right align the number
                            rangeInt.Value = rangeInt.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text
                        }
                        // else if (((FixletProperty)listBoxPropertiesSelected.Items[j]).DataType == "Date")
                        else if (previouslySelectedAttrs[j].Substring(previouslySelectedAttrs[j].IndexOf('^') + 1) == "Date")
                        {
                            Excel.Range rangeDate;
                            rangeDate = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(ExcelColumnLetter(j) + "5", ExcelColumnLetter(j) + (maxRows + 5).ToString());
                            rangeDate.NumberFormat = "yyyy/mm/dd";
                            rangeDate.Cells.HorizontalAlignment = -4152; // Right align the Date
                            rangeDate.Value = rangeDate.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text
                        }
                        // else if (((FixletProperty)listBoxPropertiesSelected.Items[j]).DataType == "Time")
                        else if (previouslySelectedAttrs[j].Substring(previouslySelectedAttrs[j].IndexOf('^') + 1) == "Time" ||
                                previouslySelectedAttrs[j].Substring(0, previouslySelectedAttrs[j].IndexOf('^')).ToLower() == "last report time" )
                        {
                            Excel.Range rangeTime;
                            rangeTime = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(ExcelColumnLetter(j) + "5", ExcelColumnLetter(j) + (maxRows + 5).ToString());
                            rangeTime.NumberFormat = "[$-409]yyyy/mm/dd hh:mm AM/PM;@";
                            rangeTime.Cells.HorizontalAlignment = -4152; // Right align the Time
                            rangeTime.Value = rangeTime.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text
                        }
                        else if (   previouslySelectedAttrs[j].Substring(0, previouslySelectedAttrs[j].IndexOf('^')).ToLower() == "ram" ||
                                    previouslySelectedAttrs[j].Substring(0, previouslySelectedAttrs[j].IndexOf('^')).ToLower() == "ram - unix" ||
                                    previouslySelectedAttrs[j].Substring(0, previouslySelectedAttrs[j].IndexOf('^')).ToLower() == "free space on system drive" ||
                                    previouslySelectedAttrs[j].Substring(0, previouslySelectedAttrs[j].IndexOf('^')).ToLower() == "total size of system drive")
                        {
                            Excel.Range rangeMB;
                            rangeMB = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(ExcelColumnLetter(j) + "5", ExcelColumnLetter(j) + (maxRows + 5).ToString());
                            rangeMB.NumberFormat = "0";
                            rangeMB.Cells.HorizontalAlignment = -4152; // Right align the number
                            rangeMB.Value = rangeMB.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text

                            rangeMB.NumberFormat = "[>=1024]#,##0.00,\" GB\";#,##0.00\" MB\"";
                            rangeMB.Value = rangeMB.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text
                        }

                    }
                }
                else
                {
                    for (int i = 0; i < colCount; i++)
                    {
                        if (((FixletProperty)listBoxPropertiesSelected.Items[i]).DataType == "Integer")
                        {
                            Excel.Range rangeInt;
                            rangeInt = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(ExcelColumnLetter(i) + "5", ExcelColumnLetter(i) + (maxRows + 5).ToString());
                            rangeInt.NumberFormat = "0";
                            rangeInt.Cells.HorizontalAlignment = -4152; // Right align the number
                            rangeInt.Value = rangeInt.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text
                        }
                        else if (((FixletProperty)listBoxPropertiesSelected.Items[i]).DataType == "Date")
                        {
                            Excel.Range rangeDate;
                            rangeDate = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(ExcelColumnLetter(i) + "5", ExcelColumnLetter(i) + (maxRows + 5).ToString());
                            rangeDate.NumberFormat = "yyyy/mm/dd";
                            rangeDate.Cells.HorizontalAlignment = -4152; // Right align the Date
                            rangeDate.Value = rangeDate.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text
                        }
                        else if (((FixletProperty)listBoxPropertiesSelected.Items[i]).DataType == "Time" ||
                                ((FixletProperty)listBoxPropertiesSelected.Items[i]).Name.ToLower() == "last report time")
                        {
                            Excel.Range rangeTime;
                            rangeTime = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(ExcelColumnLetter(i) + "5", ExcelColumnLetter(i) + (maxRows + 5).ToString());
                            rangeTime.NumberFormat = "[$-409]yyyy/mm/dd hh:mm AM/PM;@";
                            rangeTime.Cells.HorizontalAlignment = -4152; // Right align the Time
                            rangeTime.Value = rangeTime.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text
                        }
                        else if (   ((FixletProperty)listBoxPropertiesSelected.Items[i]).Name.ToLower() == "ram" || 
                                    ((FixletProperty)listBoxPropertiesSelected.Items[i]).Name.ToLower() == "ram - unix" || 
                                    ((FixletProperty)listBoxPropertiesSelected.Items[i]).Name.ToLower() == "free space on system drive" ||
                                    ((FixletProperty)listBoxPropertiesSelected.Items[i]).Name.ToLower() == "total size of system drive"   )
                        {
                            Excel.Range rangeMB;
                            rangeMB = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(ExcelColumnLetter(i) + "5", ExcelColumnLetter(i) + (maxRows + 5).ToString());
                            rangeMB.NumberFormat = "0";
                            rangeMB.Cells.HorizontalAlignment = -4152; // Right align the number
                            rangeMB.Value = rangeMB.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text

                            rangeMB.NumberFormat = "[>=1024]#,##0.00,\" GB\";#,##0.00\" MB\"";
                            rangeMB.Value = rangeMB.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text
                        }

                    }
                }

                // Formats the column width nicely to see the content, but not too wide and limit to maximum of 80
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Columns.AutoFit();
                for (int i = 0; i < colCount; i++)
                {
                    Excel.Range rangeCol;
                    rangeCol = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(ExcelColumnLetter(i) + "1", ExcelColumnLetter(i) + "1");
                    if (Convert.ToInt32(rangeCol.ColumnWidth) > maxColumnWidth)
                    {
                        rangeCol.ColumnWidth = maxColumnWidth;
                    }
                }

                #region BES Properties specific formatting
                // BES Properties specific formatting =====================================================================================================================
                if (LoadFromExcel("A2") == "BES Properties" || (listBoxObjects.SelectedItem != null && listBoxObjects.SelectedItem.ToString() == "BES Properties" && refreshQuery == false))
                {

                    double largestNumberForPropertyPercentage = 0;
                    double totalCount = 0;

                    for (int i = 0; i < maxRows; i++)
                    {
                        if (Convert.ToDouble(dataBlock[i, 2]) > largestNumberForPropertyPercentage)
                        {
                            largestNumberForPropertyPercentage = Convert.ToDouble(dataBlock[i, 2]);
                        }

                        totalCount = totalCount + Convert.ToDouble(dataBlock[i, 1]);                        
                    }
                    
                    double GraphMultiplier = 1;
                    if (largestNumberForPropertyPercentage != 0)
                    {
                        GraphMultiplier = 100 / largestNumberForPropertyPercentage;
                        GraphMultiplier = Math.Round(GraphMultiplier, 1);
                    }


                    // Do the following only if there is data returned
                    if (maxRows > 0)
                    {
                        // Format the Percent column with 2 decimals such as 100.00 or 0.73.
                        range = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("C5", Missing.Value);
                        range = range.get_Resize(maxRows, 1);
                        range.NumberFormat = "0.00";

                        // This is a special treatment for writing the bar graph for BES Property
                        // D5 is the start of the column with the bar graph
                        range = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("D5", Missing.Value);
                        range = range.get_Resize(maxRows, 1);

                        String fontName = "Britannic Bold";
                        Font testFont = new Font(fontName, 8.0f);

                        // Arial will work for the simulated Graph, but Britannic Bold is great
                        if (testFont.Name == fontName)
                            range.Font.Name = fontName;
                        else
                            range.Font.Name = "Arial";

                        range.Font.Size = "8";

                        range.FormulaR1C1 = "=REPT(\"|\",RC[-1]*" + GraphMultiplier.ToString() + ")";
                        range.Columns.AutoFit();

                        Excel.Range sortColumn = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("B4", Missing.Value);
                        range = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A5", Missing.Value);
                        range = range.get_Resize(maxRows, 4);
                        range.Sort(sortColumn, Excel.XlSortOrder.xlDescending, Type.Missing, Type.Missing, Excel.XlSortOrder.xlDescending,
                                                        Type.Missing, Excel.XlSortOrder.xlDescending,
                                                        Excel.XlYesNoGuess.xlGuess, Type.Missing,
                                                        Type.Missing, Excel.XlSortOrientation.xlSortColumns, Excel.XlSortMethod.xlPinYin);

                        // Copy the first 10 rows for charting
                        Excel.Range SourceData = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A4", "A14");
                        Excel.Range DestinationData = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("AA4", Type.Missing);
                        SourceData.Copy(DestinationData);

                        SourceData = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("B4", "B14");
                        DestinationData = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("AC4", Type.Missing);
                        SourceData.Copy(DestinationData);

                        // The property results are often too long. The following stripes 30 characters from the string
                        DestinationData = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("AB4", "AB14");
                        DestinationData.FormulaR1C1 = "=IF(LEN(RC[-1]) > 30, MID(RC[-1], 1, 20) & \"...\" & MID(RC[-1], LEN(RC[-1])-6, 7), IF(LEN(RC[-1])=0, \" \", RC[-1] ) )";

                        Excel.Range CellForOthers;
                        Excel.Range CellTotalForOthers = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("AC15", Type.Missing);
                        Excel.Range chartData;

                        if (maxRows > 10)
                        {
                            CellForOthers = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("AB15", Type.Missing);
                            CellForOthers.Value = "Others";
                            // CellTotalForOthers = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("B15", "B" + (maxRows + 4).ToString());
                            // CellTotalForOthers = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("AC15", Type.Missing);
                            CellTotalForOthers.FormulaR1C1 = "=SUM(RC[-27]:R[" + (maxRows - 11).ToString() + "]C[-27])";
                            chartData = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("AB4", "AC15");
                        }
                        else
                        {
                            chartData = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("AB4", "AC"+ (maxRows+4).ToString());
                        }
                        
                        Excel.ChartObjects xlCharts = (Excel.ChartObjects)((Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet).ChartObjects(Type.Missing);

                        // Position the Chart using cell E4 as the reference
                        Excel.Range PositioningCell = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("E4", Type.Missing);
                        // Excel.ChartObject myChart = xlCharts.Add((double)PositioningCell.Left, (double)PositioningCell.Top, 300, 300);
                        // Excel.ChartObject myChart = xlCharts.Add((double)PositioningCell.Left, (double)PositioningCell.Top, 500, 400);
                        Excel.ChartObject myChart = xlCharts.Add((double)PositioningCell.Left, (double)PositioningCell.Top, 450, 300);
                        Excel.Chart pieChart = myChart.Chart;
                        pieChart.SetSourceData(chartData);
                        pieChart.ChartType = Excel.XlChartType.xlPie;
                        // pieChart.ChartTitle.Delete();
                        pieChart.ChartTitle.Text = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A4", Type.Missing).Value.ToString();

                        // There there is only one row, Excel assumes the column/row (data versus series) incorrectly. This is the needed hint
                        if (maxRows == 1)
                            pieChart.SetSourceData(chartData, Excel.XlRowCol.xlColumns);

                        // Pie Chart, get rid of default Fill, then make the outline not visible
                        // Excel.ShapeRange shp = ((Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet).Shapes.get_Range(1);
                        Excel.ShapeRange shp = ((Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet).Shapes.get_Range(1);
                        shp.Fill.Visible = Office.MsoTriState.msoFalse;
                        shp.Line.Visible = Office.MsoTriState.msoFalse;

                        // Add Label and the percentage to the pie slices, but only if Others is not too big a slice
                        if (LoadFromActiveSheet("AC15") != "" && ((double)CellTotalForOthers.Value / totalCount * 100) > 80)
                            ((Excel.Series)pieChart.SeriesCollection(1)).ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowPercent, Type.Missing, Type.Missing, Type.Missing);
                        else
                            ((Excel.Series)pieChart.SeriesCollection(1)).ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowLabelAndPercent, Type.Missing, Type.Missing, Type.Missing);
                    }
               }
               // =============================================================================================================================================================
               #endregion BES Properties

                if (checkBoxRowHeightAutoFit.Checked)
                {
                    range.Rows.AutoFit();
                    range.VerticalAlignment = -4160; // This is how to determine value
                    // Record Macro that does the VerticalAlignment, in this case Top Align
                    // Edit Macro to see the code. Note that value is xlTop
                    // Ctrl-G to get into debug window, then type ?xlTop to decode constant

                    for (int i = 5; i < maxRows+5; i++)
                    {
                        Excel.Range rangeRow;
                        rangeRow = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A" + i, "A" + i);
                        if (Convert.ToInt32(rangeRow.RowHeight) > numericUpDownRowHeightMaximum.Value)
                        {
                            rangeRow.RowHeight = numericUpDownRowHeightMaximum.Value;
                        }
                    }
                }
                else
                {
                    // Formats the row height to the same as with Row 2. 
                    // Row 1 has the higher Title, Row 2 is the normal height
                    range.RowHeight = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "A2").Height;
                }

                // Place the cursor in cell A1 - which is at the start of the document
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", "A1").Select();


                // Save the query info into a hidden worksheet
                // saveToExcel();

                if (truncatedCells && results.Length > maxRows)
                    statusMessage(results.Length + " rows returned from query, " + maxRows + " rows written to Excel. Cells over " + maxCellLength + " chars truncated.", false, "Warning");
                else if (results.Length > maxRows)
                    statusMessage(results.Length + " rows returned from query, " + maxRows + " rows written to Excel.", false, "Warning");
                else if (truncatedCells)
                    statusMessage(results.Length + " rows returned from query. Cells over " + maxCellLength + " chars truncated.", false, "Warning");
                else
                    statusMessage(results.Length + " rows returned from query.", false, "Success");

                buttonRun.Enabled = true;

            }
            catch (Exception ex)
            {
                processError(ex);

                // MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                buttonRun.Enabled = true;
                statusMessage(ex.Message, true, "Error");
            }
        }

        private void ExecuteProperty()
        {
            try
            {
                buttonRun.Enabled = false;
                buttonRun.Update();

                statusMessage("Processing query on " + listBoxObjects.SelectedItem.ToString() + "...", true, "Normal");

                textBoxRelevance.Text = "";
                textBoxRelevance.Update();

                // Clear spreadsheet
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.ClearContents();
                DeleteCharts();

                // Computer Group Filter
                String BESComputerRelevance = "";

                if (dataGridView1.Rows.Count > 0)
                {
                    String prop = "";
                    String setOperator = "";
                    String searchValue = "";

                    if (comboBoxANDsORsForComputerGroup.Text == "OR")
                        BESComputerRelevance = "\n\telements of union of (\n";
                    else
                        BESComputerRelevance = "\n\telements of intersection of (\n";

                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        prop = dataGridView1.Rows[i].Cells[0].Value.ToString();
                        setOperator = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        searchValue = dataGridView1.Rows[i].Cells[2].Value.ToString().ToLower();

                        BESComputerRelevance = BESComputerRelevance +
                                    "\t\tmember sets of bes computer groups whose (name of it as lowercase " +
                                    setOperator + " \"" + searchValue + "\")";

                        if (i != dataGridView1.Rows.Count - 1)
                            BESComputerRelevance = BESComputerRelevance + "; \n";
                    }
                    BESComputerRelevance = BESComputerRelevance + ")";
                }

                // Add properties clause to query
                // MessageBox.Show(((FixletProperty)listBoxPropertiesSelected.Items[0]).DataType.TrimEnd('*'));

                String RelevanceStatementForProperty = " & \"$x$\" & \n\titem 1 of it as string & \"$x$\" & \n\titem 2 of it as string & \"$x$\" &  \" \") \nof (\n\tit, \n\tmultiplicity of it, \n\tmultiplicity of it as floating point * 100/(xxx as floating point)) \nof unique values \nof values of results \nof bes properties \n\twhose (\n\t\tname of it = \"" + listBoxPropertiesSelected.Items[0].ToString() + "\" and id of it = (" + ((FixletProperty)listBoxPropertiesSelected.Items[0]).DataType.TrimEnd('*') + "))";

                if (((FixletProperty)(listBoxPropertiesSelected.Items[0])).Name.ToLower() == "ram" ||
                                 ((FixletProperty)(listBoxPropertiesSelected.Items[0])).Name.ToLower() == "free space on system drive" ||
                                 ((FixletProperty)(listBoxPropertiesSelected.Items[0])).Name.ToLower() == "total size of system drive" ||
                                 ((FixletProperty)(listBoxPropertiesSelected.Items[0])).Name.ToLower() == "ram - unix")
                {
                    RelevanceStatementForProperty = "(\t(if (item 0 of it ends with \" MB\")\n\t\tthen (preceding text of first \" MB\" of item 0 of it)\n\t\telse (item 0 of it))" + RelevanceStatementForProperty;
                }
                else if (((FixletProperty)(listBoxPropertiesSelected.Items[0])).Name.ToLower() == "last report time")
                {
                    RelevanceStatementForProperty = "(\t(if ((year of date (local time zone) of it) as integer = 1980) \n\t\tthen (\" \") \n\t\telse (\n\t\t\t(year of it as string & \"/\" & \n\t\t\t month of it as two digits & \"/\" & \n\t\t\t day_of_month of it as two digits) of date (local time zone) of it & \" \" & \n\t\t\t(two digit hour of it as string & \":\" & \n\t\t\t two digit minute of it as string) of time (local time zone) of it)) \n\t\t\t\tof (item 0 of it as time)" + RelevanceStatementForProperty;
                }
                else
                {
                    RelevanceStatementForProperty = "(\titem 0 of it" + RelevanceStatementForProperty;
                }

                String[] TotalNumberOfProperties;
                TotalNumberOfProperties = bes.GetRelevanceResult("number of " + RelevanceStatementForProperty.Substring(RelevanceStatementForProperty.IndexOf("values of results")), userName, password);
                RelevanceStatementForProperty = RelevanceStatementForProperty.Replace("xxx", TotalNumberOfProperties[0]);

                if (dataGridView1.Rows.Count > 0)
                {
                    RelevanceStatementForProperty = RelevanceStatementForProperty.Replace("results \nof bes properties",
                                                "results from (" + BESComputerRelevance + ") \nof bes properties");
                }

                // Build the relevance statement here
                textBoxRelevance.Text = RelevanceStatementForProperty;

                String queryString = textBoxRelevance.Text;

                ProcessRelevanceThenWriteToExcel(RelevanceStatementForProperty, false);

            }
            catch (Exception ex)
            {
                processError(ex);
                buttonRun.Enabled = true;
                statusMessage(ex.Message, true, "Error");
            }

        }

        private void ExecuteObjectQuery(String obj)
        {
            try
            {
                buttonRun.Enabled = false;
                buttonRun.Update();

                statusMessage("Processing query on " + listBoxObjects.SelectedItem.ToString() + "...", true, "Normal");

                textBoxRelevance.Text = "";
                textBoxRelevance.Update();

                String objStr = "";
                // String unmanagedAssetFieldStr = "";

                if (obj.ToLower() == "bes unmanagedasset fields")
                    objStr = "\nof fields ";
                else
                    objStr = "\nof \n\t" + obj;

                // Clear spreadsheet
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.ClearContents();
                DeleteCharts();

                // Add whose clauses
                if (dataGridView1.Rows.Count > 0)
                {
                    objStr = objStr + " \n\t\twhose (\n\t\t\t";
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[3].Value.ToString() == "String")
                        {
                            objStr = objStr + "(" + dataGridView1.Rows[i].Cells[0].Value.ToString() + " of it as lowercase " + dataGridView1.Rows[i].Cells[1].Value.ToString().ToLower() + " \"" + dataGridView1.Rows[i].Cells[2].Value.ToString().ToLower() + "\")";
                        }
                        else if (dataGridView1.Rows[i].Cells[3].Value.ToString() == "Boolean")
                        {
                            objStr = objStr + "(" + dataGridView1.Rows[i].Cells[0].Value.ToString() + " of it " + dataGridView1.Rows[i].Cells[1].Value.ToString().ToLower() + " " + dataGridView1.Rows[i].Cells[2].Value.ToString() + ")";
                        }
                        else if (dataGridView1.Rows[i].Cells[3].Value.ToString() == "Integer")
                        {
                            objStr = objStr + "(" + dataGridView1.Rows[i].Cells[0].Value.ToString() + " of it " + dataGridView1.Rows[i].Cells[1].Value.ToString().ToLower() + " " + dataGridView1.Rows[i].Cells[2].Value.ToString() + ")";
                        }
                        else if (dataGridView1.Rows[i].Cells[3].Value.ToString() == "Date")
                        {
                            objStr = objStr + "(" + dataGridView1.Rows[i].Cells[0].Value.ToString() + " of it " + dataGridView1.Rows[i].Cells[1].Value.ToString().ToLower() + " \"" + dataGridView1.Rows[i].Cells[2].Value.ToString() + "\" as date)";
                        }
                        else if (dataGridView1.Rows[i].Cells[3].Value.ToString() == "Time")
                        {
                            objStr = objStr + "(" + dataGridView1.Rows[i].Cells[0].Value.ToString() + " of it " + dataGridView1.Rows[i].Cells[1].Value.ToString().ToLower() + " \"" + dataGridView1.Rows[i].Cells[2].Value.ToString() + "\" as time)";
                        }

                        if (i != dataGridView1.Rows.Count - 1)
                        {
                            // determine if multiple filters should be ANDed or ORed
                            // objStr = objStr + " AND ";
                            objStr = objStr + " " + comboBoxANDsORs.Text + "\n\t\t\t";
                        }
                    }
                    objStr = objStr + ")";
                }

                if (obj.ToLower() == "\n\tbes unmanagedasset fields")
                {
                    objStr = objStr + " of \n\tbes unmanagedassets";
                }

                // Add property list of query
                String PropStr = "(";
                FixletProperty fp;
                for (int i = 0; i < listBoxPropertiesSelected.Items.Count; i++)
                {
                    fp = (FixletProperty)listBoxPropertiesSelected.Items[i];
                    // MessageBox.Show("Name: " + fp.Name + "\r\n" + "Display Name: " + fp.DisplayName + "\r\n" + "Parent: " + fp.ParentName);

                    // PropStr = PropStr + fp.DisplayName + " of it";
                    if (fp.DataType == "Date")
                    {
                        PropStr = PropStr + "\n\t(if (exists " + fp.DisplayName + " of it) \n\t\tthen (" + fp.DisplayName + " of it as string) \n\t\telse (\"" + "Fri, 15 Feb 1980" + "\"))";
                    }
                    else if (fp.DataType == "Time")
                    {
                        PropStr = PropStr + "\n\t(if (exists " + fp.DisplayName + " of it) \n\t\tthen (" + fp.DisplayName + " of it as string) \n\t\telse (\"" + "Fri, 15 Feb 1980 00:00:00 -0000" + "\"))";
                    }
                    else
                    {
                        if (checkBoxConcatenation.Checked)
                        {
                            // If the Attribute is from an object, we need to check for object existance first
                            if (fp.ParentName == "" || fp.ParentName == "Common Properties" || fp.ParentName == "Extended Properties")
                                PropStr = PropStr + "\n\t(if (exists " + fp.DisplayName + " of it | false) \n\t\tthen (concatenations \"" + textBoxConcatenationSeparator.Text + "\" of (" + fp.DisplayName + " of it as string)) \n\t\telse (\"" + textBoxNull.Text + "\"))";
                            else
                                PropStr = PropStr + "\n\t(if (exists " + fp.ParentName + " of it and exists " + fp.DisplayName + " of it | false) \n\t\tthen (concatenations \"" + textBoxConcatenationSeparator.Text + "\" of (" + fp.DisplayName + " of it as string)) \n\t\telse (\"" + textBoxNull.Text + "\"))";
                        }
                        else
                        {
                            if (fp.ParentName == "" || fp.ParentName == "Common Properties" || fp.ParentName == "Extended Properties")
                                PropStr = PropStr + "\n\t(if (exists " + fp.DisplayName + " of it | false) \n\t\tthen (" + fp.DisplayName + " of it as string) \n\t\telse (\"" + textBoxNull.Text + "\"))";
                            else
                                PropStr = PropStr + "\n\t(if (exists " + fp.ParentName + " of it and exists " + fp.DisplayName + " of it | false) \n\t\tthen (" + fp.DisplayName + " of it as string) \n\t\telse (\"" + textBoxNull.Text + "\"))";
                        }
                    }

                    if (i != listBoxPropertiesSelected.Items.Count - 1)
                        PropStr = PropStr + ", ";
                }
                PropStr = PropStr + ")";

                // Construct item clause
                String itemStr = "";
                if (listBoxPropertiesSelected.Items.Count != 1)
                {
                    itemStr = "(";
                    for (int k = 0; k < listBoxPropertiesSelected.Items.Count; k++)
                    {
                        if (((FixletProperty)(listBoxPropertiesSelected.Items[k])).DataType == "Date")
                        {
                            itemStr = itemStr + "\t(if (year of it as integer = 1980) \n\t\tthen (\"" + textBoxNull.Text + "\") \n\t\telse (\n\t\t\t year of it as string & \"/\" & \n\t\t\t(month of it as two digits) as string & \"/\" & \n\t\t\t(day_of_month of it as two digits) as string) ) \n\t\t\t\tof (item " + k + " of it as date) ";

                            // itemStr = itemStr + "(year of it as string & \"/\" & (month of it as two digits) as string & \"/\" & (day_of_month of it as two digits) as string) of (item " + k + " of it as date) ";
                        }
                        else if (((FixletProperty)(listBoxPropertiesSelected.Items[k])).DataType == "Time")
                        {
                            itemStr = itemStr + "\t(if ((year of date (local time zone) of it) as integer = 1980) \n\t\tthen (\"" + textBoxNull.Text + "\") \n\t\telse (\n\t\t\t(year of it as string & \"/\" & \n\t\t\t month of it as two digits & \"/\" & \n\t\t\t day_of_month of it as two digits) of date (local time zone) of it & \" \" & \n\t\t\t(two digit hour of it as string & \":\" & \n\t\t\t two digit minute of it as string) of time (local time zone) of it)) \n\t\t\t\tof (item " + k + " of it as time)";

                            // itemStr = itemStr + "((year of it as string & \"/\" & month of it as two digits & \"/\" & day_of_month of it as two digits) of date (local time zone) of it & \" \" & (two digit hour of it as string & \":\" & two digit minute of it as string) of time (local time zone) of it) of (item " + k + " of it as time) ";
                        }
                        else
                        {
                            itemStr = itemStr + "\titem " + k + " of it as string ";
                        }

                        if (k != listBoxPropertiesSelected.Items.Count - 1)
                            itemStr = itemStr + "& \"$x$\" & \n";
                    }
                    itemStr = itemStr + ") \nof ";
                }

                textBoxRelevance.Text = itemStr + PropStr + objStr;

                String queryString = textBoxRelevance.Text;

                ProcessRelevanceThenWriteToExcel(queryString, false);

            }
            catch (Exception ex)
            {
                processError(ex);
                buttonRun.Enabled = true;
                statusMessage(ex.Message, true, "Error");
            }
        }

        private void ExecuteFixlet()
        {
            try
            {
                buttonRun.Enabled = false;
                buttonRun.Update();

                statusMessage("Processing query on " + listBoxObjects.SelectedItem.ToString() + "...", true, "Normal");

                textBoxRelevance.Text = "";
                textBoxRelevance.Update();

                String siteStr = "";
                if (listBoxObjects.SelectedItem.ToString() == "BES Fixlets")
                    siteStr = " \nof \n\tbes fixlets \n\t\twhose (\n\t\t\t(";
                else if (listBoxObjects.SelectedItem.ToString() == "Results of BES Fixlets")
                    siteStr = " \nof \n\tresults\n\t\t xxx \n\tof bes fixlets \n\t\twhose (\n\t\t\t(";
                else if (listBoxObjects.SelectedItem.ToString() == "Results of BES Actions")
                    siteStr = " \nof \n\tresults\n\t\t xxx \n\tof bes actions ";

                String siteName = "";

                // Clear spreadsheet
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.ClearContents();
                DeleteCharts();

                // Add site clause to query
                if (listBoxObjects.SelectedItem.ToString() == "BES Fixlets" || listBoxObjects.SelectedItem.ToString() == "Results of BES Fixlets")
                {

                    for (int j = 0; j < checkedListBoxSites.CheckedItems.Count; j++)
                    {
                        siteName = checkedListBoxSites.CheckedItems[j].ToString();
                        siteName = siteName.Substring(0, siteName.LastIndexOf("(") - 1);
                        if (siteName == "Patches for Windows")
                        {
                            siteName = "Enterprise Security";
                        }

                        siteStr = siteStr + "name of site of it = \"" + siteName + "\"";
                        if (j != checkedListBoxSites.CheckedItems.Count - 1)
                            siteStr = siteStr + " OR\n\t\t\t ";
                    }
                    siteStr = siteStr + ")";
                }

                // Add whose clauses
                String whoseStr = "";
                if (dataGridView1.Rows.Count > 0)
                {
                    if (listBoxObjects.SelectedItem.ToString() == "BES Fixlets")
                    {
                        whoseStr = whoseStr + " AND \n\t\t\t(";
                    }

                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[0].Value.ToString() == "Remediated")
                        {
                            whoseStr = whoseStr + "(number of results whose (exists last became relevant of it AND exists last became nonrelevant of it AND last became relevant of it < last became nonrelevant of it) of it " + dataGridView1.Rows[i].Cells[1].Value.ToString().ToLower() + " " + dataGridView1.Rows[i].Cells[2].Value.ToString() + ")";
                        }
                        else if (dataGridView1.Rows[i].Cells[3].Value.ToString() == "String")
                        {
                            whoseStr = whoseStr + "(" + dataGridView1.Rows[i].Cells[0].Value.ToString() + " of it as string as lowercase " + dataGridView1.Rows[i].Cells[1].Value.ToString().ToLower() + " \"" + dataGridView1.Rows[i].Cells[2].Value.ToString().ToLower() + "\")";
                        }
                        else if (dataGridView1.Rows[i].Cells[3].Value.ToString() == "Boolean")
                        {
                            whoseStr = whoseStr + "(" + dataGridView1.Rows[i].Cells[0].Value.ToString() + " of it " + dataGridView1.Rows[i].Cells[1].Value.ToString().ToLower() + " " + dataGridView1.Rows[i].Cells[2].Value.ToString() + ")";
                        }
                        else if (dataGridView1.Rows[i].Cells[3].Value.ToString() == "Integer")
                        {
                            whoseStr = whoseStr + "(" + dataGridView1.Rows[i].Cells[0].Value.ToString() + " of it " + dataGridView1.Rows[i].Cells[1].Value.ToString().ToLower() + " " + dataGridView1.Rows[i].Cells[2].Value.ToString() + ")";
                        }
                        else if (dataGridView1.Rows[i].Cells[3].Value.ToString() == "Date")
                        {
                            whoseStr = whoseStr + "(" + dataGridView1.Rows[i].Cells[0].Value.ToString() + " of it " + dataGridView1.Rows[i].Cells[1].Value.ToString().ToLower() + " \"" + dataGridView1.Rows[i].Cells[2].Value.ToString() + "\" as date)";
                        }
                        else if (dataGridView1.Rows[i].Cells[3].Value.ToString() == "Time")
                        {
                            whoseStr = whoseStr + "(" + dataGridView1.Rows[i].Cells[0].Value.ToString() + " of it " + dataGridView1.Rows[i].Cells[1].Value.ToString().ToLower() + " \"" + dataGridView1.Rows[i].Cells[2].Value.ToString() + "\" as time)";
                        }


                        if (i != dataGridView1.Rows.Count - 1)
                        {
                            // determine if multiple filters should be ANDed or ORed
                            // objStr = objStr + " AND ";
                            whoseStr = whoseStr + " " + comboBoxANDsORs.Text + "\n\t\t\t";
                        }
                    }
                    whoseStr = whoseStr + "))";
                    if (listBoxObjects.SelectedItem.ToString() == "BES Fixlets")
                        siteStr = siteStr + whoseStr;
                    else if (listBoxObjects.SelectedItem.ToString() == "Results of BES Fixlets")
                        siteStr = siteStr + ")";
                }
                else
                {
                    siteStr = siteStr.Replace("xxx", "");
                    if (listBoxObjects.SelectedItem.ToString() != "Results of BES Actions")
                        siteStr = siteStr + ")";
                }

                if (listBoxObjects.SelectedItem.ToString() == "Results of BES Fixlets")
                    siteStr = siteStr.Replace("xxx", " whose ((" + whoseStr);

                if (listBoxObjects.SelectedItem.ToString() == "Results of BES Actions")
                    siteStr = siteStr.Replace("xxx", " whose ((" + whoseStr);

                // siteStr = siteStr + ")";

                // Add property list of query
                String PropStr = "(";
                FixletProperty fp;
                for (int i = 0; i < listBoxPropertiesSelected.Items.Count; i++)
                {
                    fp = (FixletProperty)listBoxPropertiesSelected.Items[i];
                    // PropStr = PropStr + fp.DisplayName + " of it";
                    if (fp.Name == "Remediated")
                    {
                        PropStr = PropStr + "\n\tnumber of results \n\t\twhose (\n\t\t\texists last became relevant of it AND \n\t\t\texists last became nonrelevant of it AND \n\t\t\tlast became relevant of it < last became nonrelevant of it) \n\t\tof it";
                    }
                    else if (fp.DataType == "Date")
                    {
                        PropStr = PropStr + "\n\t(if (exists " + fp.DisplayName + " of it) \n\t\tthen (" + fp.DisplayName + " of it as string) \n\t\telse (\"" + "Fri, 15 Feb 1980" + "\"))";
                    }
                    else if (fp.DataType == "Time")
                    {
                        PropStr = PropStr + "\n\t(if (exists " + fp.DisplayName + " of it) \n\t\tthen (" + fp.DisplayName + " of it as string) \n\t\telse (\"" + "Fri, 15 Feb 1980 00:00:00 -0000" + "\"))";
                    }
                    else
                    {
                        if (checkBoxConcatenation.Checked)
                        {
                            PropStr = PropStr + "\n\t(if (exists " + fp.DisplayName + " of it | false) \n\t\tthen (concatenations \"" + textBoxConcatenationSeparator.Text + "\" of (" + fp.DisplayName + " of it as string)) \n\t\telse (\"" + textBoxNull.Text + "\"))";
                        }
                        else
                        {
                            PropStr = PropStr + "\n\t(if (exists " + fp.DisplayName + " of it | false) \n\t\tthen (" + fp.DisplayName + " of it as string) \n\t\telse (\"" + textBoxNull.Text + "\"))";
                        }
                    }

                    if (i != listBoxPropertiesSelected.Items.Count - 1)
                        PropStr = PropStr + ", ";
                }
                PropStr = PropStr + ")";

                // Construct item clause
                String itemStr = "";
                if (listBoxPropertiesSelected.Items.Count != 1)
                {
                    itemStr = "(";
                    for (int k = 0; k < listBoxPropertiesSelected.Items.Count; k++)
                    {
                        if (((FixletProperty)(listBoxPropertiesSelected.Items[k])).DataType == "Date")
                        {
                            itemStr = itemStr + "\t(if (year of it as integer = 1980) \n\t\tthen (\"" + textBoxNull.Text + "\") \n\t\telse \n\t\t\t(year of it as string & \"/\" & \n\t\t\t(month of it as two digits) as string & \"/\" & \n\t\t\t(day_of_month of it as two digits) as string) ) \n\t\t\t\tof (item " + k + " of it as date)";

                            // itemStr = itemStr + "(year of it as string & \"/\" & (month of it as two digits) as string & \"/\" & (day_of_month of it as two digits) as string) of (item " + k + " of it as date) ";
                        }
                        else if (((FixletProperty)(listBoxPropertiesSelected.Items[k])).DataType == "Time")
                        {
                            itemStr = itemStr + "\t(if ((year of date (local time zone) of it) as integer = 1980) \n\t\tthen (\"" + textBoxNull.Text + "\") \n\t\telse (\n\t\t\t(year of it as string & \"/\" & \n\t\t\t month of it as two digits & \"/\" & \n\t\t\t day_of_month of it as two digits) of date (local time zone) of it & \" \" & \n\t\t\t(two digit hour of it as string & \":\" & \n\t\t\t two digit minute of it as string) of time (local time zone) of it)) \n\t\t\t\tof (item " + k + " of it as time)";

                            // itemStr = itemStr + "((year of it as string & \"/\" & month of it as two digits & \"/\" & day_of_month of it as two digits) of date (local time zone) of it & \" \" & (two digit hour of it as string & \":\" & two digit minute of it as string) of time (local time zone) of it) of (item " + k + " of it as time) ";
                        }
                        else
                        {
                            itemStr = itemStr + "\titem " + k + " of it as string ";
                        }

                        if (k != listBoxPropertiesSelected.Items.Count - 1)
                            itemStr = itemStr + " & \"$x$\" &\n";
                    }
                    itemStr = itemStr + ") \nof ";
                }

                textBoxRelevance.Text = itemStr + PropStr + siteStr;

                String queryString = textBoxRelevance.Text;

                ProcessRelevanceThenWriteToExcel(queryString, false);

            }
            catch (Exception ex)
            {
                processError(ex);
                buttonRun.Enabled = true;
                statusMessage(ex.Message, true, "Error");
            }
        }

        private void ExecuteComputer()
        {
            try
            {
                buttonRun.Enabled = false;
                buttonRun.Update();

                statusMessage("Processing query on " + listBoxObjects.SelectedItem.ToString() + "...", true, "Normal");

                textBoxRelevance.Text = "";
                textBoxRelevance.Update();

                // Clear spreadsheet
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.ClearContents();
                DeleteCharts();

                // Add properties clause to query
                String PropStr = "";
                String BESComputerRelevance = "";


                // Computer Group Filter
                if (dataGridViewComputerGroup.Rows.Count > 0)
                {
                    String prop = "";
                    String setOperator = "";
                    String searchValue = "";

                    if (comboBoxANDsORsForComputerGroup.Text == "OR")
                        BESComputerRelevance = "\telements of union of (\n";
                    else
                        BESComputerRelevance = "\telements of intersection of (\n";

                    for (int i = 0; i < dataGridViewComputerGroup.Rows.Count; i++)
                    {
                        prop = dataGridViewComputerGroup.Rows[i].Cells[0].Value.ToString();
                        setOperator = dataGridViewComputerGroup.Rows[i].Cells[1].Value.ToString();
                        searchValue = dataGridViewComputerGroup.Rows[i].Cells[2].Value.ToString().ToLower();

                        BESComputerRelevance = BESComputerRelevance +
                                    "\t\t\tmember sets of bes computer groups whose (name of it as lowercase " +
                                    setOperator + " \"" + searchValue + "\")";

                        if (i != dataGridViewComputerGroup.Rows.Count - 1)
                            BESComputerRelevance = BESComputerRelevance + "; \n";
                    }
                    BESComputerRelevance = BESComputerRelevance + ")";

                }

                // Property Filter
                if (dataGridViewFilters.Rows.Count > 0)
                {
                    String prop = "";
                    String setOperator = "";
                    String searchValue = "";
                    String propDataType = "";

                    if (comboBoxANDsORs.Text == "OR")
                        PropStr = " \nof (\n\telements of union of (\n";
                    else
                        PropStr = " \nof (\n\telements of intersection of (\n";

                    for (int i = 0; i < dataGridViewFilters.Rows.Count; i++)
                    {
                        prop = dataGridViewFilters.Rows[i].Cells[0].Value.ToString();
                        setOperator = dataGridViewFilters.Rows[i].Cells[1].Value.ToString();
                        searchValue = dataGridViewFilters.Rows[i].Cells[2].Value.ToString().ToLower();
                        propDataType = dataGridViewFilters.Rows[i].Cells[3].Value.ToString();

                        /* PropStr = PropStr + 
                                    "(sets of computers of results whose (value of it as lowercase " + 
                                    setOperator + " \"" + searchValue + "\")" + " of bes property \"" +
                                    prop + "\")" ;  */

                        // Fixing bug reported on Forum http://forum.bigfix.com/viewtopic.php?id=5394
                        // When a property contains multiple results, the filter does not work correctly with the existing
                        // relevance construct

                        if (propDataType == "Time" || prop.ToLower() == "last report time")
                        {
                            PropStr = PropStr +
                                        "\t\t(sets of items 0 \n\t\t\tof (computers of it, values whose (it as time " +
                                        setOperator + " \"" + searchValue + "\" as time) of it)";
                        }
                        else
                        {
                            PropStr = PropStr +
                                        "\t\t(sets of items 0 \n\t\t\tof (computers of it, values whose (it as lowercase " +
                                        setOperator + " \"" + searchValue + "\") of it)";
                        }

                        if (propDataType.EndsWith("*"))
                        {
                            // propDataType.TrimEnd('*') has something that looks like: 2147497463, 5, 1
                            string item0, item1, item2;
                            item0 = propDataType.Substring(0, propDataType.IndexOf(',')) ;
                            item1 = propDataType.Substring(propDataType.IndexOf(',')+2, propDataType.LastIndexOf(',') - propDataType.IndexOf(',')-2);
                            item2 = propDataType.Substring(propDataType.LastIndexOf(',') + 2, propDataType.LastIndexOf('*') - propDataType.LastIndexOf(',') - 2);
                            // PropStr = PropStr + " \n\t\t\t\tof results of bes property whose (name of it = \"" + prop + "\" and id of it = (" + propDataType.TrimEnd('*') + ")))";
                            PropStr = PropStr + " \n\t\t\t\tof results of bes property whose (name of it = \"" + prop + "\" and \n\t\t\t\t\t(item 0 of it = " + item0 + " and item 1 of it = " + item1 + " and item 2 of it = " + item2 + ") of id of it))";
                        }
                        else
                        {
                            PropStr = PropStr + " \n\t\t\t\tof results of bes property \"" + prop + "\")";
                        }

                        if (i != dataGridViewFilters.Rows.Count - 1)
                        {
                            PropStr = PropStr + "; \n";
                        }
                    }

                    PropStr = PropStr + "),";
                }

                // No Filters
                if (dataGridViewComputerGroup.Rows.Count == 0 && dataGridViewFilters.Rows.Count == 0)
                {
                    PropStr = "\nof (\n\tbes computers, ";
                }

                // Both Computer Group and Property Filters
                else if (dataGridViewComputerGroup.Rows.Count > 0 && dataGridViewFilters.Rows.Count > 0)
                {
                    String PropertyFilter = PropStr;
                    PropStr = "\nof (\n\telements of intersection of (\n";
                    PropStr = PropStr + "\t\t" + BESComputerRelevance.Substring(13) + ";";
                    PropStr = PropStr + "\n\t\t" + PropertyFilter.Substring(20, PropertyFilter.Length-21) + "),";
                }

                // Computer Group only Filters
                else if (dataGridViewComputerGroup.Rows.Count > 0)
                {
                    PropStr = "\nof (\n" + BESComputerRelevance + ", ";
                }

                // Property only Filters
                else if (dataGridViewFilters.Rows.Count > 0)
                {
                    // PropStr already has the right statement
                }


                int colCount = listBoxPropertiesSelected.Items.Count;
                if (colCount > maxColumnsExcel)
                {
                    colCount = maxColumnsExcel;
                }

                FixletProperty fp;
                // for (int i = 0; i < listBoxPropertiesSelected.Items.Count; i++)
                for (int i = 0; i < colCount; i++)
                {
                    fp = (FixletProperty)listBoxPropertiesSelected.Items[i];
                    if (fp.DataType.EndsWith("*"))
                    {
                        // propDataType.TrimEnd('*') has something that looks like: 2147497463, 5, 1
                        string item0, item1, item2;
                        item0 = fp.DataType.Substring(0, fp.DataType.IndexOf(','));
                        item1 = fp.DataType.Substring(fp.DataType.IndexOf(',') + 2, fp.DataType.LastIndexOf(',') - fp.DataType.IndexOf(',') - 2);
                        item2 = fp.DataType.Substring(fp.DataType.LastIndexOf(',') + 2, fp.DataType.LastIndexOf('*') - fp.DataType.LastIndexOf(',') - 2);
                        // PropStr = PropStr + "\n\tbes property whose (name of it = \"" + fp.DisplayName + "\" and id of it = (" + fp.DataType.TrimEnd('*') + "))";
                        PropStr = PropStr + "\n\tbes property whose (name of it = \"" + fp.DisplayName + "\" and \n\t\t(item 0 of it = " + item0 + " and item 1 of it = " + item1 + " and item 2 of it = " + item2 + ") of id of it)";
                    }
                    else
                    {
                        PropStr = PropStr + "\n\tbes property \"" + fp.DisplayName + "\"";
                    }

                    // if (i != listBoxPropertiesSelected.Items.Count - 1)
                    if (i != colCount - 1)
                        PropStr = PropStr + ",";
                }
                PropStr = PropStr + ")";

                // Construct attr clause
                String attrStr = "";
                attrStr = "(";
                // for (int k = 0; k < listBoxPropertiesSelected.Items.Count; k++)
                for (int k = 0; k < colCount; k++)
                {
                    if (checkBoxConcatenation.Checked)
                    {
                        // Changed for version 3.3.0.0
                        // attrStr = attrStr + "\n\t(if (result (item 0 of it, item " + (k + 1) + " of it) ) \n\t\tthen (concatenation \"" + textBoxConcatenationSeparator.Text + "\" of values of result (item 0 of it, item " + (k + 1) + " of it)) \n\t\telse (\"" + textBoxNull.Text + "\"))";
                        attrStr = attrStr + "\n\t(if (exists result (item 0 of it, item " + (k + 1) + " of it) and \n\t\t exists values of result (item 0 of it, item " + (k + 1) + " of it) ) \n\t\tthen (concatenation \"" + textBoxConcatenationSeparator.Text + "\" of values of result (item 0 of it, item " + (k + 1) + " of it)) \n\t\telse (\"" + textBoxNull.Text + "\"))";
                    }
                    else
                    {
                        // Changed for version 3.3.0.0
                        // attrStr = attrStr + "\n\t(if (result (item 0 of it, item " + (k + 1) + " of it) ) \n\t\tthen (values of result (item 0 of it, item " + (k + 1) + " of it)) \n\t\telse (\"" + textBoxNull.Text + "\"))";
                        attrStr = attrStr + "\n\t(if (exists result (item 0 of it, item " + (k + 1) + " of it) and \n\t\t exists values of result (item 0 of it, item " + (k + 1) + " of it) ) \n\t\tthen (values of result (item 0 of it, item " + (k + 1) + " of it)) \n\t\telse (\"" + textBoxNull.Text + "\"))";
                    }

                    // if (k != listBoxPropertiesSelected.Items.Count - 1)
                    if (k != colCount - 1)
                        attrStr = attrStr + ", ";
                }
                attrStr = attrStr + ") ";

                // Construct item clause
                String itemStr = "(";

                // This IF statement is needed to take care of single Attribute selection.
                // The (item x of it) construct does not work if only one item available.
                // Error would have been "The tuple index 0 is out of range"
                if (colCount == 1)
                {
                    itemStr = itemStr + "it";
                }
                else
                {
                    for (int j = 0; j < colCount; j++)
                    {

                        // ===================================================================================================================================

                        if (((FixletProperty)(listBoxPropertiesSelected.Items[j])).DataType == "Date")
                        {
                            itemStr = itemStr + "\t(if (year of it as integer = 1980) \n\t\tthen (\"" + textBoxNull.Text + "\") \n\t\telse \n\t\t\t(year of it as string & \"/\" & \n\t\t\t(month of it as two digits) as string & \"/\" & \n\t\t\t(day_of_month of it as two digits) as string) ) \n\t\t\t\tof (item " + j + " of it as date) ";
                        }
                        else if (((FixletProperty)(listBoxPropertiesSelected.Items[j])).DataType == "Time"  ||
                                 ((FixletProperty)(listBoxPropertiesSelected.Items[j])).Name.ToLower() == "last report time")
                        {
                            itemStr = itemStr + "\t(if ((year of date (local time zone) of it) as integer = 1980) \n\t\tthen (\"" + textBoxNull.Text + "\") \n\t\telse (\n\t\t\t(year of it as string & \"/\" & \n\t\t\t month of it as two digits & \"/\" & \n\t\t\t day_of_month of it as two digits) of date (local time zone) of it & \" \" & \n\t\t\t(two digit hour of it as string & \":\" & \n\t\t\t two digit minute of it as string) of time (local time zone) of it)) \n\t\t\t\tof (item " + j + " of it as time)";
                        }
                        else if (((FixletProperty)(listBoxPropertiesSelected.Items[j])).Name.ToLower() == "ram" ||
                                 ((FixletProperty)(listBoxPropertiesSelected.Items[j])).Name.ToLower() == "free space on system drive" ||
                                 ((FixletProperty)(listBoxPropertiesSelected.Items[j])).Name.ToLower() == "total size of system drive" ||
                                 ((FixletProperty)(listBoxPropertiesSelected.Items[j])).Name.ToLower() == "ram - unix" )
                        {
                            itemStr = itemStr + "\t(if (item " + j + " of it ends with \" MB\")\n\t\tthen (preceding text of first \" MB\" of item " + j + " of it)\n\t\telse (item " + j + " of it)) ";
                        }
                        else
                        {
                            itemStr = itemStr + "\titem " + j + " of it as string ";
                        }

                        // ===================================================================================================================================

                        if (j != colCount - 1)
                            itemStr = itemStr + " & \"$x$\" & \n";
                    }
                }

                itemStr = itemStr + ") \nof ";

                textBoxRelevance.Text = itemStr + attrStr + PropStr;

                String queryString = textBoxRelevance.Text;

                ProcessRelevanceThenWriteToExcel(queryString, false);

            }
            catch (Exception ex)
            {
                processError(ex);
                buttonRun.Enabled = true;
                statusMessage(ex.Message, true, "Error");
            }

        }

        public void ProcessRelevanceThenWriteToExcel(String relevanceStatement, Boolean refreshQuery)
        {
            try
            {
                if (checkBoxTimeout.Checked == true)
                {
                    bes.Timeout = Convert.ToInt32(numericUpDownTimeOut.Value) * 1000;
                }
                else
                {
                    bes.Timeout = Timeout.Infinite;
                }

                HiResTimer hrt = new HiResTimer();

                hrt.Start();

                results = bes.GetRelevanceResult(relevanceStatement, userName, password);

                hrt.Stop();

                t = TimeSpan.FromMilliseconds(hrt.ElapsedMicroseconds / 1000);
                QueryTime = t.ToString().Remove(t.ToString().Length - 4);
                toolStripStatusLabelEvalTime.Text = "Query: " + QueryTime;

                if (results.Length == 0)
                {
                    statusMessage("No rows returned", false, "Success");
                    buttonRun.Enabled = true;
                    // return;
                }
                else
                {
                    statusMessage(results.Length + " rows returned, writing to Excel...", false, "Normal");
                }

                if (checkBoxSortResults.Checked == true)
                {
                    Array.Sort(results);
                }

                // Excel interface in the following method
                // Parameter False means not executed from a Refresh
                ProcessExcel(refreshQuery);

            }
            catch (Exception ex)
            {
                processError(ex);
                buttonRun.Enabled = true;
                statusMessage(ex.Message, true, "Error");
            }
        }

        public void RunRefresh()
        {
            MessageBox.Show("This is supposed to do the refresh");
        }

        private void buttonAddFilter_Click(object sender, EventArgs e)
        {
            if (comboBoxProperty.SelectedItem == null || 
                comboBoxOperator.SelectedItem == null || 
                (textBoxValue.Visible == true && textBoxValue.Text == "") ||
                (comboBoxBoolean.Visible == true && comboBoxBoolean.SelectedItem == null) ||
                (comboBoxComputerGroup.Visible == true && comboBoxComputerGroup.SelectedItem == null && comboBoxComputerGroup.Text == "")
               )
            {
                statusMessage("Please complete all filter statements", true, "Warning");
            }
            else
            {
                Boolean duplicate = false;

                String activeValue = "";

                if (comboBoxBoolean.Visible == true)
                {
                    activeValue = comboBoxBoolean.SelectedItem.ToString();
                }
                else if (comboBoxComputerGroup.Visible == true)
                {
                    if (comboBoxComputerGroup.SelectedItem != null)
                        activeValue = comboBoxComputerGroup.SelectedItem.ToString();
                    else
                        activeValue = comboBoxComputerGroup.Text;
                }
                else if (dateTimePicker1.Visible == true && dateTimePicker2.Visible == false) // Date
                {
                    activeValue = dateTimePicker1.Value.ToString("ddd, dd MMM yyyy");
                }
                else if (dateTimePicker1.Visible == true && dateTimePicker2.Visible == true) // Time
                {
                    // Constructing Time in this format: "Sat, 04 Jul 2009 00:00:00 -0700"
                    String tz = dateTimePicker2.Value.ToString("zzz");
                    tz = tz.Replace(":", "");
                    activeValue = dateTimePicker1.Value.ToString("ddd, dd MMM yyyy") + " " + dateTimePicker2.Value.ToString("HH:mm:ss") + " " + tz;
                }
                else if (labelDataType.Text == "Integer")
                {
                    string Str = textBoxValue.Text.Trim();
                    double Num;
                    bool isNum = double.TryParse(Str, out Num);
                    if (isNum)
                    {
                        activeValue = Str;
                    }
                    else
                    {
                        statusMessage("Value not an Integer", true, "Warning");
                        return;
                    }
                }
                else
                {
                    activeValue = textBoxValue.Text;
                }

                // Check to see if the filter is already added, if so ignore
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value.ToString() == ((TreeNode)comboBoxProperty.SelectedItem).Text &&
                        dataGridView1.Rows[i].Cells[1].Value.ToString() == comboBoxOperator.SelectedItem.ToString() &&
                        dataGridView1.Rows[i].Cells[2].Value.ToString() == activeValue)
                    {
                        duplicate = true;
                        break;
                    }
                        
                }

                if (duplicate == false)
                {
                    dataGridView1.Tag = listBoxObjects.SelectedItem.ToString();

                    /*
                    int newRowIndex = dataGridView1.Rows.Add(((TreeNode)comboBoxProperty.SelectedItem).Text, comboBoxOperator.SelectedItem.ToString(), activeValue, ((TreeNode)comboBoxProperty.SelectedItem).Tag.ToString());
                    // Computer Groups are handled diffrently. Saving the Computer Group filters into another different grid/table for easier retrieval later
                    if (((TreeNode)comboBoxProperty.SelectedItem).Text == "Computer Groups")
                    {
                        // dataGridViewComputerGroup.Rows.Add(((TreeNode)comboBoxProperty.SelectedItem).Text, comboBoxOperator.SelectedItem.ToString(), activeValue, ((TreeNode)comboBoxProperty.SelectedItem).Tag.ToString());
                        dataGridViewComputerGroup.Rows.Add(((TreeNode)comboBoxProperty.SelectedItem).Text, comboBoxOperator.SelectedItem.ToString(), activeValue, ((TreeNode)comboBoxProperty.SelectedItem).Tag.ToString());
                        DataGridViewCellStyle specialColor = dataGridView1.DefaultCellStyle.Clone();
                        // specialColor.BackColor = Color.LightGreen;
                        // specialColor.BackColor = System.Drawing.Color.FromArgb(199, 215, 166);
                        specialColor.BackColor = System.Drawing.Color.FromArgb(231, 243, 241);
                        dataGridView1.Rows[newRowIndex].DefaultCellStyle = specialColor;
                    }
                    else
                        dataGridViewFilters.Rows.Add(((TreeNode)comboBoxProperty.SelectedItem).Text, comboBoxOperator.SelectedItem.ToString(), activeValue, ((TreeNode)comboBoxProperty.SelectedItem).Tag.ToString());
                    */

                    int newRowIndex;

                    if (comboBoxProperty.Text == "Computer Groups")
                    {
                        newRowIndex = dataGridView1.Rows.Add(comboBoxProperty.Text, comboBoxOperator.SelectedItem.ToString(), activeValue, "Computer Groups");
                        dataGridViewComputerGroup.Rows.Add(comboBoxProperty.Text, comboBoxOperator.SelectedItem.ToString(), activeValue, "Computer Groups");
                        DataGridViewCellStyle specialColor = dataGridView1.DefaultCellStyle.Clone();
                        specialColor.BackColor = System.Drawing.Color.FromArgb(231, 243, 241); // Very light blue
                        dataGridView1.Rows[newRowIndex].DefaultCellStyle = specialColor;
                    }
                    else if (comboBoxProperty.Text == "Last Report Time")
                    {
                        dataGridView1.Rows.Add(comboBoxProperty.Text, comboBoxOperator.SelectedItem.ToString(), activeValue, "Time");
                        dataGridViewFilters.Rows.Add(comboBoxProperty.Text, comboBoxOperator.SelectedItem.ToString(), activeValue, "Time");
                    }
                    else
                    {
                        dataGridView1.Rows.Add(comboBoxProperty.Text, comboBoxOperator.SelectedItem.ToString(), activeValue, ((TreeNode)comboBoxProperty.SelectedItem).Tag.ToString());
                        dataGridViewFilters.Rows.Add(comboBoxProperty.Text, comboBoxOperator.SelectedItem.ToString(), activeValue, ((TreeNode)comboBoxProperty.SelectedItem).Tag.ToString());
                    }

                    dataGridView1.ClearSelection();
                }
                else
                {
                    statusMessage("Filter already added", true, "Warning");
                }
            }
        }

        private void buttonClearFilters_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridViewFilters.Rows.Clear();
            dataGridViewComputerGroup.Rows.Clear();
        }

        private void buttonClearSelectedFilter_Click_1(object sender, EventArgs e)
        {
            if (this.dataGridView1.SelectedRows.Count > 0)
            {
                // This for loop synchronizes the Computer Group grid with the main grid. Basically deleting the Computer Group grid if the user deletes the filter criteria
                for (int i = 0; i < dataGridViewComputerGroup.Rows.Count; i++)
                {
                    if (dataGridView1.SelectedRows[0].Cells[0].Value.ToString() == dataGridViewComputerGroup.Rows[i].Cells[0].Value.ToString() &&
                        dataGridView1.SelectedRows[0].Cells[1].Value.ToString() == dataGridViewComputerGroup.Rows[i].Cells[1].Value.ToString() &&
                        dataGridView1.SelectedRows[0].Cells[2].Value.ToString() == dataGridViewComputerGroup.Rows[i].Cells[2].Value.ToString() )
                    {
                        // dataGridViewComputerGroup.Rows.RemoveAt(dataGridViewComputerGroup.Rows[i].Index);
                        dataGridViewComputerGroup.Rows.Remove(dataGridViewComputerGroup.Rows[i]);
                        break;
                    }
                }

                for (int i = 0; i < dataGridViewFilters.Rows.Count; i++)
                {
                    if (dataGridView1.SelectedRows[0].Cells[0].Value.ToString() == dataGridViewFilters.Rows[i].Cells[0].Value.ToString() &&
                        dataGridView1.SelectedRows[0].Cells[1].Value.ToString() == dataGridViewFilters.Rows[i].Cells[1].Value.ToString() &&
                        dataGridView1.SelectedRows[0].Cells[2].Value.ToString() == dataGridViewFilters.Rows[i].Cells[2].Value.ToString())
                    {
                        // dataGridViewComputerGroup.Rows.RemoveAt(dataGridViewComputerGroup.Rows[i].Index);
                        dataGridViewFilters.Rows.Remove(dataGridViewFilters.Rows[i]);
                        break;
                    }
                }

                dataGridView1.Rows.RemoveAt(this.dataGridView1.SelectedRows[0].Index);
            }
            else
            {
                statusMessage("Select a filter to be removed", true, "Warning");
            }
        }

        private void comboBoxProperty_DropDownClosed(object sender, EventArgs e)
        {
            // This code is used to get around a bug: http://connect.microsoft.com/VisualStudio/feedback/details/615543/combobox-with-autocomplete-mode-suggest-has-a-problem
            // To reproduce:
            // - In the Add Filter screen, click Select Property combobox drop down
            // - Type a property, such as "OS"
            // - Immediately tab to next field
            // - The combobox.selectedItem is actually null, although an item seems selected
            if (comboBoxProperty.SelectedItem == null && comboBoxProperty.Text != "")
            {
                for (int i = 0; i < comboBoxProperty.Items.Count; i++)
                {
                    if (((TreeNode)comboBoxProperty.Items[i]).Text == comboBoxProperty.Text)
                    {
                        comboBoxProperty.SelectedItem = comboBoxProperty.Items[i];
                    }
                }
            }
        }

        private void comboBoxProperty_SelectedIndexChanged(object sender, EventArgs e)
        {
            statusMessage("", true, "Normal");

            labelDataType.Text = ((TreeNode)comboBoxProperty.SelectedItem).Tag.ToString();

            if (labelDataType.Text == "String")
            {
                textBoxValue.Visible = true;
                comboBoxBoolean.Visible = false;
                comboBoxComputerGroup.Visible = false; 
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
                comboBoxOperator.Items.Clear();
                comboBoxOperator.Items.Add("=");
                comboBoxOperator.Items.Add("!=");
                comboBoxOperator.Items.Add("contains");
                comboBoxOperator.Items.Add("does not contain");
                comboBoxOperator.Items.Add("starts with");
                comboBoxOperator.Items.Add("ends with");
                comboBoxOperator.Items.Add(">");
                comboBoxOperator.Items.Add(">=");
                comboBoxOperator.Items.Add("<");
                comboBoxOperator.Items.Add("<=");
                textBoxValue.Text = "";
            }
            else if (labelDataType.Text == "Boolean")
            {
                textBoxValue.Visible = false;
                comboBoxBoolean.Visible = true;
                comboBoxComputerGroup.Visible = false;
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
                comboBoxOperator.Items.Clear();
                comboBoxOperator.Items.Add("=");
                comboBoxOperator.Items.Add("!=");
                comboBoxBoolean.SelectedValue = "";
            }
            else if (labelDataType.Text == "Computer Groups")
            {
                textBoxValue.Visible = false;
                comboBoxBoolean.Visible = false;
                comboBoxComputerGroup.Visible = true;
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
                comboBoxOperator.Items.Clear();
                comboBoxOperator.Items.Add("=");
                comboBoxOperator.Items.Add("!=");
                comboBoxOperator.Items.Add("contains");
                comboBoxOperator.Items.Add("does not contain");
                comboBoxOperator.Items.Add("starts with");
                comboBoxOperator.Items.Add("ends with");
                comboBoxOperator.Items.Add(">");
                comboBoxOperator.Items.Add(">=");
                comboBoxOperator.Items.Add("<");
                comboBoxOperator.Items.Add("<=");
                comboBoxComputerGroup.SelectedValue = "";

                if (resultsComputerGroups == null)
                {
                    String computerGroupRelevance = "unique values of names of bes computer groups";
                    resultsComputerGroups = bes.GetRelevanceResult(computerGroupRelevance, userName, password);
                    for (int i = 0; i < resultsComputerGroups.Length; i++)
                    {
                        comboBoxComputerGroup.Items.Add(resultsComputerGroups[i]);
                    }
                }


            }
            else if (labelDataType.Text == "Integer")
            {
                textBoxValue.Visible = true;
                comboBoxBoolean.Visible = false;
                comboBoxComputerGroup.Visible = false;
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false;
                comboBoxOperator.Items.Clear();
                comboBoxOperator.Items.Add("=");
                comboBoxOperator.Items.Add("!=");
                comboBoxOperator.Items.Add(">");
                comboBoxOperator.Items.Add(">=");
                comboBoxOperator.Items.Add("<");
                comboBoxOperator.Items.Add("<=");
                textBoxValue.Text = "";
            }
            else if (labelDataType.Text == "Date")
            {
                textBoxValue.Visible = false;
                comboBoxBoolean.Visible = false;
                comboBoxComputerGroup.Visible = false;
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = false; 
                comboBoxOperator.Items.Clear();
                comboBoxOperator.Items.Add("=");
                comboBoxOperator.Items.Add("!=");
                comboBoxOperator.Items.Add(">");
                comboBoxOperator.Items.Add(">=");
                comboBoxOperator.Items.Add("<");
                comboBoxOperator.Items.Add("<=");
            }
            else if (labelDataType.Text == "Time" ||
                    ((TreeNode)comboBoxProperty.SelectedItem).Text.ToLower() == "last report time" )
            {
                textBoxValue.Visible = false;
                comboBoxBoolean.Visible = false;
                comboBoxComputerGroup.Visible = false;
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = true;
                comboBoxOperator.Items.Clear();
                comboBoxOperator.Items.Add("=");
                comboBoxOperator.Items.Add("!=");
                comboBoxOperator.Items.Add(">");
                comboBoxOperator.Items.Add(">=");
                comboBoxOperator.Items.Add("<");
                comboBoxOperator.Items.Add("<=");
            }
            else
            {
                textBoxValue.Visible = true;
                comboBoxBoolean.Visible = false;
                comboBoxComputerGroup.Visible = false;
                dateTimePicker1.Visible = false;
                dateTimePicker2.Visible = false; 
                comboBoxOperator.Items.Clear();
                comboBoxOperator.Items.Add("=");
                comboBoxOperator.Items.Add("!=");
                comboBoxOperator.Items.Add("contains");
                comboBoxOperator.Items.Add("does not contain");
                comboBoxOperator.Items.Add("starts with");
                comboBoxOperator.Items.Add("ends with");
                comboBoxOperator.Items.Add(">");
                comboBoxOperator.Items.Add(">=");
                comboBoxOperator.Items.Add("<");
                comboBoxOperator.Items.Add("<=");
                textBoxValue.Text = "";
            }

        }

        private void comboBoxOperator_SelectedIndexChanged(object sender, EventArgs e)
        {
            statusMessage("", true, "Normal");
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridViewComputerGroup.Rows.Clear();
            dataGridViewFilters.Rows.Clear();

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString() == "Computer Groups")
                {
                    dataGridViewComputerGroup.Rows.Add(
                        dataGridView1.Rows[i].Cells[0].Value.ToString(),
                        dataGridView1.Rows[i].Cells[1].Value.ToString(),
                        dataGridView1.Rows[i].Cells[2].Value.ToString(),
                        dataGridView1.Rows[i].Cells[3].Value.ToString());
                }
                else
                {
                    dataGridViewFilters.Rows.Add(
                        dataGridView1.Rows[i].Cells[0].Value.ToString(),
                        dataGridView1.Rows[i].Cells[1].Value.ToString(),
                        dataGridView1.Rows[i].Cells[2].Value.ToString(),
                        dataGridView1.Rows[i].Cells[3].Value.ToString());
                }
            }

            dataGridView1.ClearSelection();
        }

        private string ExcelColumnLetter(int intCol)
        {
            int intFirstLetter = ((intCol) / 26) + 64;
            int intSecondLetter = (intCol % 26) + 65;
            char letter1 = (intFirstLetter > 64) ? (char)intFirstLetter : ' ';
            return string.Concat(letter1, (char)intSecondLetter).Trim();
        }

        private void numericUpDownOpacity_ValueChanged(object sender, EventArgs e)
        {
            this.Opacity = Convert.ToDouble(numericUpDownOpacity.Value / 100);
        }

        #endregion

        #region Excel Save and Restore

        private String LoadFromExcel(String loc)
        {
            Excel.Worksheet storageWorksheet;
            try
            {
                storageWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                Excel.Range rangeStorage;
                rangeStorage = storageWorksheet.get_Range(loc, loc);
                return rangeStorage.Value.ToString();
            }
            catch (Exception ex)
            {
                String nullMsg = ex.Message;
                return "";
            }
        }

        private String LoadFromActiveSheet(String loc)
        {
            try
            {
                Excel.Range rangeStorage;
                rangeStorage = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(loc, loc);
                return rangeStorage.Value.ToString();
            }
            catch (Exception ex)
            {
                String nullMsg = ex.Message;
                return "";
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

        private void saveToExcel()
        {
            // Save the query info into a hidden worksheet 
            Excel.Worksheet storageWorksheet;

            try
            {
                // First create the hidden worksheet name "BigFixExcelConnector" if it does not already exist
                CreateHiddenExcelWorksheet();

                storageWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                // A1 is the Relevance statement
                Excel.Range rangeStorage;
                rangeStorage = storageWorksheet.get_Range("A1", "A1");
                rangeStorage.Value = textBoxRelevance.Text;

                // A2 is the selected Object, e.g. BES Computers
                rangeStorage = storageWorksheet.get_Range("A2", "A2");
                rangeStorage.Value = listBoxObjects.SelectedItem.ToString();

                // A3 are the selected Fixlet sites
                String selectedFixletSites = "";
                String singleSite = "";
                for (int j = 0; j < checkedListBoxSites.CheckedItems.Count; j++)
                {
                    singleSite = checkedListBoxSites.CheckedItems[j].ToString();
                    singleSite = singleSite.Substring(0, singleSite.LastIndexOf("(") - 1);
                    selectedFixletSites = selectedFixletSites + singleSite + "|";
                }
                rangeStorage = storageWorksheet.get_Range("A3", "A3");
                rangeStorage.Value = selectedFixletSites;

                // A4 are the selected attributes
                String selectedAttributes = "";
                String selectedAttributesForRefresh = ""; // Save this for Refresh Query purpose
                String selectedAttributesForRefreshDataType = "";

                for (int k = 0; k < listBoxPropertiesSelected.Items.Count; k++)
                {
                    selectedAttributes = selectedAttributes + ((FixletProperty)listBoxPropertiesSelected.Items[k]).ParentName + "!!" + ((FixletProperty)listBoxPropertiesSelected.Items[k]).Name + "|";
                    selectedAttributesForRefresh = selectedAttributesForRefresh + ((FixletProperty)listBoxPropertiesSelected.Items[k]).Name;
                    selectedAttributesForRefreshDataType = ((FixletProperty)listBoxPropertiesSelected.Items[k]).DataType;

                    if ( ((FixletProperty)listBoxPropertiesSelected.Items[k]).ParentName == "" ||
                         ((FixletProperty)listBoxPropertiesSelected.Items[k]).ParentName == "Common Properties" ||
                         ((FixletProperty)listBoxPropertiesSelected.Items[k]).ParentName == "Extended Properties")
                    {
                        selectedAttributesForRefresh = selectedAttributesForRefresh + "^" + selectedAttributesForRefreshDataType + "|";
                    }
                    else
                    {
                        selectedAttributesForRefresh = selectedAttributesForRefresh + " of " + ((FixletProperty)listBoxPropertiesSelected.Items[k]).ParentName + "^" + selectedAttributesForRefreshDataType + "|";
                    }
                }
                // TraverseTreeView(treeView1);
                rangeStorage = storageWorksheet.get_Range("A4", "A4");
                rangeStorage.Value = selectedAttributes;

                // A7 has the selected attributes used for Refresh
                rangeStorage = storageWorksheet.get_Range("A7", "A7");
                rangeStorage.Value = selectedAttributesForRefresh;

                // A5 has the Filter statements
                String filterStatements = "";
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    filterStatements = filterStatements +
                        dataGridView1.Rows[i].Cells[0].Value + "|x1|" +
                        dataGridView1.Rows[i].Cells[1].Value + "|x2|" +
                        dataGridView1.Rows[i].Cells[2].Value + "|x3|" +
                        dataGridView1.Rows[i].Cells[3].Value + "|x|";
                }
                rangeStorage = storageWorksheet.get_Range("A5", "A5");
                rangeStorage.Value = filterStatements;

                // The following values are saved for Refresh Query

                // A6 has the number of properties selected
                rangeStorage = storageWorksheet.get_Range("A6", "A6");
                rangeStorage.Value = listBoxPropertiesSelected.Items.Count.ToString();

                // A8 - CheckAutofit
                // A9 - CheckConcat
                // A10 - CheckSort
                // A11 - CheckTimeout
                // A12 - AutofitRowHeightMax
                // A13 - ConcatenationSeparator
                // A14 - NullSubstitution
                // A15 - TimeoutSecs
                // A16 - TimeSpan
                // A17 - Name of Report
                // A18 - FilterAndOr
                // A19 - FilterAndOrForComputerGroup
                
                rangeStorage = storageWorksheet.get_Range("A8", "A8");
                rangeStorage.Value = checkBoxRowHeightAutoFit.Checked.ToString();
                rangeStorage = storageWorksheet.get_Range("A9", "A9");
                rangeStorage.Value = checkBoxConcatenation.Checked.ToString();
                rangeStorage = storageWorksheet.get_Range("A10", "A10");
                rangeStorage.Value = checkBoxSortResults.Checked.ToString();
                rangeStorage = storageWorksheet.get_Range("A11", "A11");
                rangeStorage.Value = checkBoxTimeout.Checked.ToString();
                rangeStorage = storageWorksheet.get_Range("A12", "A12");
                rangeStorage.Value = numericUpDownRowHeightMaximum.Value.ToString();
                rangeStorage = storageWorksheet.get_Range("A13", "A13");
                rangeStorage.Value = textBoxConcatenationSeparator.Text;
                rangeStorage = storageWorksheet.get_Range("A14", "A14");
                rangeStorage.Value = textBoxNull.Text;
                rangeStorage = storageWorksheet.get_Range("A15", "A15");
                rangeStorage.Value = numericUpDownTimeOut.Value.ToString();
                rangeStorage = storageWorksheet.get_Range("A18", "A18");
                rangeStorage.Value = comboBoxANDsORs.Text;
                rangeStorage = storageWorksheet.get_Range("A19", "A19");
                rangeStorage.Value = comboBoxANDsORsForComputerGroup.Text;

                // rangeStorage = storageWorksheet.get_Range("A16", "A16");
                // rangeStorage.Value = t2.ToString().Remove(t2.ToString().Length - 4) + " / " + t.ToString().Remove(t.ToString().Length - 4);

                NeedToRefreshWizard = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error saving settings to hidden worksheet", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void RestoreQueryWizardPreferences()
        {
            try
            {
                Excel.Worksheet hiddenWorksheet;
                hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");

                checkBoxRowHeightAutoFit.Checked = Convert.ToBoolean(hiddenWorksheet.get_Range("A8", "A8").Value.ToString());
                checkBoxConcatenation.Checked = Convert.ToBoolean(hiddenWorksheet.get_Range("A9", "A9").Value.ToString());
                checkBoxSortResults.Checked = Convert.ToBoolean(hiddenWorksheet.get_Range("A10", "A10").Value.ToString());
                checkBoxTimeout.Checked = Convert.ToBoolean(hiddenWorksheet.get_Range("A11", "A11").Value.ToString());

                numericUpDownRowHeightMaximum.Value = Convert.ToInt32(hiddenWorksheet.get_Range("A12", "A12").Value.ToString());
                textBoxConcatenationSeparator.Text = hiddenWorksheet.get_Range("A13", "A13").Value.ToString();
                textBoxNull.Text = hiddenWorksheet.get_Range("A14", "A14").Value.ToString();
                numericUpDownTimeOut.Value = Convert.ToInt32(hiddenWorksheet.get_Range("A15", "A15").Value.ToString());

                textBoxRelevance.Text = hiddenWorksheet.get_Range("A1", "A1").Value.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem while restoring saved preferences from spreadsheet to do refresh.\r\n" + ex.Message , "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void loadFiltersFromExcel(String obj)
        {

            Excel.Worksheet storageWorksheet;
            try
            {
                storageWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                Excel.Range rangeStorage;
                rangeStorage = storageWorksheet.get_Range("A2", "A2");

                if (rangeStorage.Value.ToString() != obj)
                   return;

                String[] previouslySelectedFilters;
                String[] delimiters = new String[] { "|x|" };

                rangeStorage = storageWorksheet.get_Range("A5", "A5");
                previouslySelectedFilters = rangeStorage.Value.ToString().Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                foreach (String s in previouslySelectedFilters)
                {
                    int newRowIndex = dataGridView1.Rows.Add(s.Substring(0, s.IndexOf("|x1|")),
                                    s.Substring(s.IndexOf("|x1|") + 4, s.IndexOf("|x2") - s.IndexOf("|x1|") - 4),
                                    s.Substring(s.IndexOf("|x2|") + 4, s.IndexOf("|x3") - s.IndexOf("|x2|") - 4),
                                    s.Substring(s.IndexOf("|x3|") + 4));

                    if (s.Substring(0, s.IndexOf("|x1|")) == "Computer Groups")
                    {
                        dataGridViewComputerGroup.Rows.Add(s.Substring(0, s.IndexOf("|x1|")),
                                        s.Substring(s.IndexOf("|x1|") + 4, s.IndexOf("|x2") - s.IndexOf("|x1|") - 4),
                                        s.Substring(s.IndexOf("|x2|") + 4, s.IndexOf("|x3") - s.IndexOf("|x2|") - 4),
                                        s.Substring(s.IndexOf("|x3|") + 4));
                        DataGridViewCellStyle specialColor = dataGridView1.DefaultCellStyle.Clone();
                        // specialColor.BackColor = Color.LightGreen;
                        // specialColor.BackColor = System.Drawing.Color.FromArgb(199, 215, 166);
                        specialColor.BackColor = System.Drawing.Color.FromArgb(231, 243, 241);
                        dataGridView1.Rows[newRowIndex].DefaultCellStyle = specialColor;
                    }
                    else
                    {
                        dataGridViewFilters.Rows.Add(s.Substring(0, s.IndexOf("|x1|")),
                                        s.Substring(s.IndexOf("|x1|") + 4, s.IndexOf("|x2") - s.IndexOf("|x1|") - 4),
                                        s.Substring(s.IndexOf("|x2|") + 4, s.IndexOf("|x3") - s.IndexOf("|x2|") - 4),
                                        s.Substring(s.IndexOf("|x3|") + 4));
                    }
                }

                // Restore the Filters Are AND/OR combobox
                rangeStorage = storageWorksheet.get_Range("A18", "A18");
                comboBoxANDsORs.SelectedValue = rangeStorage.Value.ToString();
                comboBoxANDsORs.Text = rangeStorage.Value.ToString();

                rangeStorage = storageWorksheet.get_Range("A19", "A19");
                comboBoxANDsORs.SelectedValue = rangeStorage.Value.ToString();
                comboBoxANDsORsForComputerGroup.Text = rangeStorage.Value.ToString();

                dataGridView1.ClearSelection();
                comboBoxProperty.Select();

            }
            catch (Exception ex)
            {
                String nullMsg = ex.Message;
            }

        }

        private void loadFixletSitesFromExcel()
        {
            Excel.Worksheet storageWorksheet;
            try
            {
                storageWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                Excel.Range rangeStorage;
                // rangeStorage = storageWorksheet.get_Range("A2", "A2");

                String[] previouslySelectedFixletSites;
                char[] delimiters = new char[] { '|' };

                rangeStorage = storageWorksheet.get_Range("A3", "A3");

                previouslySelectedFixletSites = rangeStorage.Value.ToString().Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                String oneSite = "";
                foreach (String s in previouslySelectedFixletSites)
                {
                    for (int i = 0; i < checkedListBoxSites.Items.Count; i++)
                    {
                        oneSite = checkedListBoxSites.Items[i].ToString();
                        if (oneSite.Substring(0, oneSite.LastIndexOf("(") - 1) == s)
                        {
                            checkedListBoxSites.SetItemChecked(i, true);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                String nullMsg = ex.Message;
            }

        }

        private void comboBoxProperty_KeyPress(object sender, KeyPressEventArgs e)
        {
            this.AutoComplete(comboBoxProperty, e, true);
        }

        #endregion

        #region Common

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

        public Boolean RetrieveSettings()
        {
            try
            {
                webReportsURL = GetSettings("ExcelConnector", "WebReportsServer");
                userName = GetSettings("ExcelConnector", "Username");
                password = Decrypt(GetSettings("ExcelConnector", "Password"));
                refreshContentCache = GetSettings("ExcelConnector", "RefreshCache");

                if (refreshContentCache == null)
                {
                    refreshContentCache = "False";
                }

                if (GetSettings("ExcelConnector", "ConcatenationSeparator") != null)
                {
                    concatenationSeparator = GetSettings("ExcelConnector", "ConcatenationSeparator");
                }
                textBoxConcatenationSeparator.Text = concatenationSeparator;

                if (GetSettings("ExcelConnector", "NullSubstitution") != null)
                {
                    nullSubstitution = GetSettings("ExcelConnector", "NullSubstitution");
                }
                textBoxNull.Text = nullSubstitution;

                if (GetSettings("ExcelConnector", "AutofitRowHeightMax") != null)
                {
                    autofitRowHeightMax = GetSettings("ExcelConnector", "AutofitRowHeightMax");
                }
                numericUpDownRowHeightMaximum.Value = Convert.ToInt32(autofitRowHeightMax);

                if (GetSettings("ExcelConnector", "TimeoutSecs") != null)
                {
                    timeOutSecs = GetSettings("ExcelConnector", "TimeoutSecs");
                }
                numericUpDownTimeOut.Value = Convert.ToInt32(timeOutSecs);

                if (GetSettings("ExcelConnector", "CheckConcat") == null)
                    checkBoxConcatenation.Checked = true;
                else
                    checkBoxConcatenation.Checked = Convert.ToBoolean(GetSettings("ExcelConnector", "CheckConcat"));

                if (GetSettings("ExcelConnector", "CheckAutofit") == null)
                    checkBoxRowHeightAutoFit.Checked = true;
                else
                    checkBoxRowHeightAutoFit.Checked = Convert.ToBoolean(GetSettings("ExcelConnector", "CheckAutofit"));

                if (GetSettings("ExcelConnector", "CheckSort") == null)
                    checkBoxSortResults.Checked = true;
                else
                    checkBoxSortResults.Checked = Convert.ToBoolean(GetSettings("ExcelConnector", "CheckSort"));

                if (GetSettings("ExcelConnector", "CheckTimeout") == null)
                    checkBoxTimeout.Checked = true;
                else
                    checkBoxTimeout.Checked = Convert.ToBoolean(GetSettings("ExcelConnector", "CheckTimeout"));

                if ((webReportsURL == String.Empty || userName == String.Empty || password == String.Empty) && userStarted)
                {
                    MessageBox.Show("Please configure login information to IBM BigFix Web Reports first", "Login error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                else
                {
                    bes.Url = webReportsURL.TrimEnd('/') + "/webreports";
                    return true;
                }

            }
            catch (Exception ex)
            {
                if (userStarted)
                    MessageBox.Show(ex.Message, "Error retrieving settings", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
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

        private void statusMessage(String msg, Boolean clearTime, String level)
        {
            toolStripStatusLabelMessage.Text = msg;

            if (clearTime == true)
            {
                toolStripStatusLabelEvalTime.Text = "";
            }

            if (level.ToLower() == "normal")
                toolStripStatusLabelMessage.BackColor = System.Drawing.SystemColors.Control;
            else if (level.ToLower() == "success")
                toolStripStatusLabelMessage.BackColor = System.Drawing.Color.FromArgb(199, 215, 166);
            else if (level.ToLower() == "warning")
                toolStripStatusLabelMessage.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            else if (level.ToLower() == "error")
                toolStripStatusLabelMessage.BackColor = System.Drawing.Color.LightCoral;
            else
                toolStripStatusLabelMessage.BackColor = System.Drawing.SystemColors.Control;

            statusStrip1.Update();

        }

        public class FixletProperty
        {
            public string Name;
            public string DataType;
            public string DisplayName;
            public string ParentName;

            public FixletProperty() { }

            public FixletProperty(string name, string displayName, string dataType, string parentName)
            {
                Name = name;
                DataType = dataType;
                DisplayName = displayName;
                ParentName = parentName;
            }

            public override string ToString()
            {
                return DisplayName;
            }

            public string ToName()
            {
                return Name;
            }
        }

        private void processError(Exception ex)
        {
            if ((ex.Message.ToLower().Contains("login failed")) || (ex.Message.ToLower().Contains("remote name could not be resolved"))
                            || (ex.Message.ToLower().Contains("unable to connect to the remote server"))
                            || (ex.Message.ToLower().Contains("invalid uri"))
                            || (ex.Message.ToLower().Contains("uri prefix"))
                            || (ex.Message.ToLower().Contains("missing parameter"))
                            || (ex.Message.ToLower().Contains("connection was closed"))
                            )
            {
                MessageBox.Show(ex.Message + " - Have you configured login information?", "Not connected", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (ex.Message.ToLower().Contains("object reference not set to an instance of an object"))
            {
                MessageBox.Show("Error: " + ex.Message + " - Note that this IBM BigFix AddIn only works for BES 7.2 or above.", "Check IBM BigFix version", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (ex.Message.ToLower().Contains("singular expression refers to nonexistent object"))
            {
                MessageBox.Show("Error from IBM BigFix:\n" + ex.Message + "\n\nGetting the above error running the Relevance Query.\nIs it possible that some items may have been removed from the backend, such as a Retrieved Property since the query was defined?\n\nTry using the Query Wizard to generate a new report.",
                    "IBM BigFix Session Relevance Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show(ex.Message, "Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        // AutoComplete
        public void AutoComplete(ComboBox cb, System.Windows.Forms.KeyPressEventArgs e)
        {
            this.AutoComplete(cb, e, false);
        }

        public void AutoComplete(ComboBox cb, System.Windows.Forms.KeyPressEventArgs e, bool blnLimitToList)
        {
            string strFindStr = "";

            if (e.KeyChar == (char)8)
            {
                if (cb.SelectionStart <= 1)
                {
                    cb.Text = "";
                    return;
                }

                if (cb.SelectionLength == 0)
                    strFindStr = cb.Text.Substring(0, cb.Text.Length - 1);
                else
                    strFindStr = cb.Text.Substring(0, cb.SelectionStart - 1);
            }
            else
            {
                if (cb.SelectionLength == 0)
                    strFindStr = cb.Text + e.KeyChar;
                else
                    strFindStr = cb.Text.Substring(0, cb.SelectionStart) + e.KeyChar;
            }

            int intIdx = -1;

            // Search the string in the ComboBox list.

            intIdx = cb.FindString(strFindStr);

            if (intIdx != -1)
            {
                cb.SelectedText = "";
                cb.SelectedIndex = intIdx;
                cb.SelectionStart = strFindStr.Length;
                cb.SelectionLength = cb.Text.Length;
                e.Handled = true;
            }
            else
            {
                e.Handled = blnLimitToList;
            }

        }

        #endregion
    }

    #region NodeSorter Class
    // Create a node sorter that implements the IComparer interface.
    public class NodeSorter : IComparer
    {
        // Compare the length of the strings, or the strings
        // themselves, if they are the same length.
        public int Compare(object x, object y)
        {
            TreeNode tx = x as TreeNode;
            TreeNode ty = y as TreeNode;

            // Compare the length of the strings, returning the difference.
            // if (tx.Text.Length != ty.Text.Length)
            //    return tx.Text.Length - ty.Text.Length;

            // If they are the same length, call Compare.
            return string.Compare(tx.Text, ty.Text);
        }
    }
    #endregion
}