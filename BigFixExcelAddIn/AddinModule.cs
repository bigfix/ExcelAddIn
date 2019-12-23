using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Threading;
using Microsoft.Win32;
using System.Windows.Forms;
using AboutBoxDemo;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Collections;

namespace BigFixExcelConnector
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("FDFBA9E9-6C9D-45B8-8594-C8A035970615"), ProgId("BigFixExcelConnector.AddinModule")]
    public class AddinModule : AddinExpress.MSO.ADXAddinModule
    {
        public AddinModule()
        {
            InitializeComponent();
            // Please add any initialization code to the AddinInitialize event handler
        }

        private AddinExpress.MSO.ADXCommandBar adxCommandBarBF;
        private AddinExpress.XL.ADXExcelTaskPanesManager adxExcelTaskPanesManagerBF;
        private AddinExpress.XL.ADXExcelTaskPanesCollectionItem adxExcelTaskPanesCollectionItemBF;
        private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton1;
        private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton2;
        private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton3;
        private AddinExpress.MSO.ADXRibbonTab adxRibbonTab1;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup1;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButton1;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButton2;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButton3;
        private ImageList imageListRibbon;
        private FormWizard formWizard;
        private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton4;
        private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton5;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButton4;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButton5;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator1;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator2;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator3;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator4;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator5;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButton7;
        private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton6;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButton6;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButton8;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator6;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator7;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButton9;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButton10;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator8;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator9;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButton11;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator10;
        private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton7;
        private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton8;
        private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton9;
        private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton10;
        private AddinExpress.MSO.ADXCommandBarButton adxCommandBarButton11;
 
        #region Component Designer generated code
        /// <summary>
        /// Required by designer
        /// </summary>
        private System.ComponentModel.IContainer components;
 
        /// <summary>
        /// Required by designer support - do not modify
        /// the following method
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddinModule));
            this.adxCommandBarBF = new AddinExpress.MSO.ADXCommandBar(this.components);
            this.adxCommandBarButton7 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.imageListRibbon = new System.Windows.Forms.ImageList(this.components);
            this.adxCommandBarButton8 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.adxCommandBarButton3 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.adxCommandBarButton4 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.adxCommandBarButton5 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.adxCommandBarButton9 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.adxCommandBarButton10 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.adxCommandBarButton11 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.adxCommandBarButton2 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.adxCommandBarButton1 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.adxCommandBarButton6 = new AddinExpress.MSO.ADXCommandBarButton(this.components);
            this.adxExcelTaskPanesManagerBF = new AddinExpress.XL.ADXExcelTaskPanesManager(this.components);
            this.adxExcelTaskPanesCollectionItemBF = new AddinExpress.XL.ADXExcelTaskPanesCollectionItem(this.components);
            this.adxRibbonTab1 = new AddinExpress.MSO.ADXRibbonTab(this.components);
            this.adxRibbonGroup1 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.adxRibbonButton6 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator1 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButton8 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator2 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButton2 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator3 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButton4 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator4 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButton5 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator5 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButton9 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator6 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButton10 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator7 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButton11 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator8 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButton3 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator9 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButton1 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator10 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButton7 = new AddinExpress.MSO.ADXRibbonButton(this.components);
            // 
            // adxCommandBarBF
            // 
            this.adxCommandBarBF.CommandBarName = "IBM BigFix Connector";
            this.adxCommandBarBF.CommandBarTag = "213e1f18-2807-443f-a71f-be57720e9c4f";
            this.adxCommandBarBF.Controls.Add(this.adxCommandBarButton7);
            this.adxCommandBarBF.Controls.Add(this.adxCommandBarButton8);
            this.adxCommandBarBF.Controls.Add(this.adxCommandBarButton3);
            this.adxCommandBarBF.Controls.Add(this.adxCommandBarButton4);
            this.adxCommandBarBF.Controls.Add(this.adxCommandBarButton5);
            this.adxCommandBarBF.Controls.Add(this.adxCommandBarButton9);
            this.adxCommandBarBF.Controls.Add(this.adxCommandBarButton10);
            this.adxCommandBarBF.Controls.Add(this.adxCommandBarButton11);
            this.adxCommandBarBF.Controls.Add(this.adxCommandBarButton2);
            this.adxCommandBarBF.Controls.Add(this.adxCommandBarButton1);
            this.adxCommandBarBF.Controls.Add(this.adxCommandBarButton6);
            this.adxCommandBarBF.Description = "IBM BigFix Connector";
            this.adxCommandBarBF.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaExcel;
            this.adxCommandBarBF.UpdateCounter = 24;
            // 
            // adxCommandBarButton7
            // 
            this.adxCommandBarButton7.Caption = "Open Query";
            this.adxCommandBarButton7.ControlTag = "af9baf3c-44e1-4580-89f9-1565c86ebaf8";
            this.adxCommandBarButton7.Image = 13;
            this.adxCommandBarButton7.ImageList = this.imageListRibbon;
            this.adxCommandBarButton7.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxCommandBarButton7.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.adxCommandBarButton7.UpdateCounter = 4;
            this.adxCommandBarButton7.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton7_Click);
            // 
            // imageListRibbon
            // 
            this.imageListRibbon.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListRibbon.ImageStream")));
            this.imageListRibbon.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListRibbon.Images.SetKeyName(0, "applications.png");
            this.imageListRibbon.Images.SetKeyName(1, "document_edit.png");
            this.imageListRibbon.Images.SetKeyName(2, "paste.png");
            this.imageListRibbon.Images.SetKeyName(3, "1245774238_Startup Wizard.png");
            this.imageListRibbon.Images.SetKeyName(4, "1245774670_wizard.png");
            this.imageListRibbon.Images.SetKeyName(5, "1245774782_agt_utilities copy.png");
            this.imageListRibbon.Images.SetKeyName(6, "1245775183_news_subscribe.png");
            this.imageListRibbon.Images.SetKeyName(7, "Refresh.png");
            this.imageListRibbon.Images.SetKeyName(8, "programming.png");
            this.imageListRibbon.Images.SetKeyName(9, "wizard.png");
            this.imageListRibbon.Images.SetKeyName(10, "text_editor.png");
            this.imageListRibbon.Images.SetKeyName(11, "package_editors.png");
            this.imageListRibbon.Images.SetKeyName(12, "About.png");
            this.imageListRibbon.Images.SetKeyName(13, "Open.png");
            this.imageListRibbon.Images.SetKeyName(14, "Save.png");
            this.imageListRibbon.Images.SetKeyName(15, "trans_teaser2.png");
            this.imageListRibbon.Images.SetKeyName(16, "export.png");
            this.imageListRibbon.Images.SetKeyName(17, "import.png");
            this.imageListRibbon.Images.SetKeyName(18, "clear.png");
            // 
            // adxCommandBarButton8
            // 
            this.adxCommandBarButton8.Caption = "Save Query";
            this.adxCommandBarButton8.ControlTag = "b0de7532-4a5e-49ac-b18c-830f8ee1eb09";
            this.adxCommandBarButton8.Image = 14;
            this.adxCommandBarButton8.ImageList = this.imageListRibbon;
            this.adxCommandBarButton8.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxCommandBarButton8.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.adxCommandBarButton8.UpdateCounter = 4;
            this.adxCommandBarButton8.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton8_Click);
            // 
            // adxCommandBarButton3
            // 
            this.adxCommandBarButton3.Caption = "Query Wizard";
            this.adxCommandBarButton3.ControlTag = "7c04af3a-4ba2-4b9b-ba16-4387c4ca3c6a";
            this.adxCommandBarButton3.Image = 9;
            this.adxCommandBarButton3.ImageList = this.imageListRibbon;
            this.adxCommandBarButton3.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxCommandBarButton3.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.adxCommandBarButton3.UpdateCounter = 8;
            this.adxCommandBarButton3.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton3_Click);
            // 
            // adxCommandBarButton4
            // 
            this.adxCommandBarButton4.Caption = "Refresh Query";
            this.adxCommandBarButton4.ControlTag = "a7d03823-18b2-40b0-ba77-7fa48dcd7cdb";
            this.adxCommandBarButton4.Image = 7;
            this.adxCommandBarButton4.ImageList = this.imageListRibbon;
            this.adxCommandBarButton4.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxCommandBarButton4.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.adxCommandBarButton4.UpdateCounter = 8;
            this.adxCommandBarButton4.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton4_Click);
            // 
            // adxCommandBarButton5
            // 
            this.adxCommandBarButton5.Caption = "Show Code";
            this.adxCommandBarButton5.ControlTag = "60f7139d-becf-44ff-822e-79828e5629a7";
            this.adxCommandBarButton5.Image = 8;
            this.adxCommandBarButton5.ImageList = this.imageListRibbon;
            this.adxCommandBarButton5.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxCommandBarButton5.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.adxCommandBarButton5.UpdateCounter = 8;
            this.adxCommandBarButton5.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton5_Click);
            // 
            // adxCommandBarButton9
            // 
            this.adxCommandBarButton9.Caption = "Export Queries";
            this.adxCommandBarButton9.ControlTag = "9181e699-a691-4263-89d3-b1f1c86cede3";
            this.adxCommandBarButton9.Image = 16;
            this.adxCommandBarButton9.ImageList = this.imageListRibbon;
            this.adxCommandBarButton9.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxCommandBarButton9.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.adxCommandBarButton9.UpdateCounter = 4;
            this.adxCommandBarButton9.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton9_Click);
            // 
            // adxCommandBarButton10
            // 
            this.adxCommandBarButton10.Caption = "Import Queries";
            this.adxCommandBarButton10.ControlTag = "33d9e867-8636-4b2a-a869-f216075cee31";
            this.adxCommandBarButton10.Image = 17;
            this.adxCommandBarButton10.ImageList = this.imageListRibbon;
            this.adxCommandBarButton10.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxCommandBarButton10.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.adxCommandBarButton10.UpdateCounter = 4;
            this.adxCommandBarButton10.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton10_Click);
            // 
            // adxCommandBarButton11
            // 
            this.adxCommandBarButton11.Caption = "Clear Sheet";
            this.adxCommandBarButton11.ControlTag = "7885234e-bb40-4635-b453-b1dc8abfe406";
            this.adxCommandBarButton11.Image = 18;
            this.adxCommandBarButton11.ImageList = this.imageListRibbon;
            this.adxCommandBarButton11.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxCommandBarButton11.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.adxCommandBarButton11.UpdateCounter = 4;
            this.adxCommandBarButton11.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton11_Click);
            // 
            // adxCommandBarButton2
            // 
            this.adxCommandBarButton2.Caption = "Relevance Editor";
            this.adxCommandBarButton2.ControlTag = "b6e26a75-e9af-474c-a282-c20195cb5442";
            this.adxCommandBarButton2.Image = 6;
            this.adxCommandBarButton2.ImageList = this.imageListRibbon;
            this.adxCommandBarButton2.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxCommandBarButton2.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.adxCommandBarButton2.UpdateCounter = 14;
            this.adxCommandBarButton2.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton2_Click);
            // 
            // adxCommandBarButton1
            // 
            this.adxCommandBarButton1.Caption = "Configuration";
            this.adxCommandBarButton1.ControlTag = "a83548e0-b97e-4acf-bdab-b4258c481f41";
            this.adxCommandBarButton1.Image = 5;
            this.adxCommandBarButton1.ImageList = this.imageListRibbon;
            this.adxCommandBarButton1.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxCommandBarButton1.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.adxCommandBarButton1.UpdateCounter = 10;
            this.adxCommandBarButton1.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton1_Click);
            // 
            // adxCommandBarButton6
            // 
            this.adxCommandBarButton6.Caption = "About";
            this.adxCommandBarButton6.ControlTag = "48740b11-98a5-4cc7-84a3-988926273772";
            this.adxCommandBarButton6.Image = 12;
            this.adxCommandBarButton6.ImageList = this.imageListRibbon;
            this.adxCommandBarButton6.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxCommandBarButton6.Style = AddinExpress.MSO.ADXMsoButtonStyle.adxMsoButtonIconAndCaption;
            this.adxCommandBarButton6.UpdateCounter = 7;
            this.adxCommandBarButton6.Click += new AddinExpress.MSO.ADXClick_EventHandler(this.adxCommandBarButton6_Click);
            // 
            // adxExcelTaskPanesManagerBF
            // 
            this.adxExcelTaskPanesManagerBF.Items.Add(this.adxExcelTaskPanesCollectionItemBF);
            this.adxExcelTaskPanesManagerBF.SetOwner(this);
            // 
            // adxExcelTaskPanesCollectionItemBF
            // 
            this.adxExcelTaskPanesCollectionItemBF.AlwaysShowHeader = true;
            this.adxExcelTaskPanesCollectionItemBF.CloseButton = true;
            this.adxExcelTaskPanesCollectionItemBF.Enabled = false;
            this.adxExcelTaskPanesCollectionItemBF.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Top;
            this.adxExcelTaskPanesCollectionItemBF.TaskPaneClassName = "BigFixExcelConnector.ADXExcelTaskPaneSessEditor";
            // 
            // adxRibbonTab1
            // 
            this.adxRibbonTab1.Caption = "IBM BigFix";
            this.adxRibbonTab1.Controls.Add(this.adxRibbonGroup1);
            this.adxRibbonTab1.Id = "adxRibbonTab_90435f7b6aaa43379be115f0436bc9e5";
            this.adxRibbonTab1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonGroup1
            // 
            this.adxRibbonGroup1.Caption = "IBM BigFix Connector";
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButton6);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator1);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButton8);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator2);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButton2);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator3);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButton4);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator4);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButton5);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator5);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButton9);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator6);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButton10);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator7);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButton11);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator8);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButton3);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator9);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButton1);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator10);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButton7);
            this.adxRibbonGroup1.Id = "adxRibbonGroup_d9db01607ca249cfaa68b209412135ec";
            this.adxRibbonGroup1.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButton6
            // 
            this.adxRibbonButton6.Caption = "Open Query Definition";
            this.adxRibbonButton6.Description = "Open Previously Saved Query Definitions";
            this.adxRibbonButton6.Id = "adxRibbonButton_e26131608b9b4a80ae20e81828105672";
            this.adxRibbonButton6.Image = 13;
            this.adxRibbonButton6.ImageList = this.imageListRibbon;
            this.adxRibbonButton6.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton6.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton6.ScreenTip = "Open Previously Saved Query Definitions";
            this.adxRibbonButton6.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButton6.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton6_OnClick);
            // 
            // adxRibbonSeparator1
            // 
            this.adxRibbonSeparator1.Id = "adxRibbonSeparator_97b6badb7f5d49f7b66f5b9583904516";
            this.adxRibbonSeparator1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButton8
            // 
            this.adxRibbonButton8.Caption = "Save Query Definition";
            this.adxRibbonButton8.Description = "Save Query Definitions into the User Registry";
            this.adxRibbonButton8.Id = "adxRibbonButton_a3bc5b95edf1407f8020ca5bd086203a";
            this.adxRibbonButton8.Image = 14;
            this.adxRibbonButton8.ImageList = this.imageListRibbon;
            this.adxRibbonButton8.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton8.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton8.ScreenTip = "Save Query Definitions into the User Registry";
            this.adxRibbonButton8.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButton8.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton8_OnClick);
            // 
            // adxRibbonSeparator2
            // 
            this.adxRibbonSeparator2.Id = "adxRibbonSeparator_c9c10bfeca2d4a4b9c50e8ee2c341801";
            this.adxRibbonSeparator2.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButton2
            // 
            this.adxRibbonButton2.Caption = "Execute Query Wizard";
            this.adxRibbonButton2.Description = "Query Wizard will Generate Relevance Statements Automatically";
            this.adxRibbonButton2.Id = "adxRibbonButton_58e9597183e349248e8998b6daff8302";
            this.adxRibbonButton2.Image = 9;
            this.adxRibbonButton2.ImageList = this.imageListRibbon;
            this.adxRibbonButton2.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton2.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton2.ScreenTip = "Query Wizard will Generate Relevance Statements Automatically";
            this.adxRibbonButton2.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButton2.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton2_OnClick);
            // 
            // adxRibbonSeparator3
            // 
            this.adxRibbonSeparator3.Id = "adxRibbonSeparator_c89b19b4b16e43ba9ecbc5c108f1548c";
            this.adxRibbonSeparator3.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButton4
            // 
            this.adxRibbonButton4.Caption = "Refresh Generated Query";
            this.adxRibbonButton4.Description = "Re-execute Query Already Defined in Worksheet";
            this.adxRibbonButton4.Id = "adxRibbonButton_ec454026ecdb488ea9f9162379c80cce";
            this.adxRibbonButton4.Image = 7;
            this.adxRibbonButton4.ImageList = this.imageListRibbon;
            this.adxRibbonButton4.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton4.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton4.ScreenTip = "Re-execute Query Already Defined in Worksheet";
            this.adxRibbonButton4.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButton4.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton4_OnClick);
            // 
            // adxRibbonSeparator4
            // 
            this.adxRibbonSeparator4.Id = "adxRibbonSeparator_01f0184d6e6c48a78e40590021b6c081";
            this.adxRibbonSeparator4.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButton5
            // 
            this.adxRibbonButton5.Caption = "Show Relevance Code";
            this.adxRibbonButton5.Description = "Show Relevance Statement Generated by Query Wizard";
            this.adxRibbonButton5.Id = "adxRibbonButton_ce810cfbb3c6411c8afdb0d012219451";
            this.adxRibbonButton5.Image = 8;
            this.adxRibbonButton5.ImageList = this.imageListRibbon;
            this.adxRibbonButton5.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton5.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton5.ScreenTip = "Show Relevance Statement Generated by Query Wizard";
            this.adxRibbonButton5.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButton5.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton5_OnClick);
            // 
            // adxRibbonSeparator5
            // 
            this.adxRibbonSeparator5.Id = "adxRibbonSeparator_67e3843ecbcd4d2bb378c652579ae03a";
            this.adxRibbonSeparator5.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButton9
            // 
            this.adxRibbonButton9.Caption = "Export Query Definitions";
            this.adxRibbonButton9.Description = "Export Saved Query Definitions to a BCE (BigFix Connector Export) File";
            this.adxRibbonButton9.Id = "adxRibbonButton_0030f3d5cd8b4ec1a6bc2e752fe9aa30";
            this.adxRibbonButton9.Image = 16;
            this.adxRibbonButton9.ImageList = this.imageListRibbon;
            this.adxRibbonButton9.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton9.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton9.ScreenTip = "Export Saved Query Definitions to a BCE (BigFix Connector Export) File";
            this.adxRibbonButton9.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButton9.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton9_OnClick);
            // 
            // adxRibbonSeparator6
            // 
            this.adxRibbonSeparator6.Id = "adxRibbonSeparator_61c1518751f64358bced7afc662643da";
            this.adxRibbonSeparator6.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButton10
            // 
            this.adxRibbonButton10.Caption = "Import Query Definitions";
            this.adxRibbonButton10.Description = "Import Saved Query Definitions from a BCE (BigFix Connector Export) File";
            this.adxRibbonButton10.Id = "adxRibbonButton_48fe0873ab79455b81e34f611c563201";
            this.adxRibbonButton10.Image = 17;
            this.adxRibbonButton10.ImageList = this.imageListRibbon;
            this.adxRibbonButton10.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton10.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton10.ScreenTip = "Import Saved Query Definitions from a BCE (BigFix Connector Export) File";
            this.adxRibbonButton10.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButton10.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton10_OnClick);
            // 
            // adxRibbonSeparator7
            // 
            this.adxRibbonSeparator7.Id = "adxRibbonSeparator_0cc89cc905234e4680fa2bfd9f3c122e";
            this.adxRibbonSeparator7.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButton11
            // 
            this.adxRibbonButton11.Caption = "Clear Current Sheet";
            this.adxRibbonButton11.Description = "Clear Everything on the Current Worksheet";
            this.adxRibbonButton11.Id = "adxRibbonButton_b534b2b12bee4caeacb8deb0235a0676";
            this.adxRibbonButton11.Image = 18;
            this.adxRibbonButton11.ImageList = this.imageListRibbon;
            this.adxRibbonButton11.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton11.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton11.ScreenTip = "Clear Everything on the Current Worksheet";
            this.adxRibbonButton11.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButton11.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton11_OnClick);
            // 
            // adxRibbonSeparator8
            // 
            this.adxRibbonSeparator8.Id = "adxRibbonSeparator_6739bd70636147cab6246f4d5dfab5a3";
            this.adxRibbonSeparator8.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButton3
            // 
            this.adxRibbonButton3.Caption = "Session Relevance Editor";
            this.adxRibbonButton3.Description = "Session Relevance Editor Used for Testing";
            this.adxRibbonButton3.Id = "adxRibbonButton_55d2c0e2bbfc4ed39694db3693cdbc13";
            this.adxRibbonButton3.Image = 6;
            this.adxRibbonButton3.ImageList = this.imageListRibbon;
            this.adxRibbonButton3.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton3.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton3.ScreenTip = "Session Relevance Editor Used for Testing";
            this.adxRibbonButton3.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButton3.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton3_OnClick);
            // 
            // adxRibbonSeparator9
            // 
            this.adxRibbonSeparator9.Id = "adxRibbonSeparator_7d36b188941641f3b81f0e3ac2aedf57";
            this.adxRibbonSeparator9.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButton1
            // 
            this.adxRibbonButton1.Caption = "Configuration";
            this.adxRibbonButton1.Description = "Connnection Settings to IBM BigFix Web Reports";
            this.adxRibbonButton1.Id = "adxRibbonButton_a7012dff158c4e15926ce96980b13f61";
            this.adxRibbonButton1.Image = 15;
            this.adxRibbonButton1.ImageList = this.imageListRibbon;
            this.adxRibbonButton1.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton1.ScreenTip = "Connnection Settings to IBM BigFix Web Reports";
            this.adxRibbonButton1.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButton1.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton1_OnClick);
            // 
            // adxRibbonSeparator10
            // 
            this.adxRibbonSeparator10.Id = "adxRibbonSeparator_b0a0c2037abb40d190145775506e2ec8";
            this.adxRibbonSeparator10.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButton7
            // 
            this.adxRibbonButton7.Caption = "About IBM BigFix Connector";
            this.adxRibbonButton7.Description = "Some Information About the IBM BigFix Excel Connector";
            this.adxRibbonButton7.Id = "adxRibbonButton_7954cb173c2443c3b4f781b54e2fcb74";
            this.adxRibbonButton7.Image = 12;
            this.adxRibbonButton7.ImageList = this.imageListRibbon;
            this.adxRibbonButton7.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButton7.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButton7.ScreenTip = "Some Information About the IBM BigFix Excel Connector";
            this.adxRibbonButton7.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButton7.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButton7_OnClick);
            // 
            // AddinModule
            // 
            this.AddinName = "IBM BigFix Excel Connector";
            this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaExcel;
            this.AddinInitialize += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinInitialize);

        }

        #region Open Report ====================================================================================

        void adxRibbonButton6_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            OpenReport();
        }

        void adxCommandBarButton7_Click(object sender)
        {
            OpenReport();
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

        private void OpenReport()
        {
            try
            {
                CheckForOpenWorkbook();

                formWizard.NeedToRefreshWizard = true;
                FormOpen formOpenWin = new FormOpen();
                // formOpenWin.ShowDialog();
                String returnStatus = formOpenWin.RunFormOpen();

                if (returnStatus == "true")
                {
                    refreshQuery();
                }
                else if (returnStatus == "false")
                {
                    OpenReadyMessage();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        private void OpenReadyMessage()
        {
            // Clear spreadsheet
            (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.ClearContents();
            DeleteCharts();

            Excel.Range titleRow = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", "CA1");
            titleRow.Select();
            titleRow.RowHeight = 33;
            titleRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(184, 204, 228));

            Excel.Worksheet hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
            String reportName = "";
            reportName = hiddenWorksheet.get_Range("A17", "A17").Value.ToString();

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

            (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[3, 1] = "Report opened, run the Wizard or click the Refresh button";

            Excel.Range rowSeparator = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A3", "CA3");
            rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(184, 204, 228)); // 
            rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1; // xlContinuous
            rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = 4; // xlThick

            // Place the cursor in cell A1 - which is at the start of the document
            (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", "A1").Select();

        }

        #endregion

        #region Save Report ====================================================================================
        void adxRibbonButton8_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            SaveReport();
        }

        void adxCommandBarButton8_Click(object sender)
        {
            SaveReport();
        }

        private void SaveReport()
        {
            try
            {
                Excel.Worksheet hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
            }
            catch
            {
                MessageBox.Show("There is currently no Relevance statement in the worksheet. Use the Relevance Query Wizard to generate one first.", "No Relevance Statement", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            FormSave formSaveWin = new FormSave();
            formSaveWin.ShowDialog();

        }

        #endregion

        #region Execute Query Wizard ===========================================================================
        void adxRibbonButton2_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            showQueryWizard();
        }

        void adxCommandBarButton3_Click(object sender)
        {
            showQueryWizard();
        }

        private void CheckForOpenWorkbook()
        {
            // Test to make sure that a Workbook/Worksheet is opened
            try
            {
                int NumberOfWorksheets = (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.Count;
            }
            catch
            {
                MessageBox.Show("There is no worksheet opened.", "Error opening Wizard", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Workbooks.Add();
                return;
            }
        }

        private void showQueryWizard()
        {
            formWizard.userStarted = true;
            formWizard.RetrieveSettings();
            formWizard.cacheSites();

            CheckForOpenWorkbook();

            // Test to see if this Workbook/Worksheet has a previously saved Session Relevance Query
            // If yes, then retrieve and restore data
            // If no, just open the Wizard
            try
            {
                Excel.Worksheet hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                Excel.Range CellTest = hiddenWorksheet.get_Range("A6", "A6");

                if (CellTest.Value == null)
                {
                    // The hidden worksheet is called "BigFixExcelConnector"
                    // Starting in Version 3 where the Refresh feature was added, the cell A6 is also new
                    // If the hidden worksheet is there but no A6, then Refresh will not work
                    formWizard.ShowDialog();
                }
                else
                {
                    formWizard.RestoreQueryWizardPreferences();
                    formWizard.ShowDialog();
                }

            }
            catch
            {
                formWizard.ShowDialog();
            }
        }

        #endregion

        #region Refresh Generated Query ========================================================================
        void adxRibbonButton4_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            refreshQuery();
        }

        void adxCommandBarButton4_Click(object sender)
        {
            refreshQuery();
        }

        private void refreshQuery()
        {
            Excel.Range CellA1;
            Excel.Range CellA6;

            try
            {
                TimeSpan t3;
                HiResTimer totalRefreshTime = new HiResTimer();
                totalRefreshTime.Start();

                Excel.Worksheet hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                CellA1 = hiddenWorksheet.get_Range("A1", "A1");
                CellA6 = hiddenWorksheet.get_Range("A6", "A6");

                if (CellA6.Value == null)
                {
                    // The hidden worksheet is called "BigFixExcelConnector"
                    // Starting in Version 3 where the Refresh feature was added, the cell A6 is also new
                    // If the hidden worksheet is there but no A6, then Refresh will not work
                    MessageBox.Show("This worksheet was probably saved by an older version of the BigFix Excel Connector, it cannot be refreshed. Please use the Query Wizard from Version 3.0 or above.", "Cannot Refresh Query", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // Clear spreadsheet
                    (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.ClearContents();
                    DeleteCharts();
                    (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = "Processing query, please wait...";

                    formWizard.userStarted = true;
                    formWizard.RetrieveSettings();
                    formWizard.RestoreQueryWizardPreferences();
                    formWizard.ProcessRelevanceThenWriteToExcel(CellA1.Value.ToString(), true);
                }

                totalRefreshTime.Stop();
                t3 = TimeSpan.FromMilliseconds(totalRefreshTime.ElapsedMicroseconds / 1000);

                Excel.Range rangeStorage;
                rangeStorage = hiddenWorksheet.get_Range("A16", "A16");
                rangeStorage.Value = t3.ToString().Remove(t3.ToString().Length - 4) + " / " + formWizard.QueryTime;

                formWizard.NeedToRefreshWizard = true;
            }
            catch
            {
                MessageBox.Show("There is currently no Relevance statement in the worksheet. Use the Relevance Query Wizard to generate one first.", "No Relevance Statement", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        #endregion

        #region Show Relevance Code ============================================================================
        void adxRibbonButton5_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            showRelevanceCode();
        }

        void adxCommandBarButton5_Click(object sender)
        {
            showRelevanceCode();
        }

        private void showRelevanceCode()
        {
            try
            {
                Excel.Worksheet hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                Excel.Range relevanceCell = hiddenWorksheet.get_Range("A1", "A1");
                ShowRelevanceCode showCodeWin = new ShowRelevanceCode();
                showCodeWin.ShowDialog();
            }
            catch
            {
                // MessageBox.Show(ex.Message);
                MessageBox.Show("There is currently no Relevance statement in the worksheet. Use the Relevance Query Wizard to generate one first.", "No Relevance Statement", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion

        #region Export Query Definitions =======================================================================
        void adxRibbonButton9_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            ExportQueryDefinitions();
        }

        void adxCommandBarButton9_Click(object sender)
        {
            ExportQueryDefinitions();
        }

        private void ExportQueryDefinitions()
        {
            try
            {
                string strRegistrySection = "HKEY_CURRENT_USER\\Software\\BigFix\\ExcelConnector\\SavedQueries";

                // OpenFileDialog exportFile = new OpenFileDialog();
                SaveFileDialog exportFile = new SaveFileDialog();
                exportFile.CheckFileExists = false;
                exportFile.Filter = "IBM BigFix Excel Connector Export files (*.bce)|*.bce|All files (*.*)|*.*";
                exportFile.Title = "Specify the Export file";
                exportFile.FileName = System.Windows.Forms.SystemInformation.UserName + ".bce";

                if (exportFile.ShowDialog() == DialogResult.OK)
                {
                    ExportKey(strRegistrySection, exportFile.FileName);

                    // If the registry is empty, export file will not be created, so compressing and deleting not necessary
                    if (File.Exists(exportFile.FileName))
                    {
                        FileInfo myFile = new FileInfo(exportFile.FileName);

                        //Pass the file path and file name to the StreamReader constructor
                        StreamReader sr = new StreamReader(exportFile.FileName);
                        String line = "";
                        ArrayList savedReports = new ArrayList();

                        //Read the first line of text
                        line = sr.ReadLine();

                        //Continue to read until you reach end of file
                        while (line != null)
                        {
                            if (line.StartsWith("\"Name\"="))
                                savedReports.Add(line.Substring(line.IndexOf("=\"") +2, line.Length-9));
                            //Read the next line
                            line = sr.ReadLine();
                        }

                        //close the file
                        sr.Close();

                        savedReports.Sort();

                        string[] savedReportsAsString = (string[])savedReports.ToArray(typeof(string));
                        int showHowMany = 20;

                        if (savedReports.Count > showHowMany)
                        {
                            Array.Resize(ref savedReportsAsString, showHowMany);

                            MessageBox.Show("The following " + savedReports.Count +
                                " report definitions are exported to file " + exportFile.FileName +
                                "\r\n\r\n" + string.Join("\r\n", savedReportsAsString) +
                                "\r\n:\n\r\r\n" + "There are " + (savedReports.Count - showHowMany).ToString() + " more.",
                                "Export successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else if (savedReports.Count == 0)
                        {
                            MessageBox.Show("No saved report definitions to export", "Nothing to export", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                        {
                            MessageBox.Show("The following " + savedReports.Count +
                                " report definitions are exported to file " + exportFile.FileName +
                                "\r\n\r\n" + string.Join("\r\n", savedReportsAsString),
                                "Export successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }

                        Compress(myFile);
                        File.Delete(exportFile.FileName);
                        File.Move(exportFile.FileName + ".gz", exportFile.FileName);
                    }
                    else
                    {
                        MessageBox.Show("No saved report definitions to export", "Nothing to export", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Exporting File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void ExportKey(string RegKey, string SavePath)
        {
            string path = "\"" + SavePath + "\"";
            string key = "\"" + RegKey + "\"";

            var proc = new Process();
            try
            {
                proc.StartInfo.FileName = "regedit.exe";
                proc.StartInfo.UseShellExecute = false;
                proc = Process.Start("regedit.exe", "/e " + path + " " + key + "");

                if (proc != null) proc.WaitForExit();
            }
            finally
            {
                if (proc != null) proc.Dispose();
            }
        }

        #endregion

        #region Import Query Definitions =======================================================================
        void adxRibbonButton10_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            ImportQueryDefinitions();
        }

        void adxCommandBarButton10_Click(object sender)
        {
            ImportQueryDefinitions();
        }
        
        private void ImportQueryDefinitions()
        {
            try
            {
                OpenFileDialog importFile = new OpenFileDialog();
                importFile.CheckFileExists = true;
                importFile.Filter = "IBM BigFix Excel Connector Export files (*.bce)|*.bce|All files (*.*)|*.*";
                importFile.Title = "Specify the Import file";

                if (importFile.ShowDialog() == DialogResult.OK)
                {
                    File.Copy(importFile.FileName, importFile.FileName + "x.gz");
                    FileInfo myFile = new FileInfo(importFile.FileName + "x.gz");
                    Decompress(myFile);
                    ImportKey(importFile.FileName + "x");

                    //Pass the file path and file name to the StreamReader constructor
                    StreamReader sr = new StreamReader(importFile.FileName + "x");
                    String line = "";
                    ArrayList savedReports = new ArrayList();

                    //Read the first line of text
                    line = sr.ReadLine();

                    //Continue to read until you reach end of file
                    while (line != null)
                    {
                        if (line.StartsWith("\"Name\"="))
                            savedReports.Add(line.Substring(line.IndexOf("=\"") + 2, line.Length - 9));
                        //Read the next line
                        line = sr.ReadLine();
                    }

                    //close the file
                    sr.Close();

                    savedReports.Sort();

                    string[] savedReportsAsString = (string[])savedReports.ToArray(typeof(string));
                    int showHowMany = 20;

                    if (savedReports.Count > showHowMany)
                    {
                        Array.Resize(ref savedReportsAsString, showHowMany);

                        MessageBox.Show("The following " + savedReports.Count +
                            " report definitions are imported\r\n\r\n" + string.Join("\r\n", savedReportsAsString) +
                            "\r\n:\n\r\r\n" + "There are " + (savedReports.Count - showHowMany).ToString() + " more.",
                            "Import successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if (savedReports.Count == 0)
                    {
                        MessageBox.Show("No saved report definitions to import", "Nothing to import", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    else 
                    {
                        MessageBox.Show("The following " + savedReports.Count +
                            " report definitions are imported\r\n\r\n" + string.Join("\r\n", savedReportsAsString),
                            "Import successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    File.Delete(importFile.FileName + "x.gz");
                    File.Delete(importFile.FileName + "x");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error Importing File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void ImportKey(string SavePath)
        {
            string path = "\"" + SavePath + "\"";

            var proc = new Process();
            try
            {
                proc.StartInfo.FileName = "regedit.exe";
                proc.StartInfo.UseShellExecute = false;
                proc = Process.Start("regedit.exe", "/s " + path);

                if (proc != null) proc.WaitForExit();
            }
            finally
            {
                if (proc != null) proc.Dispose();
            }
        }

        #endregion

        #region Session Relevance Editor =======================================================================
        void adxRibbonButton3_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            startSessionRelevanceEditor();
        }

        void adxCommandBarButton2_Click(object sender)
        {
            startSessionRelevanceEditor();
        }

        private void startSessionRelevanceEditor()
        {
            try
            {
                this.adxExcelTaskPanesCollectionItemBF.Enabled = true;
                if (this.adxExcelTaskPanesCollectionItemBF.TaskPaneInstance == null)
                    this.adxExcelTaskPanesCollectionItemBF.CreateTaskPaneInstance();

                String comboBoxTaskPaneLocation = GetSettings("ExcelConnector", "TaskPaneLocation");

                if (comboBoxTaskPaneLocation == "Top")
                    this.adxExcelTaskPanesCollectionItemBF.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Top;
                else if (comboBoxTaskPaneLocation == "Right")
                    this.adxExcelTaskPanesCollectionItemBF.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Right;
                else if (comboBoxTaskPaneLocation == "Bottom")
                    this.adxExcelTaskPanesCollectionItemBF.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Bottom;
                else if (comboBoxTaskPaneLocation == "Left")
                    this.adxExcelTaskPanesCollectionItemBF.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Left;
                else
                    this.adxExcelTaskPanesCollectionItemBF.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Top;

                this.adxExcelTaskPanesCollectionItemBF.TaskPaneInstance.Show();
                this.adxExcelTaskPanesCollectionItemBF.TaskPaneInstance.Activate();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error opening Session Relevance Editor", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        #endregion

        #region Clear Sheet ====================================================================================
        void adxRibbonButton11_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            ClearSheet();
        }

        void adxCommandBarButton11_Click(object sender)
        {
            ClearSheet();
        }

        private void ClearSheet()
        {
            try
            {
                // Clear spreadsheet of content and formats
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.ClearContents();
                DeleteCharts();
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.ClearFormats();
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.Select();
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.RowHeight = 15;
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells.ColumnWidth = 8.43;

                // Place the cursor in cell A1 - which is at the start of the document
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", "A1").Select();

                ((Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet).Name = "IEM Connector Sheet1";
            }
            catch
            {
                // If there is an error, it is probably because the hidden BigFixExcelConnector worksheet is not available, which is OK
            }

            try
            {
                Excel.Worksheet hiddenWorksheet = (Excel.Worksheet)(AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("BigFixExcelConnector");
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.DisplayAlerts = false;
                hiddenWorksheet.Delete();
                (AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.DisplayAlerts = true;
            }
            catch
            {
                // If there is an error, it is probably because the hidden BigFixExcelConnector worksheet is not available, which is OK
            }

        }

        #endregion

        #region Configuration ==================================================================================
        void adxRibbonButton1_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            showConfiguration();
        }

        void adxCommandBarButton1_Click(object sender)
        {
            showConfiguration();
        }

        private void showConfiguration()
        {
            try
            {
                FormConfig configWin = new FormConfig();
                configWin.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        #endregion 

        #region About BigFix Connector =========================================================================
        void adxRibbonButton7_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            showAboutBox();
        }

        void adxCommandBarButton6_Click(object sender)
        {
            showAboutBox();
        }

        private void showAboutBox()
        {
            AboutBox ab = new AboutBox();
            ab.MoreRichTextBox.Text = "Developed in C# using IBM BigFix SOAP API\n\n" + 
            "Get more information from the IBM developerWorks BigFix Wiki:\n" + "https://www.ibm.com/developerworks/community/wikis/home?lang=en#/wiki/Tivoli%20Endpoint%20Manager/page/Excel%20Connector" +
                "\n\nPost questions, comments and bug reports at the IBM BigFix Forum:\n" + "https://forum.bigfix.com/t/bigfix-excel-connector/6152";
            ab.ShowDialog();
        }
        #endregion

        // ================================================================================================
        #endregion
 
        #region Add-in Express automatic code
 
        // Required by Add-in Express - do not modify
        // the methods within this region
 
        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }
 
        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }
 
        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

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

        #region Compress and Decompress Routines
        public static void Compress(FileInfo fi)
        {
            // Get the stream of the source file.
            using (FileStream inFile = fi.OpenRead())
            {
                // Prevent compressing hidden and already compressed files.
                if ((File.GetAttributes(fi.FullName) & FileAttributes.Hidden)
                        != FileAttributes.Hidden & fi.Extension != ".gz")
                {
                    // Create the compressed file.
                    using (FileStream outFile = File.Create(fi.FullName + ".gz"))
                    {
                        using (GZipStream Compress = new GZipStream(outFile,
                                CompressionMode.Compress))
                        {
                            // Copy the source file into the compression stream.
                            byte[] buffer = new byte[4096];
                            int numRead;
                            while ((numRead = inFile.Read(buffer, 0, buffer.Length)) != 0)
                            {
                                Compress.Write(buffer, 0, numRead);
                            }
                        }
                    }
                }
            }
        }

        public static void Decompress(FileInfo fi)
        {
            // Get the stream of the source file.
            using (FileStream inFile = fi.OpenRead())
            {
                // Get original file extension, for example "doc" from report.doc.gz.
                string curFile = fi.FullName;
                string origName = curFile.Remove(curFile.Length - fi.Extension.Length);

                //Create the decompressed file.
                using (FileStream outFile = File.Create(origName))
                {
                    using (GZipStream Decompress = new GZipStream(inFile,
                            CompressionMode.Decompress))
                    {
                        //Copy the decompression stream into the output file.
                        byte[] buffer = new byte[4096];
                        int numRead;
                        while ((numRead = Decompress.Read(buffer, 0, buffer.Length)) != 0)
                        {
                            outFile.Write(buffer, 0, numRead);
                        }
                    }
                }
            }
        }
        #endregion

        private void AddinModule_AddinInitialize(object sender, EventArgs e)
        {
            // http://www.add-in-express.com/forum/read.php?FID=5&TID=10278
            // The following is set to avoid the Excel: Old format or invalid type library error
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(ExcelApp.LanguageSettings.get_LanguageID(Office.MsoAppLanguageID.msoLanguageIDUI)); 

            formWizard = new FormWizard();
            Thread firstThread = new Thread(new ThreadStart(formWizard.cacheSites));
            // firstThread.Priority = ThreadPriority.Highest;
            firstThread.Start();

            // formWizard.cacheSites();
        }

        public Excel._Application ExcelApp
        {
            get
            {
                return (HostApplication as Excel._Application);
            }
        }
    }


}

