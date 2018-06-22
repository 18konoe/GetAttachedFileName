namespace GetAttachedFileName
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.AttachedFileNameGroup = this.Factory.CreateRibbonGroup();
            this.insertNameMenu = this.Factory.CreateRibbonMenu();
            this.tab1.SuspendLayout();
            this.AttachedFileNameGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabNewMailMessage";
            this.tab1.Groups.Add(this.AttachedFileNameGroup);
            this.tab1.Label = "TabNewMailMessage";
            this.tab1.Name = "tab1";
            // 
            // AttachedFileNameGroup
            // 
            this.AttachedFileNameGroup.Items.Add(this.insertNameMenu);
            this.AttachedFileNameGroup.Label = "Attached File Name";
            this.AttachedFileNameGroup.Name = "AttachedFileNameGroup";
            // 
            // insertNameMenu
            // 
            this.insertNameMenu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.insertNameMenu.Dynamic = true;
            this.insertNameMenu.Image = ((System.Drawing.Image)(resources.GetObject("insertNameMenu.Image")));
            this.insertNameMenu.Label = "Insert File Name to Message";
            this.insertNameMenu.Name = "insertNameMenu";
            this.insertNameMenu.ShowImage = true;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.AttachedFileNameGroup.ResumeLayout(false);
            this.AttachedFileNameGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AttachedFileNameGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu insertNameMenu;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
