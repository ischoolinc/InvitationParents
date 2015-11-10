namespace InvitationParents
{
	partial class Form
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
            this.buttonX1 = new DevComponents.DotNetBar.ButtonX();
            this.buttonX2 = new DevComponents.DotNetBar.ButtonX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.checkBoxX1 = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.lnkDefault = new System.Windows.Forms.LinkLabel();
            this.lnkTemplate = new System.Windows.Forms.LinkLabel();
            this.lnkUpload = new System.Windows.Forms.LinkLabel();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.cboSelTemplate = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.SuspendLayout();
            // 
            // buttonX1
            // 
            this.buttonX1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX1.BackColor = System.Drawing.Color.Transparent;
            this.buttonX1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonX1.Location = new System.Drawing.Point(166, 178);
            this.buttonX1.Name = "buttonX1";
            this.buttonX1.Size = new System.Drawing.Size(75, 23);
            this.buttonX1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonX1.TabIndex = 0;
            this.buttonX1.Text = "列印";
            this.buttonX1.Click += new System.EventHandler(this.buttonX1_Click);
            // 
            // buttonX2
            // 
            this.buttonX2.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX2.BackColor = System.Drawing.Color.Transparent;
            this.buttonX2.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonX2.Location = new System.Drawing.Point(247, 178);
            this.buttonX2.Name = "buttonX2";
            this.buttonX2.Size = new System.Drawing.Size(75, 23);
            this.buttonX2.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonX2.TabIndex = 1;
            this.buttonX2.Text = "離開";
            this.buttonX2.Click += new System.EventHandler(this.buttonX2_Click);
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(12, 12);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(291, 44);
            this.labelX1.TabIndex = 2;
            this.labelX1.Text = "說明：本功能將會列印邀請學生之家長邀請函，\r\n方便學校邀請家長加入ischool App。";
            // 
            // checkBoxX1
            // 
            this.checkBoxX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.checkBoxX1.BackgroundStyle.Class = "";
            this.checkBoxX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.checkBoxX1.Location = new System.Drawing.Point(12, 62);
            this.checkBoxX1.Name = "checkBoxX1";
            this.checkBoxX1.Size = new System.Drawing.Size(108, 23);
            this.checkBoxX1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.checkBoxX1.TabIndex = 3;
            this.checkBoxX1.Text = "上傳至電子報表";
            this.checkBoxX1.CheckedChanged += new System.EventHandler(this.checkBoxX1_CheckedChanged);
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.BackColor = System.Drawing.Color.Transparent;
            this.linkLabel1.Location = new System.Drawing.Point(233, 142);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(86, 17);
            this.linkLabel1.TabIndex = 4;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "檢視合併欄位";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // lnkDefault
            // 
            this.lnkDefault.AutoSize = true;
            this.lnkDefault.BackColor = System.Drawing.Color.Transparent;
            this.lnkDefault.Location = new System.Drawing.Point(9, 142);
            this.lnkDefault.Name = "lnkDefault";
            this.lnkDefault.Size = new System.Drawing.Size(86, 17);
            this.lnkDefault.TabIndex = 5;
            this.lnkDefault.TabStop = true;
            this.lnkDefault.Text = "下載預設範本";
            this.lnkDefault.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkDefault_LinkClicked);
            // 
            // lnkTemplate
            // 
            this.lnkTemplate.AutoSize = true;
            this.lnkTemplate.BackColor = System.Drawing.Color.Transparent;
            this.lnkTemplate.Location = new System.Drawing.Point(101, 142);
            this.lnkTemplate.Name = "lnkTemplate";
            this.lnkTemplate.Size = new System.Drawing.Size(60, 17);
            this.lnkTemplate.TabIndex = 6;
            this.lnkTemplate.TabStop = true;
            this.lnkTemplate.Text = "下載範本";
            this.lnkTemplate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkTemplate_LinkClicked);
            // 
            // lnkUpload
            // 
            this.lnkUpload.AutoSize = true;
            this.lnkUpload.BackColor = System.Drawing.Color.Transparent;
            this.lnkUpload.Location = new System.Drawing.Point(167, 142);
            this.lnkUpload.Name = "lnkUpload";
            this.lnkUpload.Size = new System.Drawing.Size(60, 17);
            this.lnkUpload.TabIndex = 7;
            this.lnkUpload.TabStop = true;
            this.lnkUpload.Text = "上傳範本";
            this.lnkUpload.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkUpload_LinkClicked);
            // 
            // labelX2
            // 
            this.labelX2.AutoSize = true;
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(13, 98);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(60, 21);
            this.labelX2.TabIndex = 8;
            this.labelX2.Text = "使用範本";
            // 
            // cboSelTemplate
            // 
            this.cboSelTemplate.DisplayMember = "Text";
            this.cboSelTemplate.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSelTemplate.FormattingEnabled = true;
            this.cboSelTemplate.ItemHeight = 19;
            this.cboSelTemplate.Location = new System.Drawing.Point(80, 96);
            this.cboSelTemplate.Name = "cboSelTemplate";
            this.cboSelTemplate.Size = new System.Drawing.Size(121, 25);
            this.cboSelTemplate.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboSelTemplate.TabIndex = 9;
            // 
            // Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(340, 216);
            this.Controls.Add(this.cboSelTemplate);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.lnkUpload);
            this.Controls.Add(this.lnkTemplate);
            this.Controls.Add(this.lnkDefault);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.checkBoxX1);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.buttonX2);
            this.Controls.Add(this.buttonX1);
            this.DoubleBuffered = true;
            this.Name = "Form";
            this.Text = "列印學生家長邀請函";
            this.Load += new System.EventHandler(this.Form_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private DevComponents.DotNetBar.ButtonX buttonX1;
		private DevComponents.DotNetBar.ButtonX buttonX2;
		private DevComponents.DotNetBar.LabelX labelX1;
		private DevComponents.DotNetBar.Controls.CheckBoxX checkBoxX1;
		private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.LinkLabel lnkDefault;
        private System.Windows.Forms.LinkLabel lnkTemplate;
        private System.Windows.Forms.LinkLabel lnkUpload;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboSelTemplate;
	}
}