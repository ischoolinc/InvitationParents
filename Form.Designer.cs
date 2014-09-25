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
			this.SuspendLayout();
			// 
			// buttonX1
			// 
			this.buttonX1.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
			this.buttonX1.BackColor = System.Drawing.Color.Transparent;
			this.buttonX1.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
			this.buttonX1.Location = new System.Drawing.Point(137, 115);
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
			this.buttonX2.Location = new System.Drawing.Point(218, 115);
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
			this.labelX1.Location = new System.Drawing.Point(19, 21);
			this.labelX1.Name = "labelX1";
			this.labelX1.Size = new System.Drawing.Size(282, 60);
			this.labelX1.TabIndex = 2;
			this.labelX1.Text = "說明：本功能將會列印邀請學生QRCODE，\r\n提供家長便利迅速與家長即時通訊App連結。";
			// 
			// checkBoxX1
			// 
			this.checkBoxX1.BackColor = System.Drawing.Color.Transparent;
			// 
			// 
			// 
			this.checkBoxX1.BackgroundStyle.Class = "";
			this.checkBoxX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
			this.checkBoxX1.Location = new System.Drawing.Point(12, 115);
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
			this.linkLabel1.Location = new System.Drawing.Point(16, 95);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(86, 17);
			this.linkLabel1.TabIndex = 4;
			this.linkLabel1.TabStop = true;
			this.linkLabel1.Text = "檢視合併欄位";
			// 
			// Form
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(305, 150);
			this.Controls.Add(this.linkLabel1);
			this.Controls.Add(this.checkBoxX1);
			this.Controls.Add(this.labelX1);
			this.Controls.Add(this.buttonX2);
			this.Controls.Add(this.buttonX1);
			this.DoubleBuffered = true;
			this.Name = "Form";
			this.Text = "列印學生家長邀請函";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private DevComponents.DotNetBar.ButtonX buttonX1;
		private DevComponents.DotNetBar.ButtonX buttonX2;
		private DevComponents.DotNetBar.LabelX labelX1;
		private DevComponents.DotNetBar.Controls.CheckBoxX checkBoxX1;
		private System.Windows.Forms.LinkLabel linkLabel1;
	}
}