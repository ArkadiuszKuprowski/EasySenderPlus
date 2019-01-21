namespace easyDMSTool
{
    partial class easyDMSToolOptionDialog
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {            
            this.serverChoose = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.serverOther_txtbox = new System.Windows.Forms.TextBox();
            this.serverOther_rbtn = new System.Windows.Forms.RadioButton();
            this.serverProd_rbtn = new System.Windows.Forms.RadioButton();
            this.serverTest_rbtn = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.useProvidedUser_rbtn = new System.Windows.Forms.RadioButton();
            this.userID_txtbox = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.useDefaultUser_rbtn = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.serverChoose.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // serverChoose
            // 
            this.serverChoose.Controls.Add(this.panel1);
            this.serverChoose.Location = new System.Drawing.Point(18, 17);
            this.serverChoose.Name = "serverChoose";
            this.serverChoose.Size = new System.Drawing.Size(212, 160);
            this.serverChoose.TabIndex = 2;
            this.serverChoose.TabStop = false;
            this.serverChoose.Text = "Servers";

            
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.serverOther_txtbox);
            this.panel1.Controls.Add(this.serverOther_rbtn);
            this.panel1.Controls.Add(this.serverProd_rbtn);
            this.panel1.Controls.Add(this.serverTest_rbtn);
            this.panel1.Location = new System.Drawing.Point(13, 26);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(192, 93);
            this.panel1.TabIndex = 4;
            // 
            // serverOther_txtbox
            // 
            this.serverOther_txtbox.Enabled = global::easyDMSTool.Properties.Settings.Default.isCheckedServerOther_rbtn;
            this.serverOther_txtbox.Location = new System.Drawing.Point(0, 69);
            this.serverOther_txtbox.Name = "serverOther_txtbox";
            this.serverOther_txtbox.Size = new System.Drawing.Size(189, 20);
            this.serverOther_txtbox.TabIndex = 3;
            this.serverOther_txtbox.Text = global::easyDMSTool.Properties.Settings.Default.serverUrlOther;
            // 
            // serverOther_rbtn
            // 
            this.serverOther_rbtn.AutoSize = true;
            this.serverOther_rbtn.Checked = global::easyDMSTool.Properties.Settings.Default.isCheckedServerOther_rbtn;
            this.serverOther_rbtn.Location = new System.Drawing.Point(0, 46);
            this.serverOther_rbtn.Name = "serverOther_rbtn";
            this.serverOther_rbtn.Size = new System.Drawing.Size(49, 17);
            this.serverOther_rbtn.TabIndex = 2;
            this.serverOther_rbtn.TabStop = true;
            this.serverOther_rbtn.Text = "other";
            this.serverOther_rbtn.UseVisualStyleBackColor = true;
            this.serverOther_rbtn.CheckedChanged += new System.EventHandler(this.serverOther_rbtn_CheckedChanged);
            // 
            // serverProd_rbtn
            // 
            this.serverProd_rbtn.AutoSize = true;
            this.serverProd_rbtn.Checked = global::easyDMSTool.Properties.Settings.Default.isCheckedServerProd_rbtn;
            this.serverProd_rbtn.Location = new System.Drawing.Point(0, 23);
            this.serverProd_rbtn.Name = "serverProd_rbtn";
            this.serverProd_rbtn.Size = new System.Drawing.Size(92, 17);
            this.serverProd_rbtn.TabIndex = 1;
            this.serverProd_rbtn.TabStop = true;
            this.serverProd_rbtn.Text = "prod - deis335";
            this.serverProd_rbtn.UseVisualStyleBackColor = true;
            this.serverProd_rbtn.CheckedChanged += new System.EventHandler(this.serverProd_rbtn_CheckedChanged);
            // 
            // serverTest_rbtn
            // 
            this.serverTest_rbtn.AutoSize = true;
            this.serverTest_rbtn.Checked = global::easyDMSTool.Properties.Settings.Default.isCheckedServerTest_rbtn;
            this.serverTest_rbtn.Location = new System.Drawing.Point(0, 0);
            this.serverTest_rbtn.Name = "serverTest_rbtn";
            this.serverTest_rbtn.Size = new System.Drawing.Size(88, 17);
            this.serverTest_rbtn.TabIndex = 0;
            this.serverTest_rbtn.Text = "test - deis366";
            this.serverTest_rbtn.UseVisualStyleBackColor = true;
            this.serverTest_rbtn.CheckedChanged += new System.EventHandler(this.serverTest_rbtn_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.userID_txtbox);
            this.groupBox1.Controls.Add(this.panel2);
            this.groupBox1.Location = new System.Drawing.Point(18, 198);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(212, 212);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "User";
            // 
            // useProvidedUser_rbtn
            // 
            this.useProvidedUser_rbtn.AutoSize = true;
            this.useProvidedUser_rbtn.Checked = global::easyDMSTool.Properties.Settings.Default.isCheckedUseProvidedUser_rbtn;
            this.useProvidedUser_rbtn.Location = new System.Drawing.Point(0, 40);
            this.useProvidedUser_rbtn.Name = "useProvidedUser_rbtn";
            this.useProvidedUser_rbtn.Size = new System.Drawing.Size(156, 17);
            this.useProvidedUser_rbtn.TabIndex = 5;
            this.useProvidedUser_rbtn.TabStop = true;
            this.useProvidedUser_rbtn.Text = "Use Below User Credentials";
            this.useProvidedUser_rbtn.UseVisualStyleBackColor = true;
            this.useProvidedUser_rbtn.CheckedChanged += new System.EventHandler(this.useProvidedUser_rbtn_CheckedChanged);
            // 
            // userID_txtbox
            // 
            this.userID_txtbox.Enabled = global::easyDMSTool.Properties.Settings.Default.isEnabledUseProvidedUser_txtbox;
            this.userID_txtbox.Location = new System.Drawing.Point(14, 110);
            this.userID_txtbox.Name = "userID_txtbox";
            this.userID_txtbox.Size = new System.Drawing.Size(124, 20);
            this.userID_txtbox.TabIndex = 0;
            this.userID_txtbox.Text = global::easyDMSTool.Properties.Settings.Default.userID;
            this.userID_txtbox.TextChanged += new System.EventHandler(this.userID_txtbox_TextChanged);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.useProvidedUser_rbtn);
            this.panel2.Controls.Add(this.useDefaultUser_rbtn);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Location = new System.Drawing.Point(13, 34);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(193, 172);
            this.panel2.TabIndex = 5;
            // 
            // useDefaultUser_rbtn
            // 
            this.useDefaultUser_rbtn.AutoSize = true;
            this.useDefaultUser_rbtn.Checked = global::easyDMSTool.Properties.Settings.Default.isCheckedUseDefaultUser_rbtn;
            this.useDefaultUser_rbtn.Location = new System.Drawing.Point(0, 17);
            this.useDefaultUser_rbtn.Name = "useDefaultUser_rbtn";
            this.useDefaultUser_rbtn.Size = new System.Drawing.Size(161, 17);
            this.useDefaultUser_rbtn.TabIndex = 4;
            this.useDefaultUser_rbtn.TabStop = true;
            this.useDefaultUser_rbtn.Text = "Use Default User Credentials";
            this.useDefaultUser_rbtn.UseVisualStyleBackColor = true;
            this.useDefaultUser_rbtn.CheckedChanged += new System.EventHandler(this.useDefaultUser_rbtn_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(136, 82);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "User ID";
            // 
            // easyDMSToolOptionDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.serverChoose);
            this.Name = "easyDMSToolOptionDialog";
            this.Size = new System.Drawing.Size(627, 475);
            this.serverChoose.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        
        
        private System.Windows.Forms.RadioButton serverTest_rbtn;
        private System.Windows.Forms.RadioButton serverProd_rbtn;
        private System.Windows.Forms.GroupBox serverChoose;
        private System.Windows.Forms.RadioButton serverOther_rbtn;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox serverOther_txtbox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox userID_txtbox;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.RadioButton useProvidedUser_rbtn;
        private System.Windows.Forms.RadioButton useDefaultUser_rbtn;
    }
}
