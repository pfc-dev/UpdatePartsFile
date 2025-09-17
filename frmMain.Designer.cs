namespace UpdatePartsFile
{
    partial class frmMain
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnChoseFile = new Button();
            txtFile = new TextBox();
            panel1 = new Panel();
            label4 = new Label();
            lblDateLastUpdated = new Label();
            label3 = new Label();
            cboPlant = new ComboBox();
            label2 = new Label();
            cboSapSystem = new ComboBox();
            btnQuit = new Button();
            btnUpdate = new Button();
            panel2 = new Panel();
            label1 = new Label();
            lblMessage = new Label();
            txtMain = new TextBox();
            panel1.SuspendLayout();
            panel2.SuspendLayout();
            SuspendLayout();
            // 
            // btnChoseFile
            // 
            btnChoseFile.Location = new Point(24, 18);
            btnChoseFile.Name = "btnChoseFile";
            btnChoseFile.Size = new Size(75, 23);
            btnChoseFile.TabIndex = 0;
            btnChoseFile.Text = "Choose File";
            btnChoseFile.UseVisualStyleBackColor = true;
            btnChoseFile.Click += btnChoseFile_Click;
            // 
            // txtFile
            // 
            txtFile.Location = new Point(134, 18);
            txtFile.Multiline = true;
            txtFile.Name = "txtFile";
            txtFile.Size = new Size(247, 51);
            txtFile.TabIndex = 1;
            // 
            // panel1
            // 
            panel1.BorderStyle = BorderStyle.FixedSingle;
            panel1.Controls.Add(label4);
            panel1.Controls.Add(lblDateLastUpdated);
            panel1.Controls.Add(label3);
            panel1.Controls.Add(cboPlant);
            panel1.Controls.Add(label2);
            panel1.Controls.Add(cboSapSystem);
            panel1.Controls.Add(btnQuit);
            panel1.Controls.Add(btnUpdate);
            panel1.Controls.Add(txtFile);
            panel1.Controls.Add(btnChoseFile);
            panel1.Location = new Point(12, 12);
            panel1.Name = "panel1";
            panel1.Size = new Size(544, 125);
            panel1.TabIndex = 2;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(26, 73);
            label4.Name = "label4";
            label4.Size = new Size(76, 15);
            label4.TabIndex = 9;
            label4.Text = "Last Updated";
            // 
            // lblDateLastUpdated
            // 
            lblDateLastUpdated.AutoSize = true;
            lblDateLastUpdated.Location = new Point(134, 73);
            lblDateLastUpdated.Name = "lblDateLastUpdated";
            lblDateLastUpdated.Size = new Size(38, 15);
            lblDateLastUpdated.TabIndex = 8;
            lblDateLastUpdated.Text = "label4";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(403, 95);
            label3.Name = "label3";
            label3.Size = new Size(34, 15);
            label3.TabIndex = 7;
            label3.Text = "Plant";
            // 
            // cboPlant
            // 
            cboPlant.FormattingEnabled = true;
            cboPlant.Items.AddRange(new object[] { "0300", "0310" });
            cboPlant.Location = new Point(453, 92);
            cboPlant.Name = "cboPlant";
            cboPlant.Size = new Size(66, 23);
            cboPlant.TabIndex = 6;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(30, 100);
            label2.Name = "label2";
            label2.Size = new Size(69, 15);
            label2.TabIndex = 5;
            label2.Text = "SAP System";
            // 
            // cboSapSystem
            // 
            cboSapSystem.FormattingEnabled = true;
            cboSapSystem.Items.AddRange(new object[] { "P01", "S4Q" });
            cboSapSystem.Location = new Point(134, 92);
            cboSapSystem.Name = "cboSapSystem";
            cboSapSystem.Size = new Size(89, 23);
            cboSapSystem.TabIndex = 4;
            cboSapSystem.SelectedValueChanged += cboSapSystem_SelectedValueChanged;
            // 
            // btnQuit
            // 
            btnQuit.Location = new Point(453, 46);
            btnQuit.Name = "btnQuit";
            btnQuit.Size = new Size(75, 23);
            btnQuit.TabIndex = 3;
            btnQuit.Text = "Quit";
            btnQuit.UseVisualStyleBackColor = true;
            btnQuit.Click += btnQuit_Click;
            // 
            // btnUpdate
            // 
            btnUpdate.Location = new Point(453, 17);
            btnUpdate.Name = "btnUpdate";
            btnUpdate.Size = new Size(75, 23);
            btnUpdate.TabIndex = 2;
            btnUpdate.Text = "Update";
            btnUpdate.UseVisualStyleBackColor = true;
            btnUpdate.Click += btnUpdate_Click;
            // 
            // panel2
            // 
            panel2.BorderStyle = BorderStyle.FixedSingle;
            panel2.Controls.Add(label1);
            panel2.Controls.Add(lblMessage);
            panel2.Location = new Point(12, 143);
            panel2.Name = "panel2";
            panel2.Size = new Size(544, 31);
            panel2.TabIndex = 3;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(8, 7);
            label1.Name = "label1";
            label1.Size = new Size(61, 15);
            label1.TabIndex = 1;
            label1.Text = "Messages:";
            // 
            // lblMessage
            // 
            lblMessage.AutoSize = true;
            lblMessage.Location = new Point(75, 7);
            lblMessage.Name = "lblMessage";
            lblMessage.Size = new Size(73, 15);
            lblMessage.TabIndex = 0;
            lblMessage.Text = "choose a file";
            // 
            // txtMain
            // 
            txtMain.BorderStyle = BorderStyle.FixedSingle;
            txtMain.Enabled = false;
            txtMain.Location = new Point(10, 186);
            txtMain.Multiline = true;
            txtMain.Name = "txtMain";
            txtMain.Size = new Size(546, 212);
            txtMain.TabIndex = 4;
            // 
            // frmMain
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(565, 416);
            Controls.Add(txtMain);
            Controls.Add(panel2);
            Controls.Add(panel1);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Name = "frmMain";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Update SAP Parts File";
            Load += frmMain_Load;
            Shown += frmMain_Shown;
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            panel2.ResumeLayout(false);
            panel2.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnChoseFile;
        private TextBox txtFile;
        private Panel panel1;
        private Button btnUpdate;
        private Panel panel2;
        private Label label1;
        private Label lblMessage;
        private Button btnQuit;
        private Label label2;
        private ComboBox cboSapSystem;
        private Label label3;
        private ComboBox cboPlant;
        private Label label4;
        private Label lblDateLastUpdated;
        private TextBox txtMain;
    }
}
