namespace DonateMonitor
{
    partial class Monitor
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
            this.lbECPAY_Status = new System.Windows.Forms.Label();
            this.Tb_MonitorOut = new System.Windows.Forms.TextBox();
            this.BtConfig = new System.Windows.Forms.Button();
            this.lbOPAY_Status = new System.Windows.Forms.Label();
            this.lbStreamlabs_Status = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lbECPAY_Status
            // 
            this.lbECPAY_Status.AutoSize = true;
            this.lbECPAY_Status.Font = new System.Drawing.Font("Noto Sans TC", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbECPAY_Status.Location = new System.Drawing.Point(12, 9);
            this.lbECPAY_Status.Name = "lbECPAY_Status";
            this.lbECPAY_Status.Size = new System.Drawing.Size(145, 27);
            this.lbECPAY_Status.TabIndex = 1;
            this.lbECPAY_Status.Text = "綠界狀態：無效";
            // 
            // Tb_MonitorOut
            // 
            this.Tb_MonitorOut.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Tb_MonitorOut.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Tb_MonitorOut.Location = new System.Drawing.Point(12, 108);
            this.Tb_MonitorOut.Multiline = true;
            this.Tb_MonitorOut.Name = "Tb_MonitorOut";
            this.Tb_MonitorOut.ReadOnly = true;
            this.Tb_MonitorOut.Size = new System.Drawing.Size(921, 572);
            this.Tb_MonitorOut.TabIndex = 2;
            // 
            // BtConfig
            // 
            this.BtConfig.Location = new System.Drawing.Point(858, 9);
            this.BtConfig.Name = "BtConfig";
            this.BtConfig.Size = new System.Drawing.Size(75, 33);
            this.BtConfig.TabIndex = 3;
            this.BtConfig.Text = "設 定";
            this.BtConfig.UseVisualStyleBackColor = true;
            this.BtConfig.Click += new System.EventHandler(this.BtConfig_Click);
            // 
            // lbOPAY_Status
            // 
            this.lbOPAY_Status.AutoSize = true;
            this.lbOPAY_Status.Font = new System.Drawing.Font("Noto Sans TC", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbOPAY_Status.Location = new System.Drawing.Point(12, 36);
            this.lbOPAY_Status.Name = "lbOPAY_Status";
            this.lbOPAY_Status.Size = new System.Drawing.Size(164, 27);
            this.lbOPAY_Status.TabIndex = 4;
            this.lbOPAY_Status.Text = "歐富寶狀態：無效";
            // 
            // lbStreamlabs_Status
            // 
            this.lbStreamlabs_Status.AutoSize = true;
            this.lbStreamlabs_Status.Font = new System.Drawing.Font("Noto Sans TC", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbStreamlabs_Status.Location = new System.Drawing.Point(12, 63);
            this.lbStreamlabs_Status.Name = "lbStreamlabs_Status";
            this.lbStreamlabs_Status.Size = new System.Drawing.Size(213, 27);
            this.lbStreamlabs_Status.TabIndex = 5;
            this.lbStreamlabs_Status.Text = "Streamlabs 狀態：無效";
            // 
            // Monitor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(945, 692);
            this.Controls.Add(this.lbStreamlabs_Status);
            this.Controls.Add(this.lbOPAY_Status);
            this.Controls.Add(this.BtConfig);
            this.Controls.Add(this.Tb_MonitorOut);
            this.Controls.Add(this.lbECPAY_Status);
            this.Font = new System.Drawing.Font("Noto Sans TC", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Monitor";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Monitor";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Monitor_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbECPAY_Status;
        private System.Windows.Forms.TextBox Tb_MonitorOut;
        private System.Windows.Forms.Button BtConfig;
        private System.Windows.Forms.Label lbOPAY_Status;
        private System.Windows.Forms.Label lbStreamlabs_Status;
    }
}