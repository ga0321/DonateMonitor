namespace DonateMonitor
{
    partial class Startup
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.Tb_ECPayAPIURL = new System.Windows.Forms.TextBox();
            this.BtnEnterMonitor = new System.Windows.Forms.Button();
            this.Tb_OPayAPIURL = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.Tb_StreamlabsKey = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.Tb_HiveBeeAPIURL = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.Tb_SoundAlertsURL = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Noto Sans TC", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(65, 75);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(221, 27);
            this.label1.TabIndex = 0;
            this.label1.Text = "綠界贊助成功動畫網址：";
            // 
            // Tb_ECPayAPIURL
            // 
            this.Tb_ECPayAPIURL.Font = new System.Drawing.Font("Noto Sans TC", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Tb_ECPayAPIURL.Location = new System.Drawing.Point(283, 75);
            this.Tb_ECPayAPIURL.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Tb_ECPayAPIURL.Name = "Tb_ECPayAPIURL";
            this.Tb_ECPayAPIURL.Size = new System.Drawing.Size(1050, 25);
            this.Tb_ECPayAPIURL.TabIndex = 3;
            // 
            // BtnEnterMonitor
            // 
            this.BtnEnterMonitor.Font = new System.Drawing.Font("Noto Sans TC", 9F);
            this.BtnEnterMonitor.Location = new System.Drawing.Point(12, 174);
            this.BtnEnterMonitor.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.BtnEnterMonitor.Name = "BtnEnterMonitor";
            this.BtnEnterMonitor.Size = new System.Drawing.Size(1327, 35);
            this.BtnEnterMonitor.TabIndex = 3;
            this.BtnEnterMonitor.Text = "進入監控頁面";
            this.BtnEnterMonitor.UseVisualStyleBackColor = true;
            this.BtnEnterMonitor.Click += new System.EventHandler(this.BtnEnterMonitor_Click);
            // 
            // Tb_OPayAPIURL
            // 
            this.Tb_OPayAPIURL.Font = new System.Drawing.Font("Noto Sans TC", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Tb_OPayAPIURL.Location = new System.Drawing.Point(283, 42);
            this.Tb_OPayAPIURL.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Tb_OPayAPIURL.Name = "Tb_OPayAPIURL";
            this.Tb_OPayAPIURL.Size = new System.Drawing.Size(1050, 25);
            this.Tb_OPayAPIURL.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Noto Sans TC", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(46, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(240, 27);
            this.label2.TabIndex = 4;
            this.label2.Text = "歐富寶贊助成功動畫網址：";
            // 
            // Tb_StreamlabsKey
            // 
            this.Tb_StreamlabsKey.Font = new System.Drawing.Font("Noto Sans TC", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Tb_StreamlabsKey.Location = new System.Drawing.Point(283, 11);
            this.Tb_StreamlabsKey.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Tb_StreamlabsKey.Name = "Tb_StreamlabsKey";
            this.Tb_StreamlabsKey.Size = new System.Drawing.Size(1050, 25);
            this.Tb_StreamlabsKey.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.label3.Location = new System.Drawing.Point(12, 11);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(274, 27);
            this.label3.TabIndex = 7;
            this.label3.Text = "Streamlabs Socket API 符記：";
            // 
            // Tb_HiveBeeAPIURL
            // 
            this.Tb_HiveBeeAPIURL.Font = new System.Drawing.Font("Noto Sans TC", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Tb_HiveBeeAPIURL.Location = new System.Drawing.Point(283, 108);
            this.Tb_HiveBeeAPIURL.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Tb_HiveBeeAPIURL.Name = "Tb_HiveBeeAPIURL";
            this.Tb_HiveBeeAPIURL.Size = new System.Drawing.Size(1050, 25);
            this.Tb_HiveBeeAPIURL.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Noto Sans TC", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(105, 106);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(181, 27);
            this.label4.TabIndex = 9;
            this.label4.Text = "HiveBee通知網址：";
            //
            // Tb_SoundAlertsURL
            //
            this.Tb_SoundAlertsURL.Font = new System.Drawing.Font("Noto Sans TC", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Tb_SoundAlertsURL.Location = new System.Drawing.Point(283, 141);
            this.Tb_SoundAlertsURL.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Tb_SoundAlertsURL.Name = "Tb_SoundAlertsURL";
            this.Tb_SoundAlertsURL.Size = new System.Drawing.Size(1050, 25);
            this.Tb_SoundAlertsURL.TabIndex = 5;
            //
            // label5
            //
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Noto Sans TC", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(46, 139);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(240, 27);
            this.label5.TabIndex = 10;
            this.label5.Text = "SoundAlerts Overlay網址：";
            //
            // Startup
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1345, 215);
            this.Controls.Add(this.Tb_SoundAlertsURL);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.Tb_HiveBeeAPIURL);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.Tb_StreamlabsKey);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Tb_OPayAPIURL);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.BtnEnterMonitor);
            this.Controls.Add(this.Tb_ECPayAPIURL);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Noto Sans TC", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Startup";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "啟動初始化";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Startup_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox Tb_ECPayAPIURL;
        private System.Windows.Forms.Button BtnEnterMonitor;
        private System.Windows.Forms.TextBox Tb_OPayAPIURL;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox Tb_StreamlabsKey;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox Tb_HiveBeeAPIURL;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox Tb_SoundAlertsURL;
        private System.Windows.Forms.Label label5;
    }
}

