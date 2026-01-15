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
            this.BtInitECPay = new System.Windows.Forms.Button();
            this.BtnEnterMonitor = new System.Windows.Forms.Button();
            this.BtInitOPay = new System.Windows.Forms.Button();
            this.Tb_OPayAPIURL = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.Tb_StreamlabsKey = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Noto Sans TC", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(30, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(221, 27);
            this.label1.TabIndex = 0;
            this.label1.Text = "綠界贊助成功動畫網址：";
            // 
            // Tb_ECPayAPIURL
            // 
            this.Tb_ECPayAPIURL.Font = new System.Drawing.Font("Noto Sans TC", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Tb_ECPayAPIURL.Location = new System.Drawing.Point(248, 13);
            this.Tb_ECPayAPIURL.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Tb_ECPayAPIURL.Name = "Tb_ECPayAPIURL";
            this.Tb_ECPayAPIURL.Size = new System.Drawing.Size(947, 25);
            this.Tb_ECPayAPIURL.TabIndex = 1;
            // 
            // BtInitECPay
            // 
            this.BtInitECPay.Font = new System.Drawing.Font("Noto Sans TC", 9F);
            this.BtInitECPay.Location = new System.Drawing.Point(1201, 7);
            this.BtInitECPay.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.BtInitECPay.Name = "BtInitECPay";
            this.BtInitECPay.Size = new System.Drawing.Size(138, 35);
            this.BtInitECPay.TabIndex = 2;
            this.BtInitECPay.Text = "初始化服務";
            this.BtInitECPay.UseVisualStyleBackColor = true;
            this.BtInitECPay.Click += new System.EventHandler(this.BtInitECPay_Click);
            // 
            // BtnEnterMonitor
            // 
            this.BtnEnterMonitor.Font = new System.Drawing.Font("Noto Sans TC", 9F);
            this.BtnEnterMonitor.Location = new System.Drawing.Point(12, 120);
            this.BtnEnterMonitor.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.BtnEnterMonitor.Name = "BtnEnterMonitor";
            this.BtnEnterMonitor.Size = new System.Drawing.Size(1327, 35);
            this.BtnEnterMonitor.TabIndex = 3;
            this.BtnEnterMonitor.Text = "進入監控頁面";
            this.BtnEnterMonitor.UseVisualStyleBackColor = true;
            this.BtnEnterMonitor.Click += new System.EventHandler(this.BtnEnterMonitor_Click);
            // 
            // BtInitOPay
            // 
            this.BtInitOPay.Font = new System.Drawing.Font("Noto Sans TC", 9F);
            this.BtInitOPay.Location = new System.Drawing.Point(1201, 46);
            this.BtInitOPay.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.BtInitOPay.Name = "BtInitOPay";
            this.BtInitOPay.Size = new System.Drawing.Size(138, 35);
            this.BtInitOPay.TabIndex = 6;
            this.BtInitOPay.Text = "初始化服務";
            this.BtInitOPay.UseVisualStyleBackColor = true;
            this.BtInitOPay.Click += new System.EventHandler(this.BtInitOPay_Click);
            // 
            // Tb_OPayAPIURL
            // 
            this.Tb_OPayAPIURL.Font = new System.Drawing.Font("Noto Sans TC", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Tb_OPayAPIURL.Location = new System.Drawing.Point(248, 52);
            this.Tb_OPayAPIURL.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Tb_OPayAPIURL.Name = "Tb_OPayAPIURL";
            this.Tb_OPayAPIURL.Size = new System.Drawing.Size(947, 25);
            this.Tb_OPayAPIURL.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Noto Sans TC", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(11, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(240, 27);
            this.label2.TabIndex = 4;
            this.label2.Text = "歐富寶贊助成功動畫網址：";
            // 
            // Tb_StreamlabsKey
            // 
            this.Tb_StreamlabsKey.Font = new System.Drawing.Font("Noto Sans TC", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Tb_StreamlabsKey.Location = new System.Drawing.Point(283, 87);
            this.Tb_StreamlabsKey.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Tb_StreamlabsKey.Name = "Tb_StreamlabsKey";
            this.Tb_StreamlabsKey.Size = new System.Drawing.Size(1050, 25);
            this.Tb_StreamlabsKey.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.label3.Location = new System.Drawing.Point(12, 85);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(274, 27);
            this.label3.TabIndex = 7;
            this.label3.Text = "Streamlabs Socket API 符記：";
            // 
            // Startup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1345, 163);
            this.Controls.Add(this.Tb_StreamlabsKey);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.BtInitOPay);
            this.Controls.Add(this.Tb_OPayAPIURL);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.BtnEnterMonitor);
            this.Controls.Add(this.BtInitECPay);
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
        private System.Windows.Forms.Button BtInitECPay;
        private System.Windows.Forms.Button BtnEnterMonitor;
        private System.Windows.Forms.Button BtInitOPay;
        private System.Windows.Forms.TextBox Tb_OPayAPIURL;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox Tb_StreamlabsKey;
        private System.Windows.Forms.Label label3;
    }
}

