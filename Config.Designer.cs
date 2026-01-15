namespace DonateMonitor
{
    partial class Config
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
            this.label1 = new System.Windows.Forms.Label();
            this.RBt_ObsOutputMode_Single = new System.Windows.Forms.RadioButton();
            this.RBt_ObsOutputMode_NextLine = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.Tb_Msg_ECPay_Msg = new System.Windows.Forms.TextBox();
            this.Bt_Save = new System.Windows.Forms.Button();
            this.Tb_Msg_OPay_Msg = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.Tb_Msg_Streamlabs_Paypal_Msg = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.Tb_Msg_Streamlabs_SubGift_Msg = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.Tb_Msg_Streamlabs_Bits_Msg = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.Tb_Msg_Custom_Anon = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.Tb_Msg_Custom_Sub_Gift = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.Tb_Msg_Custom_Bits = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.Tb_Msg_Custom_Sub_Tier1 = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.Tb_Msg_Custom_Sub_Tier2 = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.Tb_Msg_Custom_Sub_Tier3 = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.BtReset = new System.Windows.Forms.Button();
            this.BtPreview_ECPay = new System.Windows.Forms.Button();
            this.BtPreview_OPay = new System.Windows.Forms.Button();
            this.BtPreview_Streamlabs_Paypal = new System.Windows.Forms.Button();
            this.BtPreview_Streamlabs_SubGift = new System.Windows.Forms.Button();
            this.BtPreview_Streamlabs_Bits = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Noto Sans TC", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(19, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(182, 27);
            this.label1.TabIndex = 0;
            this.label1.Text = "OBS檔案輸出模式：";
            // 
            // RBt_ObsOutputMode_Single
            // 
            this.RBt_ObsOutputMode_Single.AutoSize = true;
            this.RBt_ObsOutputMode_Single.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.RBt_ObsOutputMode_Single.Location = new System.Drawing.Point(319, 24);
            this.RBt_ObsOutputMode_Single.Name = "RBt_ObsOutputMode_Single";
            this.RBt_ObsOutputMode_Single.Size = new System.Drawing.Size(256, 31);
            this.RBt_ObsOutputMode_Single.TabIndex = 2;
            this.RBt_ObsOutputMode_Single.TabStop = true;
            this.RBt_ObsOutputMode_Single.Text = "單行模式 (A: 100T B: 100T)";
            this.RBt_ObsOutputMode_Single.UseVisualStyleBackColor = true;
            // 
            // RBt_ObsOutputMode_NextLine
            // 
            this.RBt_ObsOutputMode_NextLine.AutoSize = true;
            this.RBt_ObsOutputMode_NextLine.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.RBt_ObsOutputMode_NextLine.Location = new System.Drawing.Point(207, 24);
            this.RBt_ObsOutputMode_NextLine.Name = "RBt_ObsOutputMode_NextLine";
            this.RBt_ObsOutputMode_NextLine.Size = new System.Drawing.Size(106, 31);
            this.RBt_ObsOutputMode_NextLine.TabIndex = 1;
            this.RBt_ObsOutputMode_NextLine.TabStop = true;
            this.RBt_ObsOutputMode_NextLine.Text = "換行模式";
            this.RBt_ObsOutputMode_NextLine.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.RBt_ObsOutputMode_Single);
            this.groupBox1.Controls.Add(this.RBt_ObsOutputMode_NextLine);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(668, 70);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "普通設定";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.BtPreview_Streamlabs_Bits);
            this.groupBox2.Controls.Add(this.BtPreview_Streamlabs_SubGift);
            this.groupBox2.Controls.Add(this.BtPreview_Streamlabs_Paypal);
            this.groupBox2.Controls.Add(this.BtPreview_OPay);
            this.groupBox2.Controls.Add(this.BtPreview_ECPay);
            this.groupBox2.Controls.Add(this.Tb_Msg_Streamlabs_Bits_Msg);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.Tb_Msg_Streamlabs_SubGift_Msg);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.Tb_Msg_Streamlabs_Paypal_Msg);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.Tb_Msg_OPay_Msg);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.Tb_Msg_ECPay_Msg);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Location = new System.Drawing.Point(13, 88);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(667, 240);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "OBS輸出設定";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Noto Sans TC", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(37, 30);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(107, 27);
            this.label2.TabIndex = 3;
            this.label2.Text = "綠界輸出：";
            // 
            // Tb_Msg_ECPay_Msg
            // 
            this.Tb_Msg_ECPay_Msg.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Tb_Msg_ECPay_Msg.Location = new System.Drawing.Point(141, 27);
            this.Tb_Msg_ECPay_Msg.Name = "Tb_Msg_ECPay_Msg";
            this.Tb_Msg_ECPay_Msg.Size = new System.Drawing.Size(451, 35);
            this.Tb_Msg_ECPay_Msg.TabIndex = 4;
            // 
            // Bt_Save
            // 
            this.Bt_Save.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Bt_Save.Location = new System.Drawing.Point(13, 544);
            this.Bt_Save.Name = "Bt_Save";
            this.Bt_Save.Size = new System.Drawing.Size(667, 46);
            this.Bt_Save.TabIndex = 5;
            this.Bt_Save.Text = "保 存";
            this.Bt_Save.UseVisualStyleBackColor = true;
            this.Bt_Save.Click += new System.EventHandler(this.Bt_Save_Click);
            // 
            // Tb_Msg_OPay_Msg
            // 
            this.Tb_Msg_OPay_Msg.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Tb_Msg_OPay_Msg.Location = new System.Drawing.Point(141, 68);
            this.Tb_Msg_OPay_Msg.Name = "Tb_Msg_OPay_Msg";
            this.Tb_Msg_OPay_Msg.Size = new System.Drawing.Size(451, 35);
            this.Tb_Msg_OPay_Msg.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Noto Sans TC", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(18, 71);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(126, 27);
            this.label3.TabIndex = 5;
            this.label3.Text = "歐富寶輸出：";
            // 
            // Tb_Msg_Streamlabs_Paypal_Msg
            // 
            this.Tb_Msg_Streamlabs_Paypal_Msg.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Tb_Msg_Streamlabs_Paypal_Msg.Location = new System.Drawing.Point(141, 109);
            this.Tb_Msg_Streamlabs_Paypal_Msg.Name = "Tb_Msg_Streamlabs_Paypal_Msg";
            this.Tb_Msg_Streamlabs_Paypal_Msg.Size = new System.Drawing.Size(451, 35);
            this.Tb_Msg_Streamlabs_Paypal_Msg.TabIndex = 8;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Noto Sans TC", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(11, 112);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(133, 27);
            this.label4.TabIndex = 7;
            this.label4.Text = "Paypal 輸出：";
            // 
            // Tb_Msg_Streamlabs_SubGift_Msg
            // 
            this.Tb_Msg_Streamlabs_SubGift_Msg.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Tb_Msg_Streamlabs_SubGift_Msg.Location = new System.Drawing.Point(141, 150);
            this.Tb_Msg_Streamlabs_SubGift_Msg.Name = "Tb_Msg_Streamlabs_SubGift_Msg";
            this.Tb_Msg_Streamlabs_SubGift_Msg.Size = new System.Drawing.Size(451, 35);
            this.Tb_Msg_Streamlabs_SubGift_Msg.TabIndex = 10;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Noto Sans TC", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(37, 153);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(107, 27);
            this.label5.TabIndex = 9;
            this.label5.Text = "贈訂輸出：";
            // 
            // Tb_Msg_Streamlabs_Bits_Msg
            // 
            this.Tb_Msg_Streamlabs_Bits_Msg.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Tb_Msg_Streamlabs_Bits_Msg.Location = new System.Drawing.Point(141, 191);
            this.Tb_Msg_Streamlabs_Bits_Msg.Name = "Tb_Msg_Streamlabs_Bits_Msg";
            this.Tb_Msg_Streamlabs_Bits_Msg.Size = new System.Drawing.Size(451, 35);
            this.Tb_Msg_Streamlabs_Bits_Msg.TabIndex = 12;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Noto Sans TC", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label6.Location = new System.Drawing.Point(18, 194);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(126, 27);
            this.label6.TabIndex = 11;
            this.label6.Text = "小奇點輸出：";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.Tb_Msg_Custom_Sub_Tier3);
            this.groupBox3.Controls.Add(this.label12);
            this.groupBox3.Controls.Add(this.Tb_Msg_Custom_Sub_Tier2);
            this.groupBox3.Controls.Add(this.label11);
            this.groupBox3.Controls.Add(this.Tb_Msg_Custom_Sub_Tier1);
            this.groupBox3.Controls.Add(this.label10);
            this.groupBox3.Controls.Add(this.Tb_Msg_Custom_Bits);
            this.groupBox3.Controls.Add(this.label9);
            this.groupBox3.Controls.Add(this.Tb_Msg_Custom_Sub_Gift);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.Tb_Msg_Custom_Anon);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Location = new System.Drawing.Point(13, 335);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(667, 203);
            this.groupBox3.TabIndex = 6;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "別稱設定";
            // 
            // Tb_Msg_Custom_Anon
            // 
            this.Tb_Msg_Custom_Anon.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Tb_Msg_Custom_Anon.Location = new System.Drawing.Point(73, 25);
            this.Tb_Msg_Custom_Anon.Name = "Tb_Msg_Custom_Anon";
            this.Tb_Msg_Custom_Anon.Size = new System.Drawing.Size(256, 35);
            this.Tb_Msg_Custom_Anon.TabIndex = 14;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Noto Sans TC", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label7.Location = new System.Drawing.Point(6, 28);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(69, 27);
            this.label7.TabIndex = 13;
            this.label7.Text = "匿名：";
            // 
            // Tb_Msg_Custom_Sub_Gift
            // 
            this.Tb_Msg_Custom_Sub_Gift.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Tb_Msg_Custom_Sub_Gift.Location = new System.Drawing.Point(435, 25);
            this.Tb_Msg_Custom_Sub_Gift.Name = "Tb_Msg_Custom_Sub_Gift";
            this.Tb_Msg_Custom_Sub_Gift.Size = new System.Drawing.Size(226, 35);
            this.Tb_Msg_Custom_Sub_Gift.TabIndex = 16;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Noto Sans TC", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label8.Location = new System.Drawing.Point(331, 28);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(107, 27);
            this.label8.TabIndex = 15;
            this.label8.Text = "贈禮訂閱：";
            // 
            // Tb_Msg_Custom_Bits
            // 
            this.Tb_Msg_Custom_Bits.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Tb_Msg_Custom_Bits.Location = new System.Drawing.Point(90, 66);
            this.Tb_Msg_Custom_Bits.Name = "Tb_Msg_Custom_Bits";
            this.Tb_Msg_Custom_Bits.Size = new System.Drawing.Size(239, 35);
            this.Tb_Msg_Custom_Bits.TabIndex = 18;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Noto Sans TC", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label9.Location = new System.Drawing.Point(6, 69);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(88, 27);
            this.label9.TabIndex = 17;
            this.label9.Text = "小奇點：";
            // 
            // Tb_Msg_Custom_Sub_Tier1
            // 
            this.Tb_Msg_Custom_Sub_Tier1.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Tb_Msg_Custom_Sub_Tier1.Location = new System.Drawing.Point(435, 66);
            this.Tb_Msg_Custom_Sub_Tier1.Name = "Tb_Msg_Custom_Sub_Tier1";
            this.Tb_Msg_Custom_Sub_Tier1.Size = new System.Drawing.Size(226, 35);
            this.Tb_Msg_Custom_Sub_Tier1.TabIndex = 20;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Noto Sans TC", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label10.Location = new System.Drawing.Point(350, 69);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(88, 27);
            this.label10.TabIndex = 19;
            this.label10.Text = "層級一：";
            // 
            // Tb_Msg_Custom_Sub_Tier2
            // 
            this.Tb_Msg_Custom_Sub_Tier2.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Tb_Msg_Custom_Sub_Tier2.Location = new System.Drawing.Point(435, 107);
            this.Tb_Msg_Custom_Sub_Tier2.Name = "Tb_Msg_Custom_Sub_Tier2";
            this.Tb_Msg_Custom_Sub_Tier2.Size = new System.Drawing.Size(226, 35);
            this.Tb_Msg_Custom_Sub_Tier2.TabIndex = 22;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Noto Sans TC", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label11.Location = new System.Drawing.Point(350, 110);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(88, 27);
            this.label11.TabIndex = 21;
            this.label11.Text = "層級二：";
            // 
            // Tb_Msg_Custom_Sub_Tier3
            // 
            this.Tb_Msg_Custom_Sub_Tier3.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.Tb_Msg_Custom_Sub_Tier3.Location = new System.Drawing.Point(435, 148);
            this.Tb_Msg_Custom_Sub_Tier3.Name = "Tb_Msg_Custom_Sub_Tier3";
            this.Tb_Msg_Custom_Sub_Tier3.Size = new System.Drawing.Size(226, 35);
            this.Tb_Msg_Custom_Sub_Tier3.TabIndex = 24;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Noto Sans TC", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label12.Location = new System.Drawing.Point(350, 151);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(88, 27);
            this.label12.TabIndex = 23;
            this.label12.Text = "層級三：";
            // 
            // BtReset
            // 
            this.BtReset.Font = new System.Drawing.Font("Noto Sans TC", 14F);
            this.BtReset.Location = new System.Drawing.Point(13, 596);
            this.BtReset.Name = "BtReset";
            this.BtReset.Size = new System.Drawing.Size(667, 46);
            this.BtReset.TabIndex = 7;
            this.BtReset.Text = "重置設定";
            this.BtReset.UseVisualStyleBackColor = true;
            this.BtReset.Click += new System.EventHandler(this.BtReset_Click);
            // 
            // BtPreview_ECPay
            // 
            this.BtPreview_ECPay.Font = new System.Drawing.Font("Noto Sans TC", 12F);
            this.BtPreview_ECPay.Location = new System.Drawing.Point(598, 27);
            this.BtPreview_ECPay.Name = "BtPreview_ECPay";
            this.BtPreview_ECPay.Size = new System.Drawing.Size(62, 35);
            this.BtPreview_ECPay.TabIndex = 13;
            this.BtPreview_ECPay.Text = "預覽";
            this.BtPreview_ECPay.UseVisualStyleBackColor = true;
            this.BtPreview_ECPay.Click += new System.EventHandler(this.BtPreview_ECPay_Click);
            // 
            // BtPreview_OPay
            // 
            this.BtPreview_OPay.Font = new System.Drawing.Font("Noto Sans TC", 12F);
            this.BtPreview_OPay.Location = new System.Drawing.Point(598, 69);
            this.BtPreview_OPay.Name = "BtPreview_OPay";
            this.BtPreview_OPay.Size = new System.Drawing.Size(62, 35);
            this.BtPreview_OPay.TabIndex = 14;
            this.BtPreview_OPay.Text = "預覽";
            this.BtPreview_OPay.UseVisualStyleBackColor = true;
            this.BtPreview_OPay.Click += new System.EventHandler(this.BtPreview_OPay_Click);
            // 
            // BtPreview_Streamlabs_Paypal
            // 
            this.BtPreview_Streamlabs_Paypal.Font = new System.Drawing.Font("Noto Sans TC", 12F);
            this.BtPreview_Streamlabs_Paypal.Location = new System.Drawing.Point(599, 109);
            this.BtPreview_Streamlabs_Paypal.Name = "BtPreview_Streamlabs_Paypal";
            this.BtPreview_Streamlabs_Paypal.Size = new System.Drawing.Size(62, 35);
            this.BtPreview_Streamlabs_Paypal.TabIndex = 15;
            this.BtPreview_Streamlabs_Paypal.Text = "預覽";
            this.BtPreview_Streamlabs_Paypal.UseVisualStyleBackColor = true;
            this.BtPreview_Streamlabs_Paypal.Click += new System.EventHandler(this.BtPreview_Streamlabs_Paypal_Click);
            // 
            // BtPreview_Streamlabs_SubGift
            // 
            this.BtPreview_Streamlabs_SubGift.Font = new System.Drawing.Font("Noto Sans TC", 12F);
            this.BtPreview_Streamlabs_SubGift.Location = new System.Drawing.Point(598, 150);
            this.BtPreview_Streamlabs_SubGift.Name = "BtPreview_Streamlabs_SubGift";
            this.BtPreview_Streamlabs_SubGift.Size = new System.Drawing.Size(62, 35);
            this.BtPreview_Streamlabs_SubGift.TabIndex = 16;
            this.BtPreview_Streamlabs_SubGift.Text = "預覽";
            this.BtPreview_Streamlabs_SubGift.UseVisualStyleBackColor = true;
            this.BtPreview_Streamlabs_SubGift.Click += new System.EventHandler(this.BtPreview_Streamlabs_SubGift_Click);
            // 
            // BtPreview_Streamlabs_Bits
            // 
            this.BtPreview_Streamlabs_Bits.Font = new System.Drawing.Font("Noto Sans TC", 12F);
            this.BtPreview_Streamlabs_Bits.Location = new System.Drawing.Point(598, 191);
            this.BtPreview_Streamlabs_Bits.Name = "BtPreview_Streamlabs_Bits";
            this.BtPreview_Streamlabs_Bits.Size = new System.Drawing.Size(62, 35);
            this.BtPreview_Streamlabs_Bits.TabIndex = 17;
            this.BtPreview_Streamlabs_Bits.Text = "預覽";
            this.BtPreview_Streamlabs_Bits.UseVisualStyleBackColor = true;
            this.BtPreview_Streamlabs_Bits.Click += new System.EventHandler(this.BtPreview_Streamlabs_Bits_Click);
            // 
            // Config
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(692, 647);
            this.Controls.Add(this.BtReset);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.Bt_Save);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Noto Sans TC", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "Config";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Config";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton RBt_ObsOutputMode_Single;
        private System.Windows.Forms.RadioButton RBt_ObsOutputMode_NextLine;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox Tb_Msg_ECPay_Msg;
        private System.Windows.Forms.Button Bt_Save;
        private System.Windows.Forms.TextBox Tb_Msg_OPay_Msg;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox Tb_Msg_Streamlabs_Paypal_Msg;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox Tb_Msg_Streamlabs_SubGift_Msg;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox Tb_Msg_Streamlabs_Bits_Msg;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox Tb_Msg_Custom_Anon;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox Tb_Msg_Custom_Sub_Gift;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox Tb_Msg_Custom_Bits;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox Tb_Msg_Custom_Sub_Tier3;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox Tb_Msg_Custom_Sub_Tier2;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox Tb_Msg_Custom_Sub_Tier1;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button BtReset;
        private System.Windows.Forms.Button BtPreview_ECPay;
        private System.Windows.Forms.Button BtPreview_OPay;
        private System.Windows.Forms.Button BtPreview_Streamlabs_Paypal;
        private System.Windows.Forms.Button BtPreview_Streamlabs_SubGift;
        private System.Windows.Forms.Button BtPreview_Streamlabs_Bits;
    }
}