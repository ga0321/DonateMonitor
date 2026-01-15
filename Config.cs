using DonateMonitor.ServiceListener;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static DonateMonitor.Global;

namespace DonateMonitor
{
    public partial class Config : Form
    {
        readonly Monitor _monitor = null;
        Global.VARS _VARS = new Global.VARS();
        public Config(Monitor monitor)
        {
            InitializeComponent();
            _monitor = monitor;
            _VARS = Global._VARS;
            LoadConfig();
        }

        private void LoadConfig()
        {
            if (Global.OBS_OutputMode == 1)
            {
                RBt_ObsOutputMode_NextLine.Checked = false;
                RBt_ObsOutputMode_Single.Checked = true;
            }
            else
            {
                RBt_ObsOutputMode_NextLine.Checked = true;
                RBt_ObsOutputMode_Single.Checked = false;
            }

            Tb_Msg_ECPay_Msg.Text = Global.ECPAY_OBS_Msg;
            Tb_Msg_OPay_Msg.Text = Global.OPAY_OBS_Msg;
            Tb_Msg_Streamlabs_Paypal_Msg.Text = Global.Streamlabs_Paypal_OBS_Msg;
            Tb_Msg_Streamlabs_SubGift_Msg.Text = Global.Streamlabs_SubGift_OBS_Msg;
            Tb_Msg_Streamlabs_Bits_Msg.Text = Global.Streamlabs_Bits_OBS_Msg;
            Tb_Msg_Custom_Anon.Text = Global.Custom_ANON;
            Tb_Msg_Custom_Sub_Gift.Text = Global.Custom_Sub_Gift;
            Tb_Msg_Custom_Bits.Text = Global.Custom_Bits;
            Tb_Msg_Custom_Sub_Tier1.Text = Global.Custom_Sub_Tier1;
            Tb_Msg_Custom_Sub_Tier2.Text = Global.Custom_Sub_Tier2;
            Tb_Msg_Custom_Sub_Tier3.Text = Global.Custom_Sub_Tier3;
        }

        private void SaveConfig()
        {
            if (RBt_ObsOutputMode_Single.Checked)
                Global.OBS_OutputMode = 1;
            else
                Global.OBS_OutputMode = 0;

            Global.ECPAY_OBS_Msg = Tb_Msg_ECPay_Msg.Text;
            Global.OPAY_OBS_Msg = Tb_Msg_OPay_Msg.Text;
            Global.Streamlabs_Paypal_OBS_Msg = Tb_Msg_Streamlabs_Paypal_Msg.Text;
            Global.Streamlabs_SubGift_OBS_Msg = Tb_Msg_Streamlabs_SubGift_Msg.Text;
            Global.Streamlabs_Bits_OBS_Msg = Tb_Msg_Streamlabs_Bits_Msg.Text;
            Global.Custom_ANON = Tb_Msg_Custom_Anon.Text;
            Global.Custom_Sub_Gift = Tb_Msg_Custom_Sub_Gift.Text;
            Global.Custom_Bits = Tb_Msg_Custom_Bits.Text;
            Global.Custom_Sub_Tier1 = Tb_Msg_Custom_Sub_Tier1.Text;
            Global.Custom_Sub_Tier2 = Tb_Msg_Custom_Sub_Tier2.Text;
            Global.Custom_Sub_Tier3 = Tb_Msg_Custom_Sub_Tier3.Text;
        }

        private void SaveSettings()
        {
            _VARS = Global._VARS;
            SaveConfig();
            Global.SaveSettings();
        }
        private void RestoreSettings()
        {
            Global._VARS = _VARS;
            Global.SaveSettings();
        }
        private void Bt_Save_Click(object sender, EventArgs e)
        {
            SaveSettings();
            Close();
        }

        private void BtReset_Click(object sender, EventArgs e)
        {
            var kECPAY_APIURL = Setting.Read(Setting.kECPAY_APIURL);
            var kOPAY_APIURL = Setting.Read(Setting.kOPAY_APIURL);
            var kSTREAMLABS_KEY = Setting.Read(Setting.kSTREAMLABS_KEY);

            Setting.Reset();
            Setting.Save(Setting.kECPAY_APIURL, kECPAY_APIURL);
            Setting.Save(Setting.kOPAY_APIURL, kOPAY_APIURL);
            Setting.Save(Setting.kSTREAMLABS_KEY, kSTREAMLABS_KEY);
            Global.InitSettings();
            Global.LoadSettings();
            LoadConfig();
        }

        private void BtPreview_ECPay_Click(object sender, EventArgs e)
        {
            SaveSettings();
            _monitor.AppendLogFromECPay("測試綠界", "100", "測試綠界訊息", true);
            RestoreSettings();
        }

        private void BtPreview_OPay_Click(object sender, EventArgs e)
        {
            SaveSettings();
            _monitor.AppendLogFromOPay("測試歐富寶", "100", "測試歐富寶訊息", true);
            RestoreSettings();
        }

        private void BtPreview_Streamlabs_Paypal_Click(object sender, EventArgs e)
        {
            SaveSettings();
            _monitor.AppendLogFromStreamlabs_Paypal("測試Paypal", "100", "TWD", "測試Paypal訊息", true);
            RestoreSettings();
        }

        private void BtPreview_Streamlabs_SubGift_Click(object sender, EventArgs e)
        {
            SaveSettings();
            _monitor.AppendLogFromStreamlabs_SubGift("test", "10", "測試", Streamlabs.SubPlanToText("1000"), true);
            RestoreSettings();
        }

        private void BtPreview_Streamlabs_Bits_Click(object sender, EventArgs e)
        {
            SaveSettings();
            _monitor.AppendLogFromStreamlabs_Bits("test", "100", "測試小奇點訊息", true);
            RestoreSettings();
        }
    }
}
