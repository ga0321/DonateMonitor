using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DonateMonitor
{
    public partial class Startup : Form
    {
        public Startup()
        {
            InitializeComponent();

            Tb_ECPayAPIURL.Text = Setting.Read(Setting.kECPAY_APIURL);
            Tb_OPayAPIURL.Text = Setting.Read(Setting.kOPAY_APIURL);
            Tb_StreamlabsKey.Text = Global.StreamlabsKey;
        }

        private static bool GetHttpResponse(string sApiUrl, out string osResult)
        {
            osResult = null;

            // 確保 TLS 1.2（非常重要）
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            using (var client = new HttpClient())
            {
                client.Timeout = System.Threading.Timeout.InfiniteTimeSpan;
                client.DefaultRequestHeaders.Add(
                    "User-Agent",
                    "DonateMonitor/v1.0"
                );

                var response = client.GetAsync(sApiUrl).Result;
                if (response.IsSuccessStatusCode)
                {
                    osResult = response.Content.ReadAsStringAsync().Result;
                    return true;
                }
            }

            return false;
        }

        #region ECPAY
        private static string GetECPayLoginToken(string text)
        {
            var match = Regex.Match(
                text,
                @"this\.loginToken\s*=\s*""([^""]+)"""
            );

            return match.Success ? match.Groups[1].Value : null;
        }
        private static bool InitECPay(string sApiUrl)
        {
            try
            {
                if (!GetHttpResponse(sApiUrl, out string sResult) && string.IsNullOrEmpty(sResult))
                {
                    Global.ShowError("綠界初始化失敗，請稍後再試");
                    return false;
                }
#if DEBUG
                Console.WriteLine(sResult);
#endif
                Global.ECPAY_LoginToken = GetECPayLoginToken(sResult);
                if (string.IsNullOrEmpty(Global.ECPAY_LoginToken))
                {
                    Global.ShowError("綠界初始化失敗，可能API參數已變更，請聯絡作者");
                    return false;
                }

                Global.ECPAY_ListenKey = new Uri(sApiUrl).Segments.Last();

#if DEBUG
                Console.WriteLine($"sECPAY_LoginToken: {Global.ECPAY_LoginToken}");
                Console.WriteLine($"sECPAY_ListenKey: {Global.ECPAY_ListenKey}");
#endif
                return true;
            }
            catch (Exception e)
            {
                Global.ShowError($"綠界初始化失敗: {e.Message}", true);
            }
            return false;
        }
        #endregion

        #region OPAY
        private static bool InitOPay(string sApiUrl)
        {
            try
            {
                if (!GetHttpResponse(sApiUrl, out string sResult) && string.IsNullOrEmpty(sResult))
                {
                    Global.ShowError("歐富寶初始化失敗，請稍後再試");
                    return false;
                }
#if DEBUG
                Console.WriteLine(sResult);
#endif

                Global.OPAY_ListenKey = new Uri(sApiUrl).Segments.Last();

#if DEBUG
                Console.WriteLine($"sOPAY_ListenKey: {Global.OPAY_ListenKey}");
#endif
                return true;
            }
            catch (Exception e)
            {
                Global.ShowError($"歐富寶初始化失敗: {e.Message}", true);
            }
            return false;
        }
        #endregion

        #region Streamlabs
        private void SaveStreamlabsKey()
        {
            var tb = Tb_StreamlabsKey.Text;
            if (!string.IsNullOrEmpty(tb))
            {
                Setting.Save(Setting.kSTREAMLABS_KEY, tb);
                Global.StreamlabsKey = tb;
            }
        }
        #endregion

        private void BtInitECPay_Click(object sender, EventArgs e)
        {
            var tb = Tb_ECPayAPIURL;
            string sApiUrl = tb.Text;
            if (string.IsNullOrEmpty(sApiUrl))
            {
                Global.ShowError("請輸入網址");
                return;
            }

            Setting.Save(Setting.kECPAY_APIURL, sApiUrl);
            if (InitECPay(sApiUrl))
            {
                var btn = (Button)sender;
                btn.Enabled = false;
                btn.Text = "成 功";
                tb.Enabled = false;
            }
        }

        private void BtInitOPay_Click(object sender, EventArgs e)
        {
            var tb = Tb_OPayAPIURL;
            string sApiUrl = tb.Text;
            if (string.IsNullOrEmpty(sApiUrl))
            {
                Global.ShowError("請輸入網址");
                return;
            }

            Setting.Save(Setting.kOPAY_APIURL, sApiUrl);
            if (InitOPay(sApiUrl))
            {
                var btn = (Button)sender;
                btn.Enabled = false;
                btn.Text = "成 功";
                tb.Enabled = false;
            }
        }

        private void BtnEnterMonitor_Click(object sender, EventArgs e)
        {
            SaveStreamlabsKey();
            if (!Global.IsEnableAnyService())
            {
                Global.ShowError("請至少初始化一種服務");
                return;
            }
            if (!Global.IsEnableAllService() && MessageBox.Show("有服務尚未初始化，是否繼續？", "DonateMonitor", MessageBoxButtons.YesNo) == DialogResult.No)
                return;
            this.Close();
            Global._bExit = false;
        }

        private void Startup_FormClosing(object sender, FormClosingEventArgs e)
        {
            Global._bExit = true;
        }
    }
}
