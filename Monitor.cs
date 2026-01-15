using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DonateMonitor
{
    public partial class Monitor : Form
    {
        private readonly ConcurrentQueue<string> _logQueue = new ConcurrentQueue<string>();
        private readonly ConcurrentQueue<string> _obsQueue = new ConcurrentQueue<string>();
        private readonly System.Windows.Forms.Timer _logTimer = new System.Windows.Forms.Timer();
        private readonly CancellationTokenSource _ctsServices = new CancellationTokenSource();
        private readonly ServiceListener.ECPay _servicesECPAY = new ServiceListener.ECPay();
        private readonly ServiceListener.OPay _servicesOPAY = new ServiceListener.OPay();
        private readonly ServiceListener.Streamlabs _servicesStreamlabs = new ServiceListener.Streamlabs();
        private void PrependText(TextBoxBase tb, string text)
        {
            tb.SuspendLayout();

            tb.SelectionStart = 0;
            tb.SelectionLength = 0;
            tb.SelectedText = text;

            tb.ResumeLayout();
        }
        private void InitLogPump()
        {
            // cleanup
            string obsFileName = "obs.txt";
            File.WriteAllText(obsFileName, "");

            _logTimer.Interval = 50;
            _logTimer.Tick += (s, e) =>
            {
                while (_logQueue.TryDequeue(out var msg))
                    PrependText(Tb_MonitorOut, msg + Environment.NewLine);
                while (_obsQueue.TryDequeue(out var msg))
                {
                    if (Global.OBS_OutputMode == 1)
                        File.AppendAllText(obsFileName, msg + " ", Encoding.UTF8);
                    else
                        File.AppendAllText(obsFileName, msg + Environment.NewLine, Encoding.UTF8);
                    
                }
            };
            _logTimer.Start();
        }
        private void UnInitLogPump()
        {
            _logTimer.Stop();
        }
        public void AppendLogFromECPay(string name, string amount, string msg, bool isPreview = false)
        {
            AppendLog(0, name, amount, msg, "TWD", null, isPreview);
        }
        public void AppendLogFromOPay(string name, string amount, string msg, bool isPreview = false)
        {
            AppendLog(1, name, amount, msg, "TWD", null, isPreview);
        }
        public void AppendLogFromStreamlabs_Paypal(string name, string amount, string currency, string msg, bool isPreview = false)
        {
            AppendLog(2, name, amount, msg, currency, null, isPreview);
        }
        public void AppendLogFromStreamlabs_Bits(string acc, string amount, string msg, bool isPreview = false)
        {
            AppendLog(3, acc, amount, msg, Global.Custom_Bits, null, isPreview);
        }
        public void AppendLogFromStreamlabs_SubGift(string acc, string amount, string displayName, string subplan, bool isPreview = false)
        {
            AppendLog(4, acc, amount, displayName, Global.Custom_Sub_Gift, subplan, isPreview);
        }
        private void AppendLog(int nType, string name, string amount, string msg, string currency = "TWD", string subplan = null, bool isPreview = false)
        {
            if (string.IsNullOrEmpty(msg))
                msg = "";

            string type;
            string obsMsg;
            string displayName = null;
            if (nType == 4)
            {
                type = Global.Custom_Sub_Gift;
                obsMsg = Global.Streamlabs_SubGift_OBS_Msg;
                displayName = msg;
            }
            else if (nType == 3)
            {
                type = Global.Custom_Bits;
                obsMsg = Global.Streamlabs_Bits_OBS_Msg;
            }
            else if (nType == 2)
            {
                type = "Paypal(Streamlabs)";
                obsMsg = Global.Streamlabs_Paypal_OBS_Msg;
            }
            else if (nType == 1)
            {
                type = "歐富寶";
                obsMsg = Global.OPAY_OBS_Msg;
            }
            else
            {
                type = "綠界";
                obsMsg = Global.ECPAY_OBS_Msg;
            }

            var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            var now = DateTime.Now.ToString("MM-dd HH:mm:ss");
            string formated_amount = Global.FormatAmount(amount);
            string logLine = $"[{now}][{type}] ";
            string extLogLine;
            if (nType == 4)
                extLogLine = string.Format("{0}({1}) {2}{3}{4}", displayName, name, formated_amount, currency, (!string.IsNullOrEmpty(subplan) ? "(" + subplan + ")" : ""));
            else
                extLogLine = $"{name} 贊助了 {formated_amount}{currency}, 說: {msg}";

            string obsOutput;
            if (nType == 4)
                obsOutput = string.Format(obsMsg, displayName, formated_amount, currency, subplan);
            else
                obsOutput = string.Format(obsMsg, name, formated_amount, currency, subplan);

            if (isPreview)
            {
                MessageBox.Show(obsOutput);
                return;
            }

            _logQueue.Enqueue(logLine + extLogLine);
            _obsQueue.Enqueue(obsOutput);
            CsvLogger.Write(nowFull, type, name, decimal.Parse(amount), currency, extLogLine);
        }
        public void SetActiveECPay(bool bActive)
        {
            SafeUpdateUI(() =>
            {
                lbECPAY_Status.Text = $"綠界狀態：{(bActive ? "有效" : "無效")}";
            });
        }
        public void SetActiveOPay(bool bActive)
        {
            SafeUpdateUI(() =>
            {
                lbOPAY_Status.Text = $"歐富寶狀態：{(bActive ? "有效" : "無效")}";
            });
        }
        public void SetStreamlabs(bool bActive)
        {
            SafeUpdateUI(() =>
            {
                lbStreamlabs_Status.Text = $"Streamlabs 狀態：{(bActive ? "有效" : "無效")}";
            });
        }
        private void SafeUpdateUI(Action action)
        {
            if (IsDisposed) return;

            if (InvokeRequired) BeginInvoke(action);
            else action();
        }
        private void InitServices()
        {
            if (Global.IsEnableECPAY())
            {
                _ = _servicesECPAY.StartAsync(this, _ctsServices.Token);
            }
            if (Global.IsEnableOPAY())
            {
                _ = _servicesOPAY.StartAsync(this, _ctsServices.Token);
            }
            if (Global.IsEnableStreamlabs())
            {
                _ = _servicesStreamlabs.StartAsync(this, _ctsServices.Token);
            }
        }
        private async Task UninitServices()
        {
            _ctsServices.Cancel();
            await _servicesECPAY.StopAsync();
        }
        public Monitor()
        {
            InitializeComponent();
            InitLogPump();
            InitServices();
        }
        private async void Monitor_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;                 // 先擋關閉
            UnInitLogPump();
            await UninitServices();      // 等停止完成
            e.Cancel = false;                // 放行
            Application.Exit();
        }
        private void BtConfig_Click(object sender, EventArgs e)
        {
            new Config(this).ShowDialog();
        }
    }
}
