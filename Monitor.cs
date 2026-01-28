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
        private readonly ConcurrentDictionary<string, string> _obsDict = new ConcurrentDictionary<string, string>();
        private readonly System.Windows.Forms.Timer _logTimer = new System.Windows.Forms.Timer();
        private readonly CancellationTokenSource _ctsServices = new CancellationTokenSource();
        private readonly ServiceListener.ECPay _servicesECPAY = new ServiceListener.ECPay();
        private readonly ServiceListener.OPay _servicesOPAY = new ServiceListener.OPay();
        private readonly ServiceListener.Streamlabs _servicesStreamlabs = new ServiceListener.Streamlabs();
        private readonly ServiceListener.HiveBee _servicesHiveBee = new ServiceListener.HiveBee();
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
            string obsFileName = "obs.txt";

            _logTimer.Interval = 50;
            _logTimer.Tick += (s, e) =>
            {
                while (_logQueue.TryDequeue(out var msg))
                    PrependText(Tb_MonitorOut, msg + Environment.NewLine);

                // 將所有帳號的累計記錄寫入 obs.txt (每個帳號+類型只保留最新一筆)
                if (_obsDict.Count > 0)
                {
                    var lines = _obsDict.Values.ToArray();
                    if (Global.OBS_OutputMode == 1)
                        File.WriteAllText(obsFileName, string.Join(" ", lines), Encoding.UTF8);
                    else
                        File.WriteAllText(obsFileName, string.Join(Environment.NewLine, lines), Encoding.UTF8);
                }
            };
            _logTimer.Start();
        }
        private void UnInitLogPump()
        {
            _logTimer.Stop();
        }
        private void LoadCumulativeDataFromDB()
        {
            var data = DonateDB.GetAllCumulativeData();
            foreach (var item in data)
            {
                // 檢查是否要輸出 (新訂閱/續訂要檢查設定)
                if (item.Type == Global.Type_Sub && !Global.EnableSubOutput)
                    continue;
                if (item.Type == Global.Type_Resub && !Global.EnableResubOutput)
                    continue;

                string obsKey = $"{item.Account}|{item.Type}|{item.SubPlan}";
                string formated_amount = Global.FormatAmount(item.TotalAmount.ToString());

                string obsOutput;
                if (item.Type == Global.Custom_Sub_Gift)
                {
                    // 贈訂格式: {displayName}: {amount}{currency}({subplan})
                    obsOutput = string.Format(Global.Streamlabs_SubGift_OBS_Msg, item.DisplayName, formated_amount, item.Currency, item.SubPlan);
                }
                else if (item.Type == Global.Type_Resub)
                {
                    // 續訂格式: {displayName} 續訂{months}個月({subplan})
                    obsOutput = string.Format(Global.Streamlabs_Resub_OBS_Msg, item.DisplayName, formated_amount, item.SubPlan);
                }
                else if (item.Type == Global.Type_Sub)
                {
                    // 新訂閱格式: {displayName} 新訂閱{months}個月({subplan})
                    obsOutput = string.Format(Global.Streamlabs_Sub_OBS_Msg, item.DisplayName, formated_amount, item.SubPlan);
                }
                else if (item.Type == Global.Custom_Bits)
                {
                    // 小奇點格式
                    obsOutput = string.Format(Global.Streamlabs_Bits_OBS_Msg, item.DisplayName, formated_amount, item.Currency, item.SubPlan);
                }
                else if (item.Type == Global.Type_ECPay)
                {
                    obsOutput = string.Format(Global.ECPAY_OBS_Msg, item.Account, formated_amount, item.Currency, item.SubPlan);
                }
                else if (item.Type == Global.Type_OPay)
                {
                    obsOutput = string.Format(Global.OPAY_OBS_Msg, item.Account, formated_amount, item.Currency, item.SubPlan);
                }
                else if (item.Type == Global.Type_HiveBee)
                {
                    obsOutput = string.Format(Global.HIVEBEE_OBS_Msg, item.Account, formated_amount, item.Currency, item.SubPlan);
                }
                else if (item.Type == Global.Type_Paypal)
                {
                    obsOutput = string.Format(Global.Streamlabs_Paypal_OBS_Msg, item.Account, formated_amount, item.Currency, item.SubPlan);
                }
                else
                {
                    // 其他類型使用通用格式
                    obsOutput = $"{item.Account}: {formated_amount}{item.Currency}";
                }

                _obsDict[obsKey] = obsOutput;
            }
        }
        public void AppendLogFromECPay(string name, string amount, string msg, bool isPreview = false)
        {
            string type = Global.Type_ECPay;
            decimal donateAmount = 0;
            decimal.TryParse(amount, out donateAmount);

            // 預覽時不寫入 DB
            if (!isPreview)
            {
                var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                DonateDB.Write(nowFull, type, name, name, donateAmount, "TWD", msg, null);
            }

            // 從 DB 計算累計金額
            decimal totalAmount = isPreview ? donateAmount : DonateDB.GetTotalAmount(name, type);

            // 使用累計金額輸出
            AppendLog(0, name, totalAmount.ToString(), msg, "TWD", null, isPreview);
        }
        public void AppendLogFromOPay(string name, string amount, string msg, bool isPreview = false)
        {
            string type = Global.Type_OPay;
            decimal donateAmount = 0;
            decimal.TryParse(amount, out donateAmount);

            // 預覽時不寫入 DB
            if (!isPreview)
            {
                var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                DonateDB.Write(nowFull, type, name, name, donateAmount, "TWD", msg, null);
            }

            // 從 DB 計算累計金額
            decimal totalAmount = isPreview ? donateAmount : DonateDB.GetTotalAmount(name, type);

            // 使用累計金額輸出
            AppendLog(1, name, totalAmount.ToString(), msg, "TWD", null, isPreview);
        }
        public void AppendLogFromStreamlabs_Paypal(string name, string amount, string currency, string msg, bool isPreview = false)
        {
            string type = Global.Type_Paypal;
            decimal donateAmount = 0;
            decimal.TryParse(amount, out donateAmount);

            // 預覽時不寫入 DB
            if (!isPreview)
            {
                var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                DonateDB.Write(nowFull, type, name, name, donateAmount, currency, msg, null);
            }

            // 從 DB 計算累計金額
            decimal totalAmount = isPreview ? donateAmount : DonateDB.GetTotalAmount(name, type);

            // 使用累計金額輸出
            AppendLog(2, name, totalAmount.ToString(), msg, currency, null, isPreview);
        }
        public void AppendLogFromStreamlabs_Bits(string acc, string name, string amount, string msg, bool isPreview = false)
        {
            string type = Global.Custom_Bits;
            decimal donateAmount = 0;
            decimal.TryParse(amount, out donateAmount);

            // 預覽時不寫入 DB
            if (!isPreview)
            {
                var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                DonateDB.Write(nowFull, type, acc, name, donateAmount, type, msg, null);
            }

            // 從 DB 計算累計金額 (使用 acc 帳號)
            decimal totalAmount = isPreview ? donateAmount : DonateDB.GetTotalAmount(acc, type);

            // 使用累計金額輸出
            AppendLog(3, name, totalAmount.ToString(), msg, type, null, isPreview, acc);
        }
        public void AppendLogFromStreamlabs_SubGift(string acc, string amount, string displayName, string subplan, bool isPreview = false)
        {
            // 累加 subgift 並獲取累計後的數量
            int giftAmount = 1;
            int.TryParse(amount, out giftAmount);
            if (giftAmount <= 0) giftAmount = 1;

            // 預覽時不寫入 DB
            if (!isPreview)
            {
                var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                DonateDB.Write(nowFull, Global.Custom_Sub_Gift, acc, displayName, giftAmount, Global.Custom_Sub_Gift, "", subplan);
            }

            // 從 DB 計算該帳號+層級的累計數量
            int totalCount = isPreview ? giftAmount : DonateDB.GetSubGiftCountByPlan(acc, subplan);

            // 使用累計後的數量輸出
            AppendLog(4, acc, totalCount.ToString(), displayName, Global.Custom_Sub_Gift, subplan, isPreview);
        }
        public void AppendLogFromStreamlabs_Resub(string acc, string displayName, string months, string subplan, bool isPreview = false)
        {
            string type = Global.Type_Resub;
            int monthsNum = 0;
            int.TryParse(months, out monthsNum);

            // 預覽時不寫入 DB
            if (!isPreview)
            {
                var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                DonateDB.Write(nowFull, type, acc, displayName, monthsNum, type, "", subplan);
            }

            // 檢查是否啟用輸出
            if (!Global.EnableResubOutput && !isPreview)
                return;

            // 輸出 (續訂不需要累計，直接顯示月數)
            AppendLog(6, acc, months, displayName, type, subplan, isPreview);
        }
        public void AppendLogFromStreamlabs_Sub(string acc, string displayName, string months, string subplan, bool isPreview = false)
        {
            string type = Global.Type_Sub;
            int monthsNum = 0;
            int.TryParse(months, out monthsNum);

            // 預覽時不寫入 DB
            if (!isPreview)
            {
                var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                DonateDB.Write(nowFull, type, acc, displayName, monthsNum, type, "", subplan);
            }

            // 檢查是否啟用輸出
            if (!Global.EnableSubOutput && !isPreview)
                return;

            // 輸出 (新訂閱不需要累計，直接顯示月數)
            AppendLog(7, acc, months, displayName, type, subplan, isPreview);
        }

        public void ClearSubGiftCount()
        {
            // 先匯出 CSV 備份
            string exportedFile = DonateDB.ExportToCsv();
            if (!string.IsNullOrEmpty(exportedFile))
            {
                AddLog($"已匯出備份檔案: {exportedFile}");
            }

            DonateDB.ClearAll();
            _obsDict.Clear();
            AddLog("已清除所有累計紀錄");

            // 清空 obs.txt
            string obsFileName = "obs.txt";
            File.WriteAllText(obsFileName, "");
        }
        public void AppendLogFromHiveBee(string name, string amount, string msg, bool isPreview = false)
        {
            string type = Global.Type_HiveBee;
            decimal donateAmount = 0;
            decimal.TryParse(amount, out donateAmount);

            // 預覽時不寫入 DB
            if (!isPreview)
            {
                var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                DonateDB.Write(nowFull, type, name, name, donateAmount, "TWD", msg, null);
            }

            // 從 DB 計算累計金額
            decimal totalAmount = isPreview ? donateAmount : DonateDB.GetTotalAmount(name, type);

            // 使用累計金額輸出
            AppendLog(5, name, totalAmount.ToString(), msg, "TWD", null, isPreview);
        }
        private void AppendLog(int nType, string name, string amount, string msg, string currency = "TWD", string subplan = null, bool isPreview = false, string acc = null)
        {
            if (string.IsNullOrEmpty(msg))
                msg = "";

            string type;
            string obsMsg;
            string displayName = null;
            if (nType == 7)
            {
                type = Global.Type_Sub;
                obsMsg = Global.Streamlabs_Sub_OBS_Msg;
                displayName = msg;
            }
            else if (nType == 6)
            {
                type = Global.Type_Resub;
                obsMsg = Global.Streamlabs_Resub_OBS_Msg;
                displayName = msg;
            }
            else if (nType == 5)
            {
                type = Global.Type_HiveBee;
                obsMsg = Global.HIVEBEE_OBS_Msg;
            }
            else if (nType == 4)
            {
                type = Global.Custom_Sub_Gift;
                obsMsg = Global.Streamlabs_SubGift_OBS_Msg;
                displayName = msg;
            }
            else if (nType == 3)
            {
                type = Global.Custom_Bits;
                obsMsg = Global.Streamlabs_Bits_OBS_Msg;
                if (!string.IsNullOrEmpty(acc))
                    msg += $"(帳號: {acc})";
            }
            else if (nType == 2)
            {
                type = Global.Type_Paypal;
                obsMsg = Global.Streamlabs_Paypal_OBS_Msg;
            }
            else if (nType == 1)
            {
                type = Global.Type_OPay;
                obsMsg = Global.OPAY_OBS_Msg;
            }
            else
            {
                type = Global.Type_ECPay;
                obsMsg = Global.ECPAY_OBS_Msg;
            }

            var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            var now = DateTime.Now.ToString("MM-dd HH:mm:ss");
            string formated_amount = Global.FormatAmount(amount);
            string logLine = $"[{now}][{type}] ";
            string extLogLine;
            if (nType == 7)
                extLogLine = string.Format("{0}({1}) 新訂閱{2}個月({3})", displayName, name, formated_amount, subplan);
            else if (nType == 6)
                extLogLine = string.Format("{0}({1}) 續訂{2}個月({3})", displayName, name, formated_amount, subplan);
            else if (nType == 4)
                extLogLine = string.Format("{0}({1}) 累計{2}{3}{4}", displayName, name, formated_amount, currency, (!string.IsNullOrEmpty(subplan) ? "(" + subplan + ")" : ""));
            else
                extLogLine = $"{name} 累計{formated_amount}{currency}, 說: {msg}";

            string obsOutput;
            if (nType == 7 || nType == 6)
                obsOutput = string.Format(obsMsg, displayName, formated_amount, subplan);
            else if (nType == 4)
                obsOutput = string.Format(obsMsg, displayName, formated_amount, currency, subplan);
            else
                obsOutput = string.Format(obsMsg, name, formated_amount, currency, subplan);

            if (isPreview)
            {
                MessageBox.Show(obsOutput);
                return;
            }

            _logQueue.Enqueue(logLine + extLogLine);
            // 使用 name+type+subplan 作為 key，相同帳號+類型+層級只保留最新一筆
            string obsKey = $"{name}|{type}|{subplan ?? ""}";
            _obsDict[obsKey] = obsOutput;
            // DB 寫入已在各個 AppendLogFrom* 方法中完成
        }
        public void AddLog(string sLog)
        {
            var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            _logQueue.Enqueue($"[{nowFull}] {sLog}");
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
        public void SetActiveStreamlabs(bool bActive)
        {
            SafeUpdateUI(() =>
            {
                lbStreamlabs_Status.Text = $"Streamlabs 狀態：{(bActive ? "有效" : "無效")}";
            });
        }
        public void SetActiveHivebee(bool bActive)
        {
            SafeUpdateUI(() =>
            {
                lbHiveBee_Status.Text = $"HiveBee 狀態：{(bActive ? "有效" : "無效")}";
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
            if (Global.IsEnableHiveBee())
            {
                _ = _servicesHiveBee.StartAsync(this, _ctsServices.Token);
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
            LoadCumulativeDataFromDB();
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

        private void BtClearDonateDB_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("確定要清除所有累計紀錄嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                ClearSubGiftCount();
            }
        }
    }
}
