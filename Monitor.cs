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
        private readonly Plugin.SoundAlerts _pluginSoundAlerts = new Plugin.SoundAlerts();
        private volatile bool _dataGridDirty = false;
        private bool _dataOperationInProgress = false;
        private volatile bool _dataSyncing = false;
        private readonly ConcurrentQueue<Action> _pendingEvents = new ConcurrentQueue<Action>();
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

                // 操作結束後，將排隊中的事件寫入 DB
                if (!_pendingEvents.IsEmpty && !_dataOperationInProgress)
                {
                    _dataSyncing = true;
                    try
                    {
                        while (_pendingEvents.TryDequeue(out var action))
                        {
                            try { action(); } catch (Exception ex) { Global.WriteErrorLog($"[Flush] {ex}"); }
                        }
                    }
                    finally
                    {
                        _dataSyncing = false;
                    }
                    _dataGridDirty = true;
                }

                // 有新資料時刷新資料管理頁（操作中／同步中不刷新，避免重入）
                if (_dataGridDirty && !_dataOperationInProgress && !_dataSyncing)
                {
                    _dataGridDirty = false;
                    try { LoadDonateData(); } catch { }
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
                else if (item.Type == Global.Type_SoundAlerts)
                {
                    obsOutput = string.Format(Global.SoundAlerts_OBS_Msg, item.Account, formated_amount, item.Currency, item.SubPlan);
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
            if (!isPreview && _dataOperationInProgress) { _pendingEvents.Enqueue(() => AppendLogFromECPay(name, amount, msg)); return; }
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
            if (!isPreview && _dataOperationInProgress) { _pendingEvents.Enqueue(() => AppendLogFromOPay(name, amount, msg)); return; }
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
            if (!isPreview && _dataOperationInProgress) { _pendingEvents.Enqueue(() => AppendLogFromStreamlabs_Paypal(name, amount, currency, msg)); return; }
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
            if (!isPreview && _dataOperationInProgress) { _pendingEvents.Enqueue(() => AppendLogFromStreamlabs_Bits(acc, name, amount, msg)); return; }
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
            if (!isPreview && _dataOperationInProgress) { _pendingEvents.Enqueue(() => AppendLogFromStreamlabs_SubGift(acc, amount, displayName, subplan)); return; }
            // 累加 subgift 並獲取累計後的數量
            int giftAmount = 1;
            int.TryParse(amount, out giftAmount);
            if (giftAmount <= 0) giftAmount = 1;

            int totalCount;
            if (!isPreview)
            {
                var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                if (subplan == Global.Custom_Sub_Tier1)
                {
                    // 層一：累計合併為單筆記錄
                    totalCount = DonateDB.AccumulateSubGift(nowFull, acc, displayName, giftAmount, subplan);
                }
                else
                {
                    // 層二、層三：逐筆寫入
                    DonateDB.Write(nowFull, Global.Custom_Sub_Gift, acc, displayName, giftAmount, Global.Custom_Sub_Gift, "", subplan);
                    totalCount = giftAmount;
                }
            }
            else
            {
                totalCount = giftAmount;
            }

            // 使用累計後的數量輸出
            AppendLog(4, acc, totalCount.ToString(), displayName, Global.Custom_Sub_Gift, subplan, isPreview);
        }
        public void AppendLogFromStreamlabs_Resub(string acc, string displayName, string months, string subplan, bool isPreview = false)
        {
            if (!isPreview && _dataOperationInProgress) { _pendingEvents.Enqueue(() => AppendLogFromStreamlabs_Resub(acc, displayName, months, subplan)); return; }
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
            if (!isPreview && _dataOperationInProgress) { _pendingEvents.Enqueue(() => AppendLogFromStreamlabs_Sub(acc, displayName, months, subplan)); return; }
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
            if (!isPreview && _dataOperationInProgress) { _pendingEvents.Enqueue(() => AppendLogFromHiveBee(name, amount, msg)); return; }
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
        public void AppendLogFromSoundAlerts(string name, string amount, string costType, bool isPreview = false)
        {
            if (!isPreview && _dataOperationInProgress) { _pendingEvents.Enqueue(() => AppendLogFromSoundAlerts(name, amount, costType)); return; }
            string type = Global.Type_SoundAlerts;
            decimal donateAmount = 0;
            decimal.TryParse(amount, out donateAmount);

            if (!isPreview)
            {
                var nowFull = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                DonateDB.Write(nowFull, type, name, name, donateAmount, costType, "", null);
            }

            decimal totalAmount = isPreview ? donateAmount : DonateDB.GetTotalAmount(name, type);

            AppendLog(8, name, totalAmount.ToString(), "", costType, null, isPreview);
        }
        private void AppendLog(int nType, string name, string amount, string msg, string currency = "TWD", string subplan = null, bool isPreview = false, string acc = null)
        {
            if (string.IsNullOrEmpty(msg))
                msg = "";

            string type;
            string obsMsg;
            string displayName = null;
            if (nType == 8)
            {
                type = Global.Type_SoundAlerts;
                obsMsg = Global.SoundAlerts_OBS_Msg;
            }
            else if (nType == 7)
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
            // 標記資料管理頁需要刷新
            _dataGridDirty = true;
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
        public void SetActiveSoundAlerts(bool bActive)
        {
            SafeUpdateUI(() =>
            {
                lbSoundAlerts_Status.Text = $"SoundAlerts 狀態：{(bActive ? "有效" : "無效")}";
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
            if (Global.IsEnableSoundAlerts())
            {
                _ = _pluginSoundAlerts.StartAsync(this, _ctsServices.Token);
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
            LoadDonateData();
            AddLog($"已讀取 {_donateRecords.Count} 筆累計資料");
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
            if (_dataSyncing) { MessageBox.Show("資料同步中，請稍後", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            _dataOperationInProgress = true;
            try
            {
                var result = MessageBox.Show("確定要清除所有累計紀錄嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    ClearSubGiftCount();
                    LoadDonateData();
                }
            }
            catch (Exception ex)
            {
                Global.WriteErrorLog($"[資料管理] 清除失敗: {ex}");
                MessageBox.Show($"清除資料失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _dataOperationInProgress = false;
            }
        }

        private List<DonateRecord> _donateRecords = new List<DonateRecord>();

        private void LoadDonateData()
        {
            _donateRecords = DonateDB.GetAllRecords();
            dgvDonateData.DataSource = null;
            dgvDonateData.DataSource = _donateRecords;

            // 設定欄位標題
            if (dgvDonateData.Columns.Count > 0)
            {
                dgvDonateData.Columns["Id"].HeaderText = "ID";
                dgvDonateData.Columns["Id"].ReadOnly = true;
                dgvDonateData.Columns["DateTime"].HeaderText = "時間";
                dgvDonateData.Columns["Account"].HeaderText = "帳號";
                dgvDonateData.Columns["DisplayName"].HeaderText = "顯示名稱";
                dgvDonateData.Columns["Amount"].HeaderText = "金額";
                dgvDonateData.Columns["Currency"].HeaderText = "幣別";
                dgvDonateData.Columns["Message"].HeaderText = "訊息";
                dgvDonateData.Columns["SubPlan"].HeaderText = "方案";

                // 將 Type 欄位替換為下拉選單
                int typeColumnIndex = dgvDonateData.Columns["Type"].Index;
                dgvDonateData.Columns.Remove("Type");

                var typeComboColumn = new DataGridViewComboBoxColumn
                {
                    Name = "Type",
                    HeaderText = "類型",
                    DataPropertyName = "Type",
                    DisplayIndex = typeColumnIndex,
                    FlatStyle = FlatStyle.Flat
                };

                // 添加支援的類型選項
                typeComboColumn.Items.AddRange(new string[]
                {
                    Global.Type_ECPay,
                    Global.Type_OPay,
                    Global.Type_HiveBee,
                    Global.Type_Paypal,
                    Global.Type_Sub,
                    Global.Type_Resub,
                    Global.Custom_Sub_Gift,
                    Global.Custom_Bits,
                    Global.Type_SoundAlerts
                });

                dgvDonateData.Columns.Insert(typeColumnIndex, typeComboColumn);
            }
        }

        private void BtRefreshData_Click(object sender, EventArgs e)
        {
            try
            {
                LoadDonateData();
                AddLog("已重新載入資料");
            }
            catch (Exception ex)
            {
                Global.WriteErrorLog($"[資料管理] 載入失敗: {ex}");
                MessageBox.Show($"載入資料失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtAddData_Click(object sender, EventArgs e)
        {
            if (_dataSyncing) { MessageBox.Show("資料同步中，請稍後", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            _dataOperationInProgress = true;
            try
            {
                using (var form = new AddDonateRecord())
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        int newId = DonateDB.Insert(
                            form.RecordDateTime,
                            form.RecordType,
                            form.RecordAccount,
                            form.RecordDisplayName,
                            form.RecordAmount,
                            form.RecordCurrency,
                            form.RecordMessage,
                            form.RecordSubPlan
                        );
                        LoadDonateData();
                        ReloadObsData();
                        AddLog($"已新增資料 (ID: {newId})");
                    }
                }
            }
            catch (Exception ex)
            {
                Global.WriteErrorLog($"[資料管理] 新增失敗: {ex}");
                MessageBox.Show($"新增資料失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _dataOperationInProgress = false;
            }
        }

        private void BtDeleteSelected_Click(object sender, EventArgs e)
        {
            if (dgvDonateData.SelectedRows.Count == 0)
            {
                MessageBox.Show("請先選取要刪除的資料列", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (_dataSyncing) { MessageBox.Show("資料同步中，請稍後", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            _dataOperationInProgress = true;
            try
            {
                var result = MessageBox.Show($"確定要刪除選取的 {dgvDonateData.SelectedRows.Count} 筆資料嗎？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result != DialogResult.Yes)
                    return;

                int deleteCount = 0;
                foreach (DataGridViewRow row in dgvDonateData.SelectedRows)
                {
                    if (row.DataBoundItem is DonateRecord record)
                    {
                        DonateDB.DeleteById(record.Id);
                        deleteCount++;
                    }
                }

                LoadDonateData();
                ReloadObsData();
                AddLog($"已刪除 {deleteCount} 筆資料");
            }
            catch (Exception ex)
            {
                Global.WriteErrorLog($"[資料管理] 刪除失敗: {ex}");
                MessageBox.Show($"刪除資料失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _dataOperationInProgress = false;
            }
        }

        /// <summary>
        /// ComboBox（類型下拉）選擇後立即 commit，觸發 CellValueChanged
        /// </summary>
        private void DgvDonateData_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvDonateData.IsCurrentCellDirty)
                dgvDonateData.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        /// <summary>
        /// 儲存格值變更時自動儲存該筆記錄
        /// </summary>
        private void DgvDonateData_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (_dataSyncing) return;

            // LoadDonateData 會重設 DataSource 觸發此事件，用旗標擋掉
            if (_dataOperationInProgress) return;

            _dataOperationInProgress = true;
            try
            {
                if (e.RowIndex >= _donateRecords.Count) return;
                var record = _donateRecords[e.RowIndex];

                DonateDB.UpdateById(
                    record.Id,
                    record.DateTime,
                    record.Type,
                    record.Account,
                    record.DisplayName,
                    record.Amount,
                    record.Currency,
                    record.Message,
                    record.SubPlan
                );

                ReloadObsData();
                AddLog($"已自動儲存 ID:{record.Id}（{record.Account}）");
            }
            catch (Exception ex)
            {
                Global.WriteErrorLog($"[資料管理] 自動儲存失敗: {ex}");
            }
            finally
            {
                _dataOperationInProgress = false;
            }
        }

        private void DgvDonateData_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != System.Windows.Forms.MouseButtons.Right)
                return;
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            var col = dgvDonateData.Columns[e.ColumnIndex];
            if (col.DataPropertyName != "Amount" && col.Name != "Amount")
                return;

            // 選取該儲存格
            dgvDonateData.CurrentCell = dgvDonateData[e.ColumnIndex, e.RowIndex];

            // 取得螢幕座標並顯示右鍵選單
            var cellRect = dgvDonateData.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
            var screenPoint = dgvDonateData.PointToScreen(new Point(cellRect.Left + e.X, cellRect.Top + e.Y));
            cmsAmount.Tag = e.RowIndex;
            cmsAmount.Show(screenPoint);
        }

        private void TsmiAddAmount_Click(object sender, EventArgs e)
        {
            int rowIndex = (int)cmsAmount.Tag;
            if (rowIndex < 0 || rowIndex >= _donateRecords.Count)
                return;

            if (_dataSyncing) { MessageBox.Show("資料同步中，請稍後", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            _dataOperationInProgress = true;
            try
            {
                var record = _donateRecords[rowIndex];
                string input = ShowInputDialog("增加金額", $"目前金額: {record.Amount}\n請輸入要增加的金額：");
                if (input == null)
                    return;

                if (!decimal.TryParse(input, out decimal addAmount))
                {
                    MessageBox.Show("請輸入有效的數字", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 使用 DB 原子操作 Amount = Amount + delta，避免 read-modify-write 競爭
                decimal newAmount = DonateDB.AddAmountById(record.Id, addAmount);
                LoadDonateData();
                ReloadObsData();
                AddLog($"已增加金額 {addAmount} 至 ID:{record.Id}（{record.Account}），新金額: {newAmount}");
            }
            catch (Exception ex)
            {
                Global.WriteErrorLog($"[資料管理] 增加金額失敗: {ex}");
                MessageBox.Show($"增加金額失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _dataOperationInProgress = false;
            }
        }

        private string ShowInputDialog(string title, string prompt)
        {
            using (var form = new Form())
            {
                form.Text = title;
                form.StartPosition = FormStartPosition.CenterParent;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.ClientSize = new Size(350, 130);
                form.Font = new Font("Noto Sans TC", 10F);

                var label = new Label { Text = prompt, Left = 12, Top = 12, AutoSize = true };
                var textBox = new TextBox { Left = 12, Top = 60, Width = 320 };
                var btnOk = new Button { Text = "確定", DialogResult = DialogResult.OK, Left = 152, Top = 92, Width = 80 };
                var btnCancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Left = 240, Top = 92, Width = 80 };

                form.Controls.AddRange(new Control[] { label, textBox, btnOk, btnCancel });
                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;

                return form.ShowDialog() == DialogResult.OK ? textBox.Text.Trim() : null;
            }
        }

        private void ReloadObsData()
        {
            _obsDict.Clear();
            LoadCumulativeDataFromDB();
            WriteObsFile();
        }

        private void WriteObsFile()
        {
            string obsFileName = "obs.txt";
            if (_obsDict.Count > 0)
            {
                var lines = _obsDict.Values.ToArray();
                if (Global.OBS_OutputMode == 1)
                    File.WriteAllText(obsFileName, string.Join(" ", lines), Encoding.UTF8);
                else
                    File.WriteAllText(obsFileName, string.Join(Environment.NewLine, lines), Encoding.UTF8);
            }
            else
            {
                File.WriteAllText(obsFileName, "", Encoding.UTF8);
            }
        }
    }
}
