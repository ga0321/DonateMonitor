using Newtonsoft.Json.Linq;
using SocketIOClient;
using SocketIOClient.Transport;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace DonateMonitor.ServiceListener
{
    internal class Streamlabs : IServiceListener
    {
        private SocketIOClient.SocketIO _connection = null;
        public static string SubPlanToText(string plan)
        {
            if (plan == "3000")
                return Global.Custom_Sub_Tier3;
            else if (plan == "2000")
                return Global.Custom_Sub_Tier2;
            else
                return Global.Custom_Sub_Tier1;
        }
        public async Task StartAsync(Monitor monitor, CancellationToken token)
        {
            try
            {
                _connection = new SocketIOClient.SocketIO(
                    "https://sockets.streamlabs.com",
                    new SocketIOOptions
                    {
                        // Streamlabs 是 https://sockets.streamlabs.com?token=xxx
                        Query = new Dictionary<string, string>
                        {
                            ["token"] = Global.StreamlabsKey
                        },

                        // 常見相容性設定：強制走 WebSocket + Engine.IO v3（對應 socket.io 2.x）
                        Transport = TransportProtocol.WebSocket,
                        EIO = (SocketIO.Core.EngineIO)3,

                        Reconnection = true,
                        ReconnectionAttempts = int.MaxValue,   // 🔥 無限次
                        ReconnectionDelay = 1000,               // 第一次重連 1 秒
                        ReconnectionDelayMax = 10000,            // 最長 10 秒
                        RandomizationFactor = 0.5                // 抖動（避免打爆 server）
                    }
                );

                // === 連線成功 ===
                _connection.OnConnected += (sender, e) =>
                {
                    monitor.SetStreamlabs(true);
#if DEBUG
                    Console.WriteLine("Streamlabs OnConnected");
#endif
                };

                // === 斷線 ===
                _connection.OnDisconnected += (sender, reason) =>
                {
                    monitor.SetStreamlabs(false);
#if DEBUG
                    Console.WriteLine($"Streamlabs OnDisconnected: {reason}");
#endif
                };

                // === 重新連線成功 ===
                _connection.OnReconnected += (sender, e) =>
                {
                    monitor.SetStreamlabs(true);
#if DEBUG
                    Console.WriteLine("Streamlabs OnReconnected");
#endif
                };

                // === 連線錯誤 / 發生錯誤 ===
                _connection.OnError += (sender, error) =>
                {
                    monitor.SetStreamlabs(false);
                    Global.WriteErrorLog(error);
#if DEBUG
                    Console.WriteLine($"Streamlabs OnError: {error}");
#endif
                };

                _connection.On("event", resp =>
                {
                    string raw;
                    try { raw = resp.GetValue<string>(); }
                    catch { raw = resp.ToString(); }

#if DEBUG
                    Console.WriteLine(raw);
#endif

                    // ✅ 最外層是陣列：[ {type, message, ...} ]
                    var arr = JArray.Parse(raw);

                    foreach (var ev in arr)
                    {
                        var type = ev["type"]?.ToString();
#if DEBUG
                        Console.WriteLine($"type={type}");
#endif

                        if (type == "donation")
                        {
                            if (ev["message"] is JArray msgs)
                            {
                                foreach (var m in msgs)
                                {
                                    var from = m["from"]?.ToString();
                                    var amount = m["amount"]?.ToString();
                                    //var formatted = m["formattedAmount"]?.ToString();
                                    var msg = m["message"]?.ToString();
                                    var currency = m["currency"]?.ToString();

                                    monitor.AppendLogFromStreamlabs_Paypal(from, amount, currency, msg);
                                }
                            }
                        }
                        else if (type == "bits")
                        {
                            if (ev["message"] is JArray msgs)
                            {
                                foreach (var m in msgs)
                                {
                                    var from = m["name"]?.ToString();
                                    var amount = m["amount"]?.ToString();
                                    var msg = m["message"]?.ToString();
                                    monitor.AppendLogFromStreamlabs_Bits(from, amount, msg);
                                }
                            }
                        }
                        else if (type == "subscription" || type == "subMysteryGift")
                        {
                            if (ev["message"] is JArray msgs)
                            {
                                foreach (var m in msgs)
                                {
                                    var subType = m["sub_type"]?.ToString();
                                    if (string.IsNullOrEmpty(subType))
                                        continue;

                                    var condition = m["condition"]?.ToString();
                                    if (string.IsNullOrEmpty(condition))
                                        continue;

                                    bool isAnon = (condition == "ANON_SUBSCRIPTION_GIFT" || condition == "MIN_ANON_SUBMYSTERYGIFT");
                                    var gifter_ac = isAnon ? Global.Custom_ANON : m["gifter"]?.ToString();
                                    var gifter_display = isAnon ? Global.Custom_ANON : m["gifter_display_name"]?.ToString();
                                    var sub_plan = m["sub_plan"]?.ToString();

                                    if (subType == "subgift")
                                    {
                                        
                                        var amount = m["months"]?.ToString();
#if DEBUG
                                        gifter_ac += "(gifter_ac)";
                                        gifter_display += "(gifter_display)";
#endif

                                        monitor.AppendLogFromStreamlabs_SubGift(gifter_ac, amount, gifter_display, SubPlanToText(sub_plan));
                                    }
                                    else if (subType == "submysterygift")
                                    {
                                        var amount = m["amount"]?.ToString();
#if DEBUG
                                        gifter_ac += "(gifter_ac)";
                                        gifter_display += "(gifter_display)";
#endif
                                        monitor.AppendLogFromStreamlabs_SubGift(gifter_ac, amount, gifter_display, SubPlanToText(sub_plan));
                                    }
                                }
                            }
                        }
                    }
                });

                await _connection.ConnectAsync();

                // 🔥 關鍵：等到 token 被取消，才離開 StartAsync
                await Task.Delay(Timeout.Infinite, token);
            }
            catch (OperationCanceledException)
            {
                await StopAsync();
            }
            catch (Exception ex)
            {
                Global.ShowError($"啟動Streamlabs服務失敗: {ex.Message}", true);
            }
        }

        public async Task StopAsync()
        {
            if (_connection != null)
            {
                await _connection.DisconnectAsync();
                _connection = null;
            }
        }
    }
}
