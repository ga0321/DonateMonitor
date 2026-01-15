using SocketIOClient;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace DonateMonitor.ServiceListener
{
    internal class OPay : IServiceListener
    {
        private SocketIOClient.SocketIO _connection = null;
        public async Task StartAsync(Monitor monitor, CancellationToken token)
        {
            try
            {
                _connection = new SocketIOClient.SocketIO(
//#if DEBUG
//                    $"https://socket-stage.opay.tw/web/live/{Global.OPAY_ListenKey}",
//#else
                    $"https://socket.opay.tw/web/live/{Global.OPAY_ListenKey}",
//#endif
                    new SocketIOOptions
                    {
                        Auth = new {},
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
                    monitor.SetActiveOPay(true);
#if DEBUG
                    Console.WriteLine("OPay OnConnected");
#endif
                };

                // === 斷線 ===
                _connection.OnDisconnected += (sender, reason) =>
                {
                    monitor.SetActiveOPay(false);
#if DEBUG
                    Console.WriteLine($"OPay OnDisconnected: {reason}");
#endif
                };

                // === 重新連線成功 ===
                _connection.OnReconnected += (sender, e) =>
                {
                    monitor.SetActiveOPay(true);
#if DEBUG
                    Console.WriteLine("OPay OnReconnected");
#endif
                };

                // === 連線錯誤 / 發生錯誤 ===
                _connection.OnError += (sender, error) =>
                {
                    Global.WriteErrorLog(error);
                    monitor.SetActiveOPay(false);
#if DEBUG
                    Console.WriteLine($"OPay OnError: {error}");
#endif
                };

                // === 收到 Server 通知 ===
                _connection.On("notify", response =>
                {
                    var resp = response.ToString();
#if DEBUG
                    Console.WriteLine(resp);
#endif
                    var arr = JArray.Parse(resp);

                    if (!(arr[0]?["data"]?["lstDonate"] is JArray lstDonate))
                        return;

                    foreach (var d in lstDonate)
                    {
                        //string donateid = d["donateid"]?.ToString();
                        string name = d["name"]?.ToString();
                        string amount = d["amount"]?.ToString();
                        string msg = d["msg"]?.ToString();

                        monitor.AppendLogFromOPay(name, amount, msg);
                    }
                });

                await _connection.ConnectAsync();

                // 🔥 關鍵：等到 token 被取消，才離開 StartAsync
                await Task.Delay(Timeout.Infinite, token);
            }
            catch (OperationCanceledException)
            {
                monitor.SetActiveOPay(false);
                await StopAsync();
            }
            catch (Exception ex)
            {
                Global.ShowError($"啟動歐富寶服務失敗: {ex.Message}", true);
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
