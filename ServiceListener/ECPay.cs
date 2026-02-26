using Microsoft.AspNetCore.SignalR.Client;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace DonateMonitor.ServiceListener
{
    internal class ECPay : IServiceListener
    {
        private HubConnection _connection;
        public async Task StartAsync(Monitor monitor, CancellationToken token)
        {
            try
            {
                _connection = new HubConnectionBuilder()
            .WithUrl(
#if DEBUG
                "https://signalr-stage.ecpay.com.tw/donateHub",
#else
                "https://signalr.ecpay.com.tw/donateHub",
#endif
                options =>
                {
                    options.AccessTokenProvider = () =>
                        Task.FromResult(Global.ECPAY_LoginToken);
                }
            )
            .WithAutomaticReconnect()
            .Build();

                _connection.On<string, string, string>(
                    Global.ECPAY_ListenKey,
                    (name, amount, msg) =>
                    {
                        try
                        {
#if DEBUG
                            Console.WriteLine(
                                $"name={name}, amount={amount}, msg={msg}"
                            );
#endif
                            monitor.AppendLogFromECPay(name, amount, msg);
                        }
                        catch (Exception ex)
                        {
                            Global.WriteErrorLog($"[ECPay] event handle error: {ex}");
                        }
                    }
                );

                _connection.Closed += async (error) =>
                {
                    monitor.SetActiveECPay(false);
                    monitor.AddLog("與綠界伺服器連線中斷");
#if DEBUG
                    Console.WriteLine($"ECPay closed");
#endif
                };

                _connection.Reconnecting += (error) =>
                {
                    monitor.SetActiveECPay(false);
                    monitor.AddLog("嘗試重新與綠界伺服器連線中...");
#if DEBUG
                    Console.WriteLine($"ECPay Reconnecting");
#endif
                    return Task.CompletedTask;
                };

                _connection.Reconnected += (connectionId) =>
                {
                    monitor.SetActiveECPay(true);
                    monitor.AddLog("與綠界伺服器重新連線成功");
#if DEBUG
                    Console.WriteLine($"ECPay Reconnected");
#endif
                    return Task.CompletedTask;
                };

                await _connection.StartAsync(token);
#if DEBUG
                Console.WriteLine("ECPay connected");
#endif
                monitor.AddLog("與綠界伺服器連線成功");
                monitor.SetActiveECPay(true);

                // 🔥 關鍵：等到 token 被取消，才離開 StartAsync
                await Task.Delay(Timeout.Infinite, token);
            }
            catch (OperationCanceledException)
            {
                monitor.SetActiveECPay(false);
                await StopAsync();
            }
            catch (Exception ex)
            {
                monitor.AddLog($"啟動綠界服務失敗: {ex.Message}");
            }
        }

        public async Task StopAsync()
        {
            if (_connection != null)
            {
                await _connection.StopAsync();
                if (_connection != null)
                    await _connection.DisposeAsync();
                _connection = null;
            }
        }
    }
}
