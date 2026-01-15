using Microsoft.AspNetCore.SignalR.Client;
using System;
using System.Collections.Generic;
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
#if DEBUG
                        Console.WriteLine(
                            $"name={name}, amount={amount}, msg={msg}"
                        );
#endif
                        monitor.AppendLogFromECPay(name, amount, msg);
                    }
                );

                _connection.Closed += async (error) =>
                {
                    monitor.SetActiveECPay(false);
#if DEBUG
                    Console.WriteLine($"ECPay closed");
#endif
                };

                _connection.Reconnecting += (error) =>
                {
                    monitor.SetActiveECPay(false);
#if DEBUG
                    Console.WriteLine($"ECPay Reconnecting");
#endif
                    return Task.CompletedTask;
                };

                _connection.Reconnected += (connectionId) =>
                {
                    monitor.SetActiveECPay(true);
#if DEBUG
                    Console.WriteLine($"ECPay Reconnected");
#endif
                    return Task.CompletedTask;
                };


                await _connection.StartAsync(token);
#if DEBUG
                Console.WriteLine("ECPay connected");
#endif
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
                Global.ShowError($"啟動綠界服務失敗: {ex.Message}", true);
            }
        }

        public async Task StopAsync()
        {
            if (_connection != null)
            {
                await _connection.StopAsync();
                await _connection.DisposeAsync();
                _connection = null;
            }
        }
    }
}
