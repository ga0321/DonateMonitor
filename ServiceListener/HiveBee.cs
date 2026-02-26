using Microsoft.AspNet.SignalR.Client;
using Microsoft.AspNet.SignalR.Client.Transports;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace DonateMonitor.ServiceListener
{
    internal class HiveBee : IServiceListener
    {
        private readonly string _hubName = "notificationHub";
        private readonly string _connectionData = Uri.EscapeDataString(@"[{""name"":""notificationhub""}]");  // URL-encoded connectionData

        private readonly HttpClient _http = new HttpClient
        {
            BaseAddress = new Uri("https://www.hivebee.com.tw")
        };
        private HubConnection _conn;
        private IHubProxy _hub;

        private readonly SemaphoreSlim _startStopLock = new SemaphoreSlim(1, 1);
        private CancellationToken _ct;

        public async Task StartAsync(Monitor monitor, CancellationToken ct)
        {
            _ct = ct;

            // 無限維持：斷線就重連；cancel 才退出
            int attempt = 0;
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    attempt++;
                    await StartOnceAsync(monitor, ct).ConfigureAwait(false);

                    // StartOnceAsync 連上後會等到 Closed 才 return（或 cancel）
                    attempt = 0; // reset backoff after a successful full session
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested)
                {
                    break;
                }
                catch (Exception ex)
                {
                    monitor.AddLog($"HiveBee服務啟動失敗: {ex.Message}");
                    var delay = ComputeBackoff(attempt);
                    await Task.Delay(delay, ct).ConfigureAwait(false);
                }
            }

            await StopAsync().ConfigureAwait(false);
        }

        public async Task StopAsync()
        {
            await _startStopLock.WaitAsync().ConfigureAwait(false);
            try
            {
                if (_conn != null)
                {
                    try { _conn.Stop(); } catch { /* ignore */ }
                    try { _conn.Dispose(); } catch { /* ignore */ }
                    _conn = null;
                    _hub = null;
                }
            }
            finally
            {
                _startStopLock.Release();
            }
        }

        public void Dispose()
        {
            try { StopAsync().GetAwaiter().GetResult(); } catch { }
            _http.Dispose();
            _startStopLock.Dispose();
        }

        // -------------------------
        // Core session (one connect)
        // -------------------------
        private async Task StartOnceAsync(Monitor monitor, CancellationToken ct)
        {
            await _startStopLock.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                // 確保上一條乾淨關閉
                if (_conn != null)
                {
                    try { _conn.Stop(); } catch { }
                    try { _conn.Dispose(); } catch { }
                    _conn = null;
                    _hub = null;
                }

                // 1) App negotiate -> RedirectUrl + AccessToken
                var appNego = await AppNegotiateAsync(ct).ConfigureAwait(false);
                var redirectUrl = appNego.Value<string>("RedirectUrl");
                var accessToken = appNego.Value<string>("AccessToken");

                if (string.IsNullOrEmpty(redirectUrl) || string.IsNullOrEmpty(accessToken))
                    throw new InvalidOperationException("HiveBee App negotiate returned empty RedirectUrl/AccessToken.");

                // 2) Service base + redirect query (must NOT put query inside HubConnection url)
                var serviceBase = StripQuery(redirectUrl);        // https://.../aspnetclient
                var redirectQs = GetQuery(redirectUrl);           // _=...&asrs_request_id=...&asrs.op=%2Fsignalr

                // 3) Service negotiate (optional but nice for verification; matches F12)
                //    Use Authorization: Bearer <AccessToken>
                await ServiceNegotiateAsync(serviceBase, redirectQs, accessToken, ct).ConfigureAwait(false);

                // 4) Build HubConnection to service (/aspnetclient) with useDefaultUrl:false
                //    - QueryString: keep redirect query + access_token (for connect/ws like F12)
                //    - Header: Authorization: Bearer (for service negotiate like F12)
                var qs = BuildServiceQuery(redirectQs, accessToken);

                _conn = new HubConnection(serviceBase, qs, useDefaultUrl: false);
                _conn.Headers["Authorization"] = "Bearer " + accessToken;

                _hub = _conn.CreateHubProxy(_hubName);

                _hub.On<string, int, object>("NotifyDonateObj", (type, value, data) =>
                {
                    try
                    {
                        var jo = data as JObject ?? JObject.FromObject(data);
#if DEBUG
                        Console.WriteLine($"HiveBee type={type} value={value} data={data}");
#endif
                        try { Global.WriteDebugLog($"[HiveBee] {data?.ToString()}"); } catch { }

                        // 1 = text, 6 = video
                        var alertType = jo.Value<string>("Type"); // 測試環境是AlertType WTF
                        if (type.Equals("tool") && (alertType == "1" || alertType == "6"))
                        {
                            string name = jo.Value<string>("Name") ?? "";
                            string amount = jo["Amount"]?.ToString() ?? "0";
                            string text = jo.Value<string>("Text") ?? "";
                            monitor.AppendLogFromHiveBee(name, amount, text);
                        }
                    }
                    catch (Exception ex)
                    {
                        Global.WriteErrorLog($"[HiveBee] event handle error: {ex}");
                    }
                });

                // Events
                _conn.Closed += async () =>
                {
                    monitor.SetActiveECPay(false);
                    monitor.AddLog("與HiveBee伺服器連線中斷");
#if DEBUG
                    Console.WriteLine($"HiveBee closed");
#endif
                };

                _conn.Reconnecting += () =>
                {
                    monitor.SetActiveECPay(false);
                    monitor.AddLog("嘗試重新與HiveBee伺服器連線中...");
#if DEBUG
                    Console.WriteLine($"HiveBee Reconnecting");
#endif
                };

                _conn.Reconnected += () =>
                {
                    monitor.SetActiveECPay(true);
                    monitor.AddLog("與HiveBee伺服器重新連線成功");
#if DEBUG
                    Console.WriteLine($"HiveBee Reconnected");
#endif
                };

                _conn.Error += (e) =>
                {
                    monitor.SetActiveECPay(false);
                    monitor.AddLog("與HiveBee伺服器連線錯誤");
#if DEBUG
                    Console.WriteLine($"HiveBee Error: {e.Message}");
#endif
                };

                // 5) Start websocket (align to F12 transport=webSockets)
                await _conn.Start(new WebSocketTransport()).ConfigureAwait(false);

                // 6) joinRoom(widgetToken)
                var joined = await _hub.Invoke<bool>("joinRoom", Global.HiveBeeKey).ConfigureAwait(false);
                if (!joined)
                    throw new InvalidOperationException("HiveBee joinRoom returned false (token invalid or permission denied).");

                monitor.SetActiveHivebee(true);
                monitor.AddLog("與HiveBee伺服器連線成功");

                // 7) Wait until closed or cancelled
                var closedTcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                Action closedHandler = () => closedTcs.TrySetResult(true);
                _conn.Closed += closedHandler;

                try
                {
                    using (ct.Register(() => closedTcs.TrySetCanceled()))
                    {
                        await closedTcs.Task.ConfigureAwait(false);
                    }
                }
                finally
                {
                    _conn.Closed -= closedHandler;
                }
            }
            finally
            {
                _startStopLock.Release();
            }
        }

        // -------------------------
        // Negotiation
        // -------------------------
        private async Task<JObject> AppNegotiateAsync(CancellationToken ct)
        {
            // /signalr/negotiate?clientProtocol=2.1&connectionData=[{"name":"notificationhub"}]&_=...
            var url = $"/signalr/negotiate?clientProtocol=2.1&connectionData={_connectionData}&_={NowMs()}";
            var txt = await _http.GetStringAsync(url).ConfigureAwait(false);
            ct.ThrowIfCancellationRequested();
            return JObject.Parse(txt);
        }

        private static async Task<JObject> ServiceNegotiateAsync(
            string serviceBase,
            string redirectQs,
            string accessToken,
            CancellationToken ct)
        {
            // https://.../aspnetclient/negotiate?clientProtocol=2.1&{redirectQs}&connectionData=...&_=...
            var url =
                $"{serviceBase}/negotiate?clientProtocol=2.1&{redirectQs}&connectionData=%5B%7B%22name%22%3A%22notificationhub%22%7D%5D&_={NowMs()}";

            using (var http = new HttpClient())
            {
                http.DefaultRequestHeaders.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);

                var txt = await http.GetStringAsync(url).ConfigureAwait(false);
                ct.ThrowIfCancellationRequested();
                return JObject.Parse(txt);
            }
        }

        private static string BuildServiceQuery(string redirectQs, string accessToken)
        {
            // Keep Azure redirect params + access_token (for connect/websocket)
            // Note: do NOT put query inside HubConnection url; pass via ctor queryString.
            var qs = (redirectQs ?? "").Trim('&');

            if (!string.IsNullOrEmpty(qs))
                qs += "&";

            qs += "access_token=" + Uri.EscapeDataString(accessToken);

            return qs;
        }

        // -------------------------
        // Helpers
        // -------------------------
        private static long NowMs() => DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();

        private static string StripQuery(string url)
        {
            var i = url.IndexOf('?');
            return i >= 0 ? url.Substring(0, i) : url;
        }

        private static string GetQuery(string url)
        {
            var i = url.IndexOf('?');
            return i >= 0 ? url.Substring(i + 1) : "";
        }

        private static TimeSpan ComputeBackoff(int attempt)
        {
            // 1s,2s,4s,8s,16s,30s cap + jitter
            var pow = Math.Min(attempt, 5);
            var seconds = Math.Min(30, (int)Math.Pow(2, pow));
            var jitter = new Random().Next(0, 500);
            return TimeSpan.FromSeconds(seconds) + TimeSpan.FromMilliseconds(jitter);
        }
    }
}
