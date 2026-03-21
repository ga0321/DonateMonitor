using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json.Linq;
using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace DonateMonitor.Plugin
{
    internal class SoundAlerts : IPlugin
    {
        private WebView2 _webView = null;
        private Monitor _monitor = null;

        // 去重：記錄最後一筆事件
        private string _lastEventKey = null;
        private DateTime _lastEventTime = DateTime.MinValue;
        private readonly TimeSpan _dedupeWindow = TimeSpan.FromSeconds(3);

        private string _userDataFolder =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WebView2UserData");

        public async Task StartAsync(Monitor monitor, CancellationToken token)
        {
            _monitor = monitor;

            try
            {
                // 在 UI thread 建立 WebView2（StartAsync 由 UI thread 呼叫）
                _webView = new WebView2();
                _webView.Size = new Size(1, 1);
                _webView.Visible = false;
                monitor.Controls.Add(_webView);

                monitor.AddLog("正在初始化SoundAlerts...");

                var env = await CoreWebView2Environment.CreateAsync(null, _userDataFolder);
                await _webView.EnsureCoreWebView2Async(env);

                _webView.CoreWebView2.Settings.IsWebMessageEnabled = true;
                _webView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
                _webView.CoreWebView2.DOMContentLoaded += async (s, e) =>
                {
                    await InjectBridgeAsync();
                };

                monitor.AddLog("SoundAlerts WebView2 已就緒");
                monitor.SetActiveSoundAlerts(true);

                // 導航到 Overlay URL
                _webView.CoreWebView2.Navigate(Global.SoundAlertsOverlayUrl);

                // 等到 token 被取消
                await Task.Delay(Timeout.Infinite, token);
            }
            catch (OperationCanceledException)
            {
                await StopAsync();
            }
            catch (Exception ex)
            {
                monitor.AddLog($"啟動SoundAlerts失敗: {ex.Message}");
                Global.WriteErrorLog($"[SoundAlerts] Start failed: {ex}");
                monitor.SetActiveSoundAlerts(false);
            }
        }

        public async Task StopAsync()
        {
            if (_webView != null)
            {
                try
                {
                    _webView.CoreWebView2?.Stop();
                    _webView.Dispose();
                }
                catch { }
                _webView = null;
            }
            _monitor?.SetActiveSoundAlerts(false);
            await Task.CompletedTask;
        }

        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string raw = e.WebMessageAsJson;

                var obj = JObject.Parse(raw);
                var kind = (string)obj["kind"];

#if DEBUG
                if (kind != "console-log" && kind != "bridge-log")
                    Console.WriteLine($"[SoundAlerts] {kind}: {raw.Substring(0, Math.Min(raw.Length, 200))}");
#endif

                if (kind == "bridge-log" || kind == "bridge-error")
                {
                    try { Global.WriteDebugLog($"[SoundAlerts][{kind}] {(string)obj["message"]}"); } catch { }
                    return;
                }

                // WebSocket 訊息
                if (kind == "ws-message")
                {
                    var json = obj["json"];
                    if (json != null && json.Type == JTokenType.Object)
                    {
                        var parsed = ParseSoundAlert((JObject)json);
                        TryEmitEvent(parsed);
                    }
                    return;
                }

                // Fetch 回應 (Firestore)
                if (kind == "fetch-response")
                {
                    string url = (string)obj["url"];
                    if (!string.IsNullOrEmpty(url) &&
                        url.IndexOf("firestore", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        var json = obj["json"];
                        if (json != null && json.Type == JTokenType.Object)
                        {
                            var parsed = ParseSoundAlert((JObject)json);
                            TryEmitEvent(parsed);
                        }
                        else
                        {
                            string rawText = (string)obj["raw"];
                            if (!string.IsNullOrEmpty(rawText))
                            {
                                var parsed = ParseSoundAlertFromRawByRegex(rawText);
                                TryEmitEvent(parsed);
                            }
                        }
                    }
                    return;
                }

                // XHR 回應
                if (kind == "xhr-response")
                {
                    string url = (string)obj["url"];
                    if (!string.IsNullOrEmpty(url) &&
                        (url.IndexOf("alert", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         url.IndexOf("sound", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         url.IndexOf("event", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        var json = obj["json"];
                        if (json != null && json.Type == JTokenType.Object)
                        {
                            var parsed = ParseSoundAlert((JObject)json);
                            TryEmitEvent(parsed);
                        }
                    }
                    return;
                }
            }
            catch (Exception ex)
            {
                Global.WriteErrorLog($"[SoundAlerts] WebMessage parse error: {ex}");
            }
        }

        private void TryEmitEvent(Tuple<string, int, string> parsed)
        {
            if (parsed == null) return;
            if (string.IsNullOrEmpty(parsed.Item1) && parsed.Item2 == 0)
                return;

            // 去重
            string key = $"{parsed.Item1}|{parsed.Item2}|{parsed.Item3}";
            var now = DateTime.Now;
            if (key == _lastEventKey && (now - _lastEventTime) < _dedupeWindow)
                return;

            _lastEventKey = key;
            _lastEventTime = now;

            try { Global.WriteDebugLog($"[SoundAlerts] Event: user={parsed.Item1}, cost={parsed.Item2}, type={parsed.Item3}"); } catch { }

            _monitor?.AppendLogFromSoundAlerts(parsed.Item1, parsed.Item2.ToString(), parsed.Item3);
        }

        #region JavaScript Bridge

        private async Task InjectBridgeAsync()
        {
            if (_webView?.CoreWebView2 == null)
                return;

            string script = @"
(function () {
    try {
        if (window.__saBridgeInstalled) { return; }
        window.__saBridgeInstalled = true;

        function post(obj) {
            try { if (window.chrome && window.chrome.webview) window.chrome.webview.postMessage(obj); } catch (e) {}
        }

        function log(msg) { post({ kind: 'bridge-log', message: msg + ' @ ' + location.href }); }

        function safeClone(value) {
            try { return JSON.parse(JSON.stringify(value)); }
            catch (e) { try { return String(value); } catch (e2) { return null; } }
        }

        function tryParseJson(text) {
            try { return JSON.parse(text); } catch (e) { return null; }
        }

        // WebSocket hook
        try {
            var OrigWebSocket = window.WebSocket;
            window.WebSocket = function(url, protocols) {
                var ws = protocols ? new OrigWebSocket(url, protocols) : new OrigWebSocket(url);
                post({ kind: 'ws-open', href: location.href, url: String(url) });
                ws.addEventListener('message', function(evt) {
                    try {
                        var raw = evt && typeof evt.data !== 'undefined' ? evt.data : null;
                        var parsed = typeof raw === 'string' ? tryParseJson(raw) : null;
                        post({ kind: 'ws-message', href: location.href, url: String(url), raw: typeof raw === 'string' ? raw : null, json: parsed });
                    } catch (e) {}
                });
                return ws;
            };
            window.WebSocket.prototype = OrigWebSocket.prototype;
            log('WebSocket hooked');
        } catch (e) { log('WebSocket hook failed: ' + e); }

        // fetch hook
        try {
            var _origFetch = window.fetch;
            window.fetch = function() {
                var args = arguments;
                return _origFetch.apply(this, args).then(function(resp) {
                    try {
                        var reqUrl = args[0];
                        if (reqUrl && typeof reqUrl !== 'string' && reqUrl.url) reqUrl = reqUrl.url;
                        var clone = resp.clone();
                        clone.text().then(function(txt) {
                            var parsed = tryParseJson(txt);
                            post({ kind: 'fetch-response', href: location.href, url: String(reqUrl), status: resp.status, raw: txt, json: parsed });
                        }).catch(function(){});
                    } catch (e) {}
                    return resp;
                });
            };
            log('fetch hooked');
        } catch (e) { log('fetch hook failed: ' + e); }

        // XHR hook
        try {
            var _open = XMLHttpRequest.prototype.open;
            var _send = XMLHttpRequest.prototype.send;
            XMLHttpRequest.prototype.open = function(method, url) {
                this.__sa_method = method;
                this.__sa_url = url;
                return _open.apply(this, arguments);
            };
            XMLHttpRequest.prototype.send = function() {
                this.addEventListener('load', function() {
                    try {
                        var txt = this.responseText;
                        var parsed = tryParseJson(txt);
                        post({ kind: 'xhr-response', href: location.href, method: this.__sa_method || null, url: this.__sa_url || null, status: this.status, raw: txt, json: parsed });
                    } catch (e) {}
                });
                return _send.apply(this, arguments);
            };
            log('xhr hooked');
        } catch (e) { log('xhr hook failed: ' + e); }

        log('bridge installed');
    } catch (err) {
        try { if (window.chrome && window.chrome.webview) window.chrome.webview.postMessage({ kind: 'bridge-error', message: String(err && err.stack ? err.stack : err), href: location.href }); } catch (e) {}
    }
})();
";
            await _webView.CoreWebView2.ExecuteScriptAsync(script);
            try { Global.WriteDebugLog("[SoundAlerts] Bridge injected"); } catch { }
        }

        #endregion

        #region Parsing

        private Tuple<string, int, string> ParseSoundAlert(JObject json)
        {
            try
            {
                JToken fields = null;

                if (json["document"] != null && json["document"]["fields"] != null)
                    fields = json["document"]["fields"];
                else if (json["documentChange"] != null &&
                         json["documentChange"]["document"] != null &&
                         json["documentChange"]["document"]["fields"] != null)
                    fields = json["documentChange"]["document"]["fields"];

                if (fields == null)
                    return Tuple.Create("", 0, "");

                string lastUser = "";
                int cost = 0;
                string costType = "";

                if (fields["last_user"]?["stringValue"] != null)
                    lastUser = fields["last_user"]["stringValue"].ToString();

                if (fields["cost"]?["integerValue"] != null)
                    int.TryParse(fields["cost"]["integerValue"].ToString(), out cost);

                if (fields["cost_type"]?["stringValue"] != null)
                    costType = fields["cost_type"]["stringValue"].ToString();

                return Tuple.Create(lastUser, cost, costType);
            }
            catch
            {
                return Tuple.Create("", 0, "");
            }
        }

        private Tuple<string, int, string> ParseSoundAlertFromRawByRegex(string raw)
        {
            try
            {
                if (string.IsNullOrEmpty(raw))
                    return Tuple.Create("", 0, "");

                if (raw.IndexOf("/history/", StringComparison.OrdinalIgnoreCase) < 0)
                    return Tuple.Create("", 0, "");

                string lastUser = "";
                int cost = 0;
                string costType = "";

                var mUser = Regex.Match(raw,
                    @"""last_user""\s*:\s*\{\s*""stringValue""\s*:\s*""(?<v>[^""]*)""",
                    RegexOptions.Singleline);
                if (mUser.Success)
                    lastUser = mUser.Groups["v"].Value;

                var mCost = Regex.Match(raw,
                    @"""cost""\s*:\s*\{\s*""integerValue""\s*:\s*""(?<v>\d+)""",
                    RegexOptions.Singleline);
                if (mCost.Success)
                    int.TryParse(mCost.Groups["v"].Value, out cost);

                var mCostType = Regex.Match(raw,
                    @"""cost_type""\s*:\s*\{\s*""stringValue""\s*:\s*""(?<v>[^""]*)""",
                    RegexOptions.Singleline);
                if (mCostType.Success)
                    costType = mCostType.Groups["v"].Value;

                return Tuple.Create(lastUser, cost, costType);
            }
            catch
            {
                return Tuple.Create("", 0, "");
            }
        }

        #endregion
    }
}
