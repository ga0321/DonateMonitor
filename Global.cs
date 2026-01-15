using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DonateMonitor
{
    static class Setting
    {
        static public readonly string kECPAY_APIURL = "ECPAY_APIURL";
        static public readonly string kOPAY_APIURL = "OPAY_APIURL";
        static public readonly string kSTREAMLABS_KEY = "STREAMLABS_KEY";
        static public readonly string kOBS_OUTPUT_MODE = "OBS_OUTPUT_MODE";
        static public readonly string kOBS_ECPAY_OUTPUT_MSG = "OBS_ECPAY_OUTPUT_MSG";
        static public readonly string kOBS_OPAY_OUTPUT_MSG = "OBS_OPAY_OUTPUT_MSG";
        static public readonly string kOBS_STREAMLABS_PAYPAL_OUTPUT_MSG = "OBS_STREAMLABS_PAYPAL_OUTPUT_MSG";
        static public readonly string kOBS_STREAMLABS_BITS_OUTPUT_MSG = "OBS_STREAMLABS_BITS_OUTPUT_MSG";
        static public readonly string kOBS_STREAMLABS_SUBGIFT_OUTPUT_MSG = "OBS_STREAMLABS_SUB_GIFT_OUTPUT_MSG";
        static public readonly string kCUSTOM_ANON = "CUSTOM_ANON";
        static public readonly string kCUSTOM_SUB_TIER1 = "CUSTOM_SUB_TIER1";
        static public readonly string kCUSTOM_SUB_TIER2 = "CUSTOM_SUB_TIER2";
        static public readonly string kCUSTOM_SUB_TIER3 = "CUSTOM_SUB_TIER3";
        static public readonly string kCUSTOM_SUB_GIFT = "CUSTOM_SUB_GIFT";
        static public readonly string kCUSTOM_BITS = "CUSTOM_BITS";
        static public string Read(string sKey)
        {
            try
            {
                return ConfigurationManager.AppSettings[sKey];
            }
            catch
            {
                return null;
            }
        }
        static public void Save(string key, string value)
        {
            var config = ConfigurationManager.OpenExeConfiguration(
                ConfigurationUserLevel.None
            );

            if (config.AppSettings.Settings[key] == null)
                config.AppSettings.Settings.Add(key, value);
            else
                config.AppSettings.Settings[key].Value = value;

            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }
        static public void Reset()
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings.Clear();
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }
    }
    static class Global
    {
        static public bool _bExit = false;
        static private string _sECPAY_LoginToken = null;
        static private string _sECPAY_ListenKey = null;
        static private string _sOPAY_ListenKey = null;
        public struct VARS
        {
            public string _sStreamlabsKey;
            public int _nOBS_OutputMode;
            public string _sECPAY_OBS_Msg;
            public string _sOPAY_OBS_Msg;
            public string _sStreamlabs_Paypal_OBS_Msg;
            public string _sStreamlabs_Bits_OBS_Msg;
            public string _sStreamlabs_SubGift_OBS_Msg;
            public string _sCustom_ANON;
            public string _sCustom_Sub_Tier1;
            public string _sCustom_Sub_Tier2;
            public string _sCustom_Sub_Tier3;
            public string _sCustom_Sub_Gift;
            public string _sCustom_Bits;
        }
        static VARS _rVARS = new VARS();
        public static VARS _VARS = new VARS();

        #region GlobalVars
        public static string ECPAY_LoginToken
        {
            get => Volatile.Read(ref _sECPAY_LoginToken);
            set => Interlocked.Exchange(ref _sECPAY_LoginToken, value);
        }
        public static string ECPAY_ListenKey
        {
            get => Volatile.Read(ref _sECPAY_ListenKey);
            set => Interlocked.Exchange(ref _sECPAY_ListenKey, value);
        }
        public static string OPAY_ListenKey
        {
            get => Volatile.Read(ref _sOPAY_ListenKey);
            set => Interlocked.Exchange(ref _sOPAY_ListenKey, value);
        }
        public static string StreamlabsKey
        {
            get => Volatile.Read(ref _VARS._sStreamlabsKey);
            set => Interlocked.Exchange(ref _VARS._sStreamlabsKey, value);
        }
        public static int OBS_OutputMode
        {
            get => Volatile.Read(ref _VARS._nOBS_OutputMode);
            set => Interlocked.Exchange(ref _VARS._nOBS_OutputMode, value);
        }
        public static string ECPAY_OBS_Msg
        {
            get => Volatile.Read(ref _VARS._sECPAY_OBS_Msg);
            set => Interlocked.Exchange(ref _VARS._sECPAY_OBS_Msg, value);
        }
        public static string OPAY_OBS_Msg
        {
            get => Volatile.Read(ref _VARS._sOPAY_OBS_Msg);
            set => Interlocked.Exchange(ref _VARS._sOPAY_OBS_Msg, value);
        }
        public static string Streamlabs_Paypal_OBS_Msg
        {
            get => Volatile.Read(ref _VARS._sStreamlabs_Paypal_OBS_Msg);
            set => Interlocked.Exchange(ref _VARS._sStreamlabs_Paypal_OBS_Msg, value);
        }
        public static string Streamlabs_Bits_OBS_Msg
        {
            get => Volatile.Read(ref _VARS._sStreamlabs_Bits_OBS_Msg);
            set => Interlocked.Exchange(ref _VARS._sStreamlabs_Bits_OBS_Msg, value);
        }
        public static string Streamlabs_SubGift_OBS_Msg
        {
            get => Volatile.Read(ref _VARS._sStreamlabs_SubGift_OBS_Msg);
            set => Interlocked.Exchange(ref _VARS._sStreamlabs_SubGift_OBS_Msg, value);
        }
        public static string Custom_ANON
        {
            get => Volatile.Read(ref _VARS._sCustom_ANON);
            set => Interlocked.Exchange(ref _VARS._sCustom_ANON, value);
        }
        public static string Custom_Sub_Tier1
        {
            get => Volatile.Read(ref _VARS._sCustom_Sub_Tier1);
            set => Interlocked.Exchange(ref _VARS._sCustom_Sub_Tier1, value);
        }
        public static string Custom_Sub_Tier2
        {
            get => Volatile.Read(ref _VARS._sCustom_Sub_Tier2);
            set => Interlocked.Exchange(ref _VARS._sCustom_Sub_Tier2, value);
        }
        public static string Custom_Sub_Tier3
        {
            get => Volatile.Read(ref _VARS._sCustom_Sub_Tier3);
            set => Interlocked.Exchange(ref _VARS._sCustom_Sub_Tier3, value);
        }
        public static string Custom_Sub_Gift
        {
            get => Volatile.Read(ref _VARS._sCustom_Sub_Gift);
            set => Interlocked.Exchange(ref _VARS._sCustom_Sub_Gift, value);
        }
        public static string Custom_Bits
        {
            get => Volatile.Read(ref _VARS._sCustom_Bits);
            set => Interlocked.Exchange(ref _VARS._sCustom_Bits, value);
        }
        static public void InitSettings()
        {
            _rVARS._sStreamlabsKey = null;
            _rVARS._nOBS_OutputMode = 0;
            _rVARS._sECPAY_OBS_Msg = "{0}: {1}{2}";
            _rVARS._sOPAY_OBS_Msg = "{0}: {1}{2}";
            _rVARS._sStreamlabs_Paypal_OBS_Msg = "{0}: {1}{2}";
            _rVARS._sStreamlabs_Bits_OBS_Msg = "{0}: {1}{2}";
            _rVARS._sStreamlabs_SubGift_OBS_Msg = "{0}: {1}{2}({3})";
            _rVARS._sCustom_ANON = "匿名";
            _rVARS._sCustom_Sub_Tier1 = "層一";
            _rVARS._sCustom_Sub_Tier2 = "層二";
            _rVARS._sCustom_Sub_Tier3 = "層三";
            _rVARS._sCustom_Sub_Gift = "贈訂";
            _rVARS._sCustom_Bits = "小奇點";

            _VARS = _rVARS;
        }
        static public void LoadSettings()
        {
            string sVar = Setting.Read(Setting.kOBS_OUTPUT_MODE);
            if (!string.IsNullOrEmpty(sVar))
            {
                if (sVar.Equals("1"))
                    OBS_OutputMode = 1;
                else
                    OBS_OutputMode = 0;
            }
            else
                OBS_OutputMode = 0;

            sVar = Setting.Read(Setting.kSTREAMLABS_KEY);
            if (!string.IsNullOrEmpty(sVar))
                StreamlabsKey = sVar;

            sVar = Setting.Read(Setting.kOBS_ECPAY_OUTPUT_MSG);
            if (!string.IsNullOrEmpty(sVar))
                ECPAY_OBS_Msg = sVar;

            sVar = Setting.Read(Setting.kOBS_OPAY_OUTPUT_MSG);
            if (!string.IsNullOrEmpty(sVar))
                OPAY_OBS_Msg = sVar;

            sVar = Setting.Read(Setting.kOBS_STREAMLABS_PAYPAL_OUTPUT_MSG);
            if (!string.IsNullOrEmpty(sVar))
                Streamlabs_Paypal_OBS_Msg = sVar;

            sVar = Setting.Read(Setting.kOBS_STREAMLABS_BITS_OUTPUT_MSG);
            if (!string.IsNullOrEmpty(sVar))
                Streamlabs_Bits_OBS_Msg = sVar;

            sVar = Setting.Read(Setting.kOBS_STREAMLABS_SUBGIFT_OUTPUT_MSG);
            if (!string.IsNullOrEmpty(sVar))
                Streamlabs_SubGift_OBS_Msg = sVar;

            sVar = Setting.Read(Setting.kCUSTOM_ANON);
            if (!string.IsNullOrEmpty(sVar))
                Custom_ANON = sVar;

            sVar = Setting.Read(Setting.kCUSTOM_SUB_TIER1);
            if (!string.IsNullOrEmpty(sVar))
                Custom_Sub_Tier1 = sVar;

            sVar = Setting.Read(Setting.kCUSTOM_SUB_TIER2);
            if (!string.IsNullOrEmpty(sVar))
                Custom_Sub_Tier2 = sVar;

            sVar = Setting.Read(Setting.kCUSTOM_SUB_TIER3);
            if (!string.IsNullOrEmpty(sVar))
                Custom_Sub_Tier3 = sVar;

            sVar = Setting.Read(Setting.kCUSTOM_SUB_GIFT);
            if (!string.IsNullOrEmpty(sVar))
                Custom_Sub_Gift = sVar;

            sVar = Setting.Read(Setting.kCUSTOM_BITS);
            if (!string.IsNullOrEmpty(sVar))
                Custom_Bits = sVar;
        }
        static public void SaveSettings()
        {
            Setting.Save(Setting.kOBS_OUTPUT_MODE, OBS_OutputMode.ToString());
            Setting.Save(Setting.kOBS_ECPAY_OUTPUT_MSG, ECPAY_OBS_Msg);
            Setting.Save(Setting.kOBS_OPAY_OUTPUT_MSG, OPAY_OBS_Msg);
            Setting.Save(Setting.kOBS_STREAMLABS_PAYPAL_OUTPUT_MSG, Streamlabs_Paypal_OBS_Msg);
            Setting.Save(Setting.kOBS_STREAMLABS_BITS_OUTPUT_MSG, Streamlabs_Bits_OBS_Msg);
            Setting.Save(Setting.kOBS_STREAMLABS_SUBGIFT_OUTPUT_MSG, Streamlabs_SubGift_OBS_Msg);
            Setting.Save(Setting.kCUSTOM_ANON, Custom_ANON);
            Setting.Save(Setting.kCUSTOM_SUB_TIER1, Custom_Sub_Tier1);
            Setting.Save(Setting.kCUSTOM_SUB_TIER2, Custom_Sub_Tier2);
            Setting.Save(Setting.kCUSTOM_SUB_TIER3, Custom_Sub_Tier3);
            Setting.Save(Setting.kCUSTOM_SUB_GIFT, Custom_Sub_Gift);
            Setting.Save(Setting.kCUSTOM_BITS, Custom_Bits);
        }
        #endregion
        static public void WriteErrorLog(string msg)
        {
            string path = "errors.log";
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}";

            File.AppendAllText(path, line + Environment.NewLine);
        }
        static public void ShowError(string sMsg, bool bWriteLog = false)
        {
            if (bWriteLog)
                WriteErrorLog(sMsg);
            MessageBox.Show(sMsg, "DonateMonitor", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        static public void SetControlsEnabled(Control parent, bool enabled)
        {
            foreach (Control c in parent.Controls)
            {
                c.Enabled = enabled;

                if (c.HasChildren)
                    SetControlsEnabled(c, enabled);
            }
        }
        static public bool IsEnableECPAY()
        { 
            return !string.IsNullOrEmpty(ECPAY_LoginToken);
        }
        static public bool IsEnableOPAY()
        {
            return !string.IsNullOrEmpty(OPAY_ListenKey);
        }
        static public bool IsEnableStreamlabs()
        {
            return !string.IsNullOrEmpty(StreamlabsKey);
        }
        static public bool IsEnableAnyService()
        {
            return IsEnableECPAY() || IsEnableOPAY() || IsEnableStreamlabs();
        }
        static public bool IsEnableAllService()
        {
            return IsEnableECPAY() && IsEnableOPAY() && IsEnableStreamlabs();
        }
        static public string FormatAmount(string amount)
        {
            try
            {
                return Math.Truncate(Math.Round(decimal.Parse(amount), 2, MidpointRounding.AwayFromZero)).ToString();
            }
            catch
            {
                return amount; 
            }
        }
    }
}
