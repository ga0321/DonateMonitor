using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

namespace DonateMonitor
{
    public static class DonateDB
    {
        private static readonly string DbPath = "donate.db";
        private static readonly string ConnectionString = $"Data Source={DbPath};Version=3;";

        static DonateDB()
        {
            InitializeDatabase();
        }

        private static void InitializeDatabase()
        {
            if (!File.Exists(DbPath))
            {
                SQLiteConnection.CreateFile(DbPath);
            }

            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();

                string sql = @"
                    CREATE TABLE IF NOT EXISTS DonateLog (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        DateTime TEXT NOT NULL,
                        Type TEXT NOT NULL,
                        Account TEXT,
                        DisplayName TEXT,
                        Amount REAL,
                        Currency TEXT,
                        Message TEXT,
                        SubPlan TEXT,
                        CreatedAt TEXT DEFAULT (datetime('now', 'localtime'))
                    )";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 寫入捐贈記錄
        /// </summary>
        public static void Write(
            string datetime,
            string type,
            string account,
            string displayName,
            decimal amount,
            string currency,
            string message,
            string subPlan = null)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                string sql = @"
                    INSERT INTO DonateLog (DateTime, Type, Account, DisplayName, Amount, Currency, Message, SubPlan)
                    VALUES (@datetime, @type, @account, @displayName, @amount, @currency, @message, @subPlan)";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@datetime", datetime);
                    cmd.Parameters.AddWithValue("@type", type ?? "");
                    cmd.Parameters.AddWithValue("@account", account ?? "");
                    cmd.Parameters.AddWithValue("@displayName", displayName ?? "");
                    cmd.Parameters.AddWithValue("@amount", (double)amount);
                    cmd.Parameters.AddWithValue("@currency", currency ?? "");
                    cmd.Parameters.AddWithValue("@message", message ?? "");
                    cmd.Parameters.AddWithValue("@subPlan", subPlan ?? "");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 獲取指定帳號和類型的累計金額
        /// </summary>
        public static decimal GetTotalAmount(string account, string type)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                string sql = "SELECT COALESCE(SUM(Amount), 0) FROM DonateLog WHERE Account = @account AND Type = @type";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@account", account);
                    cmd.Parameters.AddWithValue("@type", type);
                    var result = cmd.ExecuteScalar();
                    return Convert.ToDecimal(result);
                }
            }
        }

        /// <summary>
        /// 獲取指定帳號的贈訂累計數量
        /// </summary>
        public static int GetSubGiftCount(string account)
        {
            return (int)GetTotalAmount(account, Global.Custom_Sub_Gift);
        }

        /// <summary>
        /// 獲取指定帳號和層級的贈訂累計數量
        /// </summary>
        public static int GetSubGiftCountByPlan(string account, string subPlan)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                string sql = "SELECT COALESCE(SUM(Amount), 0) FROM DonateLog WHERE Account = @account AND Type = @type AND SubPlan = @subPlan";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@account", account);
                    cmd.Parameters.AddWithValue("@type", Global.Custom_Sub_Gift);
                    cmd.Parameters.AddWithValue("@subPlan", subPlan ?? "");
                    var result = cmd.ExecuteScalar();
                    return Convert.ToInt32(result);
                }
            }
        }

        /// <summary>
        /// 清除指定類型的所有記錄
        /// </summary>
        public static void ClearByType(string type)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                string sql = "DELETE FROM DonateLog WHERE Type = @type";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@type", type);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 清除所有贈訂記錄
        /// </summary>
        public static void ClearAllSubGift()
        {
            ClearByType(Global.Custom_Sub_Gift);
        }

        /// <summary>
        /// 清除所有記錄
        /// </summary>
        public static void ClearAll()
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                string sql = "DELETE FROM DonateLog";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 取得所有累計資料 (用於程式啟動時載入)
        /// </summary>
        /// <returns>List of (Account, DisplayName, Type, TotalAmount, Currency, SubPlan)</returns>
        public static List<(string Account, string DisplayName, string Type, decimal TotalAmount, string Currency, string SubPlan)> GetAllCumulativeData()
        {
            var result = new List<(string, string, string, decimal, string, string)>();
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                // 依照 Account, Type, SubPlan 分組取得累計金額，同時取得最新的 DisplayName 和 Currency
                string sql = @"
                    SELECT Account, Type, SubPlan, SUM(Amount) as TotalAmount,
                           (SELECT DisplayName FROM DonateLog d2 WHERE d2.Account = d1.Account AND d2.Type = d1.Type AND COALESCE(d2.SubPlan,'') = COALESCE(d1.SubPlan,'') ORDER BY d2.Id DESC LIMIT 1) as DisplayName,
                           (SELECT Currency FROM DonateLog d2 WHERE d2.Account = d1.Account AND d2.Type = d1.Type AND COALESCE(d2.SubPlan,'') = COALESCE(d1.SubPlan,'') ORDER BY d2.Id DESC LIMIT 1) as Currency
                    FROM DonateLog d1
                    GROUP BY Account, Type, SubPlan";
                using (var cmd = new SQLiteCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add((
                            reader["Account"]?.ToString() ?? "",
                            reader["DisplayName"]?.ToString() ?? "",
                            reader["Type"]?.ToString() ?? "",
                            Convert.ToDecimal(reader["TotalAmount"]),
                            reader["Currency"]?.ToString() ?? "",
                            reader["SubPlan"]?.ToString() ?? ""
                        ));
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 匯出所有記錄為 CSV 檔案
        /// </summary>
        /// <returns>匯出的檔案路徑，如果沒有記錄則返回 null</returns>
        public static string ExportToCsv()
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                string sql = "SELECT DateTime, Type, Account, DisplayName, Amount, Currency, Message, SubPlan FROM DonateLog ORDER BY Id";
                using (var cmd = new SQLiteCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.HasRows)
                        return null;

                    string fileName = $"donate_backup_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                    using (var writer = new StreamWriter(fileName, false, System.Text.Encoding.UTF8))
                    {
                        // 寫入標題列
                        writer.WriteLine("DateTime,Type,Account,DisplayName,Amount,Currency,Message,SubPlan");

                        // 寫入資料列
                        while (reader.Read())
                        {
                            string line = string.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\",{4},\"{5}\",\"{6}\",\"{7}\"",
                                EscapeCsvField(reader["DateTime"]?.ToString()),
                                EscapeCsvField(reader["Type"]?.ToString()),
                                EscapeCsvField(reader["Account"]?.ToString()),
                                EscapeCsvField(reader["DisplayName"]?.ToString()),
                                reader["Amount"],
                                EscapeCsvField(reader["Currency"]?.ToString()),
                                EscapeCsvField(reader["Message"]?.ToString()),
                                EscapeCsvField(reader["SubPlan"]?.ToString()));
                            writer.WriteLine(line);
                        }
                    }
                    return fileName;
                }
            }
        }

        private static string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "";
            return field.Replace("\"", "\"\"");
        }
    }
}
