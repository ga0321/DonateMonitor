using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

namespace DonateMonitor
{
    public static class DonateDB
    {
        private static readonly string DbPath = "donate.db";
        private static readonly string ConnectionString = $"Data Source={DbPath};Version=3;Busy Timeout=60000;";

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
        /// 累計贈訂記錄：同一 Account+SubPlan 只保留一筆，Amount 累加。
        /// 回傳累計後的總數量。
        /// </summary>
        public static int AccumulateSubGift(
            string datetime,
            string account,
            string displayName,
            int amount,
            string subPlan)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    string type = Global.Custom_Sub_Gift;

                    // 取得目前累計
                    int currentTotal = 0;
                    string sumSql = "SELECT COALESCE(SUM(Amount), 0) FROM DonateLog WHERE Account = @account AND Type = @type AND SubPlan = @subPlan";
                    using (var cmd = new SQLiteCommand(sumSql, conn, transaction))
                    {
                        cmd.Parameters.AddWithValue("@account", account ?? "");
                        cmd.Parameters.AddWithValue("@type", type);
                        cmd.Parameters.AddWithValue("@subPlan", subPlan ?? "");
                        currentTotal = Convert.ToInt32(cmd.ExecuteScalar());
                    }

                    // 刪除舊的行（可能有多筆歷史資料）
                    string deleteSql = "DELETE FROM DonateLog WHERE Account = @account AND Type = @type AND SubPlan = @subPlan";
                    using (var cmd = new SQLiteCommand(deleteSql, conn, transaction))
                    {
                        cmd.Parameters.AddWithValue("@account", account ?? "");
                        cmd.Parameters.AddWithValue("@type", type);
                        cmd.Parameters.AddWithValue("@subPlan", subPlan ?? "");
                        cmd.ExecuteNonQuery();
                    }

                    // 寫入合併後的單筆記錄
                    int newTotal = currentTotal + amount;
                    string insertSql = @"
                        INSERT INTO DonateLog (DateTime, Type, Account, DisplayName, Amount, Currency, Message, SubPlan)
                        VALUES (@datetime, @type, @account, @displayName, @amount, @currency, @message, @subPlan)";
                    using (var cmd = new SQLiteCommand(insertSql, conn, transaction))
                    {
                        cmd.Parameters.AddWithValue("@datetime", datetime);
                        cmd.Parameters.AddWithValue("@type", type);
                        cmd.Parameters.AddWithValue("@account", account ?? "");
                        cmd.Parameters.AddWithValue("@displayName", displayName ?? "");
                        cmd.Parameters.AddWithValue("@amount", newTotal);
                        cmd.Parameters.AddWithValue("@currency", type);
                        cmd.Parameters.AddWithValue("@message", "");
                        cmd.Parameters.AddWithValue("@subPlan", subPlan ?? "");
                        cmd.ExecuteNonQuery();
                    }

                    transaction.Commit();
                    return newTotal;
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

        /// <summary>
        /// 取得所有記錄（用於 DataGridView 顯示）
        /// </summary>
        public static List<DonateRecord> GetAllRecords()
        {
            var result = new List<DonateRecord>();
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                string sql = "SELECT Id, DateTime, Type, Account, DisplayName, Amount, Currency, Message, SubPlan FROM DonateLog ORDER BY Id DESC";
                using (var cmd = new SQLiteCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(new DonateRecord
                        {
                            Id = Convert.ToInt32(reader["Id"]),
                            DateTime = reader["DateTime"]?.ToString() ?? "",
                            Type = reader["Type"]?.ToString() ?? "",
                            Account = reader["Account"]?.ToString() ?? "",
                            DisplayName = reader["DisplayName"]?.ToString() ?? "",
                            Amount = Convert.ToDecimal(reader["Amount"]),
                            Currency = reader["Currency"]?.ToString() ?? "",
                            Message = reader["Message"]?.ToString() ?? "",
                            SubPlan = reader["SubPlan"]?.ToString() ?? ""
                        });
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 依 ID 刪除記錄
        /// </summary>
        public static void DeleteById(int id)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                string sql = "DELETE FROM DonateLog WHERE Id = @id";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 新增一筆空白記錄並回傳新 ID
        /// </summary>
        public static int Insert(string datetime, string type, string account, string displayName, decimal amount, string currency, string message, string subPlan)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                string sql = @"
                    INSERT INTO DonateLog (DateTime, Type, Account, DisplayName, Amount, Currency, Message, SubPlan)
                    VALUES (@datetime, @type, @account, @displayName, @amount, @currency, @message, @subPlan);
                    SELECT last_insert_rowid();";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@datetime", datetime ?? "");
                    cmd.Parameters.AddWithValue("@type", type ?? "");
                    cmd.Parameters.AddWithValue("@account", account ?? "");
                    cmd.Parameters.AddWithValue("@displayName", displayName ?? "");
                    cmd.Parameters.AddWithValue("@amount", (double)amount);
                    cmd.Parameters.AddWithValue("@currency", currency ?? "");
                    cmd.Parameters.AddWithValue("@message", message ?? "");
                    cmd.Parameters.AddWithValue("@subPlan", subPlan ?? "");
                    return Convert.ToInt32(cmd.ExecuteScalar());
                }
            }
        }

        /// <summary>
        /// 依 ID 更新記錄
        /// </summary>
        public static void UpdateById(int id, string datetime, string type, string account, string displayName, decimal amount, string currency, string message, string subPlan)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                string sql = @"
                    UPDATE DonateLog SET
                        DateTime = @datetime,
                        Type = @type,
                        Account = @account,
                        DisplayName = @displayName,
                        Amount = @amount,
                        Currency = @currency,
                        Message = @message,
                        SubPlan = @subPlan
                    WHERE Id = @id";
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.Parameters.AddWithValue("@datetime", datetime ?? "");
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
    }

    /// <summary>
    /// 捐贈記錄資料模型
    /// </summary>
    public class DonateRecord
    {
        public int Id { get; set; }
        public string DateTime { get; set; }
        public string Type { get; set; }
        public string Account { get; set; }
        public string DisplayName { get; set; }
        public decimal Amount { get; set; }
        public string Currency { get; set; }
        public string Message { get; set; }
        public string SubPlan { get; set; }
    }
}
