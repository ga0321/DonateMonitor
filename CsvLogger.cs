using System;
using System.IO;
using System.Text;

namespace DonateMonitor
{
    public static class CsvLogger
    {
        private static readonly object _lock = new object();

        public static void Write(
            string datetime,
            string type,
            string name,
            decimal amount,
            string currency,
            string message)
        {
            // 1️⃣ 依日期建立資料夾
            string dir = "csv";
            Directory.CreateDirectory(dir);

            // 2️⃣ CSV 檔案路徑
            string path = Path.Combine(dir, $"{DateTime.Now.ToString("yyyy-MM-dd")}.csv");

            // 3️⃣ CSV 標頭
            const string header = "時間,類型,帳號,種類,數量,訊息";

            // 4️⃣ 組 CSV 一行（處理逗號與引號）
            string line = string.Format(
                "{0},{1},{2},{3},{4},{5}",
                datetime,
                Escape(type),
                Escape(name),
                Escape(currency),
                amount,
                Escape(message)
            );

            lock (_lock)
            {
                bool fileExists = File.Exists(path);

                using (var sw = new StreamWriter(path, true, Encoding.UTF8))
                {
                    // 第一次建立檔案 → 寫標頭
                    if (!fileExists)
                    {
                        sw.WriteLine(header);
                    }

                    // 寫資料
                    sw.WriteLine(line);
                }
            }
        }

        // CSV 安全處理（有逗號 / 換行 / 雙引號時）
        private static string Escape(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
            {
                value = value.Replace("\"", "\"\"");
                return $"\"{value}\"";
            }

            return value;
        }
    }
}