using System.Globalization;
using Cloud_Invoice_Carrier.Models;
using Microsoft.Data.Sqlite;

namespace Cloud_Invoice_Carrier;

/// <summary>
/// 本機 SQLite 列印／收款紀錄（參考 puddingDog01 PhotoDetailStore）。
/// 預設路徑：程式目錄 report/carrier_records.db
/// </summary>
internal static class CarrierRecordStore
{
    private const string ReportFolderName = "report";
    private const string DefaultDbFileName = "carrier_records.db";

    private static string GetDbPath()
    {
        try
        {
            var configPath = Path.Combine(AppPaths.ExeDirectory, "carrier_db_path.txt");
            if (File.Exists(configPath))
            {
                var line = File.ReadAllText(configPath).Trim();
                if (!string.IsNullOrWhiteSpace(line))
                    return line.Trim();
            }
        }
        catch
        {
            // 沿用預設
        }

        var reportDir = Path.Combine(AppPaths.ExeDirectory, ReportFolderName);
        Directory.CreateDirectory(reportDir);
        return Path.Combine(reportDir, DefaultDbFileName);
    }

    private static void EnsureTable(SqliteConnection conn)
    {
        const string sql = @"
CREATE TABLE IF NOT EXISTS PrintRecord (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Date TEXT NOT NULL,
    Time TEXT NOT NULL,
    ProjectName TEXT NOT NULL,
    MachineName TEXT NOT NULL,
    Amount INTEGER NOT NULL,
    TemplateName TEXT NOT NULL,
    CreatedAt TEXT NOT NULL,
    Copies INTEGER NOT NULL DEFAULT 1,
    IsTest INTEGER NOT NULL DEFAULT 0
);
";
        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        cmd.ExecuteNonQuery();

        using var idx = conn.CreateCommand();
        idx.CommandText = "CREATE INDEX IF NOT EXISTS IX_PrintRecord_Date ON PrintRecord(Date)";
        idx.ExecuteNonQuery();
    }

    public static string? InsertRecord(
        string templateName,
        DateTime whenLocal,
        int amount,
        string projectName,
        string machineName,
        int copies = 1,
        bool isTest = false)
    {
        try
        {
            var path = GetDbPath();
            using var conn = new SqliteConnection($"Data Source={path}");
            conn.Open();
            EnsureTable(conn);

            var dateStr = whenLocal.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            var timeStr = whenLocal.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
            copies = Math.Clamp(copies, 0, 99);

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
INSERT INTO PrintRecord (Date, Time, ProjectName, MachineName, Amount, TemplateName, CreatedAt, Copies, IsTest)
VALUES (@date, @time, @projectName, @machineName, @amount, @templateName, @createdAt, @copies, @isTest);
";
            cmd.Parameters.AddWithValue("@date", dateStr);
            cmd.Parameters.AddWithValue("@time", timeStr);
            cmd.Parameters.AddWithValue("@projectName", projectName ?? "");
            cmd.Parameters.AddWithValue("@machineName", machineName ?? "");
            cmd.Parameters.AddWithValue("@amount", amount);
            cmd.Parameters.AddWithValue("@templateName", templateName ?? "");
            cmd.Parameters.AddWithValue("@createdAt", DateTime.UtcNow.ToString("o"));
            cmd.Parameters.AddWithValue("@copies", copies);
            cmd.Parameters.AddWithValue("@isTest", isTest ? 1 : 0);
            cmd.ExecuteNonQuery();
            return null;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }

    public static (List<PrintRecordViewRow> Rows, int TotalPrintSheets, int TotalTestSheets, int TotalAmount, int TotalCount) GetPrintRecordsForView(
        DateTime date,
        string rangeType)
    {
        var rows = new List<PrintRecordViewRow>();
        var totalPrintSheets = 0;
        var totalTestSheets = 0;
        var totalAmount = 0;
        var totalCount = 0;

        try
        {
            var path = GetDbPath();
            if (!File.Exists(path))
                return (rows, 0, 0, 0, 0);

            var dateStr = date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            var dateFilter = rangeType switch
            {
                "week" => "Date >= @start AND Date <= @end",
                "month" => "Date LIKE @monthPrefix",
                _ => "Date = @dateStr"
            };

            var sql = $@"
SELECT Id, Date, Time, ProjectName, MachineName, Amount, TemplateName, Copies, IsTest
FROM PrintRecord
WHERE {dateFilter}
ORDER BY Date, Time, Id";

            using var conn = new SqliteConnection($"Data Source={path}");
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@dateStr", dateStr);

            if (rangeType == "week")
            {
                var start = date;
                while (start.DayOfWeek != DayOfWeek.Monday)
                    start = start.AddDays(-1);
                var end = start.AddDays(6);
                cmd.Parameters.AddWithValue("@start", start.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                cmd.Parameters.AddWithValue("@end", end.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            }

            if (rangeType == "month")
                cmd.Parameters.AddWithValue("@monthPrefix", date.ToString("yyyy-MM", CultureInfo.InvariantCulture) + "%");

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var copies = reader.IsDBNull(7) ? 1 : reader.GetInt32(7);
                var isTest = !reader.IsDBNull(8) && reader.GetInt32(8) != 0;
                var amount = reader.GetInt32(5);
                totalPrintSheets += copies;
                totalAmount += amount;
                totalCount += 1;
                if (isTest)
                    totalTestSheets += copies;

                rows.Add(new PrintRecordViewRow
                {
                    Id = reader.GetInt64(0),
                    Date = reader.GetString(1),
                    Time = reader.GetString(2),
                    ProjectName = reader.GetString(3),
                    MachineName = reader.GetString(4),
                    Amount = amount,
                    TemplateName = reader.GetString(6),
                    Copies = copies,
                    IsTest = isTest
                });
            }
        }
        catch
        {
            // 回傳空
        }

        return (rows, totalPrintSheets, totalTestSheets, totalAmount, totalCount);
    }

    public static string GetDbPathForDisplay() => GetDbPath();
}
