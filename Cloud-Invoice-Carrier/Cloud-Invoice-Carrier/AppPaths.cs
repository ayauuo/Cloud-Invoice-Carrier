namespace Cloud_Invoice_Carrier;

/// <summary>
/// 單檔發佈時：ContentRoot 為解壓資源目錄，ExeDirectory 為 exe 所在目錄（.env 等使用者設定）。
/// </summary>
internal static class AppPaths
{
    public static string ExeDirectory { get; } = InitializeExeDirectory();

    public static string ContentRoot => AppContext.BaseDirectory;

    private static string InitializeExeDirectory()
    {
        var processPath = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(processPath))
            return Path.GetDirectoryName(processPath) ?? AppContext.BaseDirectory;

        return Application.StartupPath;
    }

    public static string? FindFile(string relativePath)
    {
        var normalized = NormalizeRelative(relativePath);
        foreach (var root in new[] { ExeDirectory, ContentRoot })
        {
            var path = Path.Combine(root, normalized);
            if (File.Exists(path))
                return path;
        }

        return null;
    }

    public static string CombineContent(string relativePath)
    {
        return Path.Combine(ContentRoot, NormalizeRelative(relativePath));
    }

    public static bool IsUnderAppRoots(string candidatePath)
    {
        var full = Path.GetFullPath(candidatePath);
        return IsUnderDirectory(ExeDirectory, full) || IsUnderDirectory(ContentRoot, full);
    }

    private static string NormalizeRelative(string relativePath) =>
        relativePath.Replace('/', Path.DirectorySeparatorChar).TrimStart(Path.DirectorySeparatorChar);

    private static bool IsUnderDirectory(string directory, string candidatePath)
    {
        var root = Path.GetFullPath(directory.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar);
        var full = Path.GetFullPath(candidatePath);
        return full.StartsWith(root, StringComparison.OrdinalIgnoreCase);
    }
}
