namespace Cloud_Invoice_Carrier;

/// <summary>依 picture/head、picture/back 目錄實際檔案解析版型圖路徑（支援 png、jpg 等）。</summary>
internal static class CarrierTemplateImageResolver
{
    public const int TemplateCount = 12;

    private static readonly string[] Extensions = [".png", ".jpg", ".jpeg", ".webp"];

    public static string[] ResolveHeadImages(string contentRoot) =>
        ResolveFolderImages("picture/head", TemplateCount, contentRoot);

    public static string[] ResolveBackImages(string contentRoot) =>
        ResolveFolderImages("picture/back", TemplateCount, contentRoot);

    private static string[] ResolveFolderImages(string folder, int count, string contentRoot)
    {
        var results = new string[count];
        for (var i = 0; i < count; i++)
        {
            var index = i + 1;
            results[i] = ResolveOne(folder, index, contentRoot) ?? $"{folder}/{index}.jpg";
        }

        return results;
    }

    private static string? ResolveOne(string folder, int index, string contentRoot)
    {
        foreach (var root in new[] { AppPaths.ExeDirectory, contentRoot })
        {
            foreach (var ext in Extensions)
            {
                var relative = $"{folder}/{index}{ext}";
                var path = Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar));
                if (File.Exists(path))
                    return relative;
            }
        }

        return null;
    }
}
