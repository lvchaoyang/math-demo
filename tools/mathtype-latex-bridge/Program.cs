using System.Text;
using System.Text.RegularExpressions;

static int Usage()
{
    Console.Error.WriteLine("Usage: MathTypeLatexBridge --ole <path> [--mode sdk|heuristic]");
    Console.Error.WriteLine("  --ole   Path to MathType OLE .bin file");
    Console.Error.WriteLine("  --mode  sdk (default) or heuristic");
    return 1;
}

static string? GetArg(string[] args, string name)
{
    for (var i = 0; i < args.Length - 1; i++)
    {
        if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))
            return args[i + 1];
    }
    return null;
}

static bool HasArg(string[] args, string name)
{
    return args.Any(a => string.Equals(a, name, StringComparison.OrdinalIgnoreCase));
}

static bool LooksLikeLatex(string s)
{
    if (string.IsNullOrWhiteSpace(s) || s.Length < 2) return false;
    var score = 0;
    if (s.Contains('\\')) score += 1;
    if (Regex.IsMatch(s, @"\\[a-zA-Z]+")) score += 2;
    if (s.Contains('{') && s.Contains('}')) score += 1;
    if (s.Contains('^') || s.Contains('_')) score += 1;
    return score >= 3;
}

static bool IsNoisy(string s)
{
    var t = s.ToLowerInvariant();
    var markers = new[]
    {
        "design science",
        "teX input language".ToLowerInvariant(),
        "winallbasiccodepages",
        "winallcodepages",
        "times new roman",
        "courier new",
        "mt extra",
        "dsmt",
    };
    return markers.Any(m => t.Contains(m));
}

static string? TryHeuristic(byte[] data)
{
    var encodings = new[] { Encoding.UTF8, Encoding.Unicode, Encoding.Latin1 };
    foreach (var enc in encodings)
    {
        var text = enc.GetString(data);
        if (string.IsNullOrWhiteSpace(text)) continue;

        foreach (Match m in Regex.Matches(text, @"\$\$(.{1,2000}?)\$\$", RegexOptions.Singleline))
        {
            var seg = m.Groups[1].Value.Trim();
            if (LooksLikeLatex(seg) && !IsNoisy(seg)) return seg;
        }
        foreach (Match m in Regex.Matches(text, @"(?<!\$)\$(.{1,1000}?)(?<!\$)\$", RegexOptions.Singleline))
        {
            var seg = m.Groups[1].Value.Trim();
            if (LooksLikeLatex(seg) && !IsNoisy(seg)) return seg;
        }
        foreach (Match m in Regex.Matches(text, @"([^\r\n]{3,2000})"))
        {
            var seg = m.Groups[1].Value.Trim();
            if (LooksLikeLatex(seg) && !IsNoisy(seg)) return seg;
        }
    }
    return null;
}

static string? TrySdk(string olePath)
{
    // TODO: 在这里接入你们的 MathType SDK 调用（建议封装成独立类）
    // 约定：成功返回纯 LaTeX 文本；失败返回 null。
    //
    // 可选实现方式：
    // 1) 通过 MathType SDK 提供的 API（C/.NET）将 OLE/MTEF 转 TeX
    // 2) 若你们已有内部 DLL/COM 封装，在此直接调用
    return null;
}

if (args.Length == 0 || HasArg(args, "--help") || HasArg(args, "-h"))
    Environment.Exit(Usage());

var olePath = GetArg(args, "--ole");
if (string.IsNullOrWhiteSpace(olePath))
    Environment.Exit(Usage());

if (!File.Exists(olePath))
{
    Console.Error.WriteLine($"Input not found: {olePath}");
    Environment.Exit(2);
}

var mode = (GetArg(args, "--mode") ?? "sdk").Trim().ToLowerInvariant();

try
{
    string? latex = null;
    if (mode == "sdk")
    {
        latex = TrySdk(olePath);
        if (string.IsNullOrWhiteSpace(latex))
        {
            // SDK 未接好时，给出明确错误，便于 Parser 分类统计
            Console.Error.WriteLine("sdk_not_implemented_or_failed");
            Environment.Exit(10);
        }
    }
    else if (mode == "heuristic")
    {
        var data = File.ReadAllBytes(olePath);
        latex = TryHeuristic(data);
        if (string.IsNullOrWhiteSpace(latex))
        {
            Console.Error.WriteLine("heuristic_no_candidate");
            Environment.Exit(11);
        }
    }
    else
    {
        Console.Error.WriteLine($"unsupported_mode:{mode}");
        Environment.Exit(12);
    }

    Console.OutputEncoding = Encoding.UTF8;
    Console.WriteLine(latex!.Trim());
    Environment.Exit(0);
}
catch (Exception ex)
{
    Console.Error.WriteLine($"bridge_exception:{ex.Message}");
    Environment.Exit(20);
}
