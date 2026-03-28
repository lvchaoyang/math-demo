using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

static int Usage()
{
    Console.Error.WriteLine("Usage: WmfGdiRender <input.wmf|emf> <output.png> [--dpi <72-1200>]");
    Console.Error.WriteLine("  Renders a standalone WMF/EMF using GDI+ (Windows).");
    return 1;
}

static float ParseDpi(string[] args, ref int i)
{
    if (i + 1 >= args.Length) return -1;
    if (!float.TryParse(args[i + 1], out var d)) return -1;
    i++;
    if (d < 72) d = 72;
    if (d > 1200) d = 1200;
    return d;
}

if (args.Length < 2)
    Environment.Exit(Usage());

var inputPath = args[0];
var outputPath = args[1];
float dpi = 300f;

for (var i = 2; i < args.Length; i++)
{
    if (args[i] == "--dpi")
    {
        var d = ParseDpi(args, ref i);
        if (d < 0) Environment.Exit(Usage());
        dpi = d;
    }
}

if (!File.Exists(inputPath))
{
    Console.Error.WriteLine($"Input not found: {inputPath}");
    Environment.Exit(2);
}

try
{
    using var mf = new Metafile(inputPath);
    // .NET Metafile Width/Height: hundredths of a millimeter (typical for metafile frame)
    float wMm = mf.Width / 100f;
    float hMm = mf.Height / 100f;
    if (wMm <= 0 || hMm <= 0)
    {
        Console.Error.WriteLine("Invalid metafile dimensions in header.");
        Environment.Exit(3);
    }

    int w = Math.Max(1, (int)Math.Ceiling(wMm * dpi / 25.4f));
    int h = Math.Max(1, (int)Math.Ceiling(hMm * dpi / 25.4f));

    using var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
    bmp.SetResolution(dpi, dpi);
    using (var g = Graphics.FromImage(bmp))
    {
        g.Clear(Color.White);
        g.PageUnit = GraphicsUnit.Pixel;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
        g.DrawImage(mf, 0, 0, w, h);
    }

    Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? ".");
    bmp.Save(outputPath, ImageFormat.Png);
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex.Message);
    Environment.Exit(4);
}
