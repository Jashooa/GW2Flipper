namespace GW2Flipper.Utility;

using System.Drawing;
using System.Drawing.Imaging;

using global::GW2Flipper.Extensions;

using NLog;

using TesseractOCR;

internal static class OCR
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

    private static readonly Dictionary<string, string> StringFixes = new()
    {
        { "Uscku's", "Usoku's" },
        { " Qil", " Oil" },
        { "lcebrood", "Icebrood" },
        { "Baim", "Balm" },
        { "Cre\n", "Ore\n" },
        { "Axiquiot\n", "Axiquiotl\n" },
        { "Piguant", "Piquant" },
        { "Grawil", "Grawl" },
    };

    public static string ReadName(Bitmap bitmap, Color color)
    {
        using var engine = new Engine("./tessdata", "eng_best", TesseractOCR.Enums.EngineMode.Default);
        _ = engine.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890 + -—/ '\",()");

        var ms = new MemoryStream();
        bitmap.BinarizeByColor(color, 0.55).Save(ms, ImageFormat.Bmp);
        var image = TesseractOCR.Pix.Image.LoadFromMemory(ms).Scale(3.125f, 3.125f);

        using var page = engine.Process(image, TesseractOCR.Enums.PageSegMode.SingleBlock);

        var text = page.Text.Replace("\n", " ").Replace("\r", string.Empty).Trim();

        if (page.MeanConfidence < 0.1)
        {
            Logger.Debug($"Mean confidence: {page.MeanConfidence} !!!!!!!!!");
            // image.Save($"./logs/images/{DateTime.Now:HH-mm-ss-ffff}_OCR.png", TesseractOCR.Enums.ImageFormat.Png);
        }
        else if (page.MeanConfidence < 0.9)
        {
            Logger.Debug($"Mean confidence: {page.MeanConfidence} !!!");
            // image.Save($"./logs/images/{DateTime.Now:HH-mm-ss-ffff}_OCR.png", TesseractOCR.Enums.ImageFormat.Png);
        }
        else
        {
            Logger.Debug($"Mean confidence: {page.MeanConfidence}");
        }

        var pretext = text;
        text = StringRepair(text);
        if (pretext != text)
        {
            Logger.Debug($"Repaired {pretext} to {text}");
        }

        return text;
    }

    private static string StringRepair(string text)
    {
        text += '\n';

        foreach (var fix in StringFixes)
        {
            text = text.Replace(fix.Key, fix.Value);
        }

        return text.TrimEnd();
    }

    /*public static string ReadNameIron(Bitmap bitmap, Color color)
    {
        var ocr = new IronTesseract();
        // ocr.Configuration.WhiteListCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890 +-/'\",()";
        ocr.Configuration.PageSegmentationMode = TesseractPageSegmentationMode.SingleBlock;
        ocr.Configuration.TesseractVariables["debug_file"] = "./logs/tesseract.log";

        using var input = new OcrInput(bitmap.BinarizeByColor(color, 0.6));

        var result = ocr.Read(input);
        var text = result.Text.Replace("\n", " ").Replace("\r", string.Empty).Trim();

        if (result.Confidence < 90)
        {
            Logger.Debug($"Mean confidence: {result.Confidence} !!!!!!!!!");
            // _ = input.SaveAsImages($"{DateTime.Now:HH-mm-ss-ffff}");
        }
        else
        {
            Logger.Debug($"Mean confidence: {result.Confidence}");
        }

        Logger.Debug($"Text: {text}");

        return text;
    }

    public static string ReadNumberIron(Bitmap bitmap)
    {
        var ocr = new IronTesseract();
        ocr.Configuration.WhiteListCharacters = "1234567890";
        ocr.Configuration.PageSegmentationMode = TesseractPageSegmentationMode.SingleBlock;
        ocr.Configuration.TesseractVariables["classify_bln_numeric_mode"] = "1";
        ocr.Configuration.TesseractVariables["debug_file"] = "./logs/tesseract.log";

        using var input = new OcrInput(bitmap);

        _ = input.EnhanceResolution(300);
        _ = input.Contrast();
        _ = input.Invert();

        var result = ocr.Read(input);

        var text = result.Text;

        if (result.Confidence < 90)
        {
            Logger.Debug($"Mean confidence: {result.Confidence} !!!!!!!!!");
            // _ = input.SaveAsImages($"{DateTime.Now:HH-mm-ss-ffff}");
        }
        else
        {
            Logger.Debug($"Mean confidence: {result.Confidence}");
        }

        Logger.Debug($"Text: {text}");

        return text;
    }*/
}
