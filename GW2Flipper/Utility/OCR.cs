namespace GW2Flipper.Utility;

using System.Drawing;
using System.Drawing.Imaging;

using global::GW2Flipper.Extensions;

using NLog;

using TesseractOCR;

internal static class OCR
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

    public static string ReadName(Bitmap bitmap, Color color)
    {
        using var engine = new Engine("./tessdata", "eng_best", TesseractOCR.Enums.EngineMode.Default);
        _ = engine.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890 + -—/ '\",():");

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

        return text;
    }

    public static bool NameCompare(string name, string compare, Dictionary<string, string> stringFixes)
    {
        var same = string.Equals(name, compare, StringComparison.OrdinalIgnoreCase);

        if (!same)
        {
            var fileName = Path.Combine("logs", "wrongnames.txt");
            var entry = $"Item name: {name} Captured name: {compare}";
            var isMatch = false;

            if (File.Exists(fileName))
            {
                foreach (var line in File.ReadLines(fileName))
                {
                    if (string.Equals(line, entry))
                    {
                        isMatch = true;
                        break;
                    }
                }
            }

            if (!isMatch && !compare.Contains(name))
            {
                File.AppendAllText(fileName, entry + "\n");

                var lines = File.ReadAllLines(fileName);
                var linesOrdered = lines.OrderBy(line => line).ToList();
                File.WriteAllLines(fileName, linesOrdered);
            }
        }

        var pretext = compare;
        compare = StringRepair(compare, stringFixes);
        if (pretext != compare)
        {
            Logger.Debug($"Repaired {pretext} to {compare}");
        }

        name = name.RemoveDiacritics().StripPunctuation().Replace(" ", string.Empty);
        compare = compare.StripPunctuation().Replace(" ", string.Empty);

        Logger.Debug($"Item name: {name} Captured name: {compare}");

        return string.Equals(name, compare, StringComparison.OrdinalIgnoreCase);
    }

    private static string StringRepair(string text, Dictionary<string, string> stringFixes)
    {
        text += '\n';

        foreach (var fix in stringFixes)
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
