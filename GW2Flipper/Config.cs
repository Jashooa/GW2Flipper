namespace GW2Flipper;

using System.Text.Json;

internal class Config
{
    private const string CONFIG = "config.json";
    private const string BLACKLIST = "blacklist.json";
    private const string OCRFIXES = "ocrfixes.json";

    public static Config? Instance { get; private set; }

    public string ApiKey { get; set; } = string.Empty;

    public string CharacterName { get; set; } = string.Empty;

    public int BuysPerSellLoop { get; set; }

    public double Quantity { get; set; }

    public int MaxSpend { get; set; }

    public bool UndercutBuys { get; set; }

    public bool UndercutSells { get; set; }

    public bool BuyIfSelling { get; set; }

    public bool SellForBestIfUnderRange { get; set; }

    public bool IgnoreSellsLessThanBuys { get; set; }

    public double SellsLessThanBuysRange { get; set; }

    public bool RemoveUnprofitable { get; set; }

    public double MinStringSimilarity { get; set; }

    public double ErrorRange { get; set; }

    public double ErrorRangeInverse { get; set; }

    public double ProfitRange { get; set; }

    public int UpdateListTime { get; set; }

    public List<Dictionary<string, string>> Arguments { get; set; } = null!;

    public List<int> Blacklist { get; set; } = null!;

    public Dictionary<string, string> OcrFixes { get; set; } = null!;

    public static void Load()
    {
        var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, CONFIG);
        var blacklistPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, BLACKLIST);
        var ocrfixesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, OCRFIXES);

        if (!File.Exists(configPath))
        {
            throw new FileNotFoundException($"{CONFIG} is missing.");
        }

        if (!File.Exists(blacklistPath))
        {
            throw new FileNotFoundException($"{BLACKLIST} is missing.");
        }

        if (!File.Exists(ocrfixesPath))
        {
            throw new FileNotFoundException($"{OCRFIXES} is missing.");
        }

        var options = new JsonSerializerOptions
        {
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true,
        };

        Instance = JsonSerializer.Deserialize<Config>(File.ReadAllText(configPath), options);

        if (Instance == null)
        {
            throw new Exception("Couldn't load config");
        }

        Instance.ErrorRangeInverse = 1 + 1 - Instance.ErrorRange;
        Instance.Blacklist = JsonSerializer.Deserialize<List<int>>(File.ReadAllText(blacklistPath), options)!;
        Instance.OcrFixes = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(ocrfixesPath), options)!;
    }
}
