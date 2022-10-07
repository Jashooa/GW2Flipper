namespace GW2Flipper;

using System.Text.Json;

internal class Config
{
    private const string PATH = "config.json";
    private const string BLACKLIST = "blacklist.json";

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

    public bool RemoveUnprofitable { get; set; }

    public double MinStringSimilarity { get; set; }

    public double ErrorRange { get; set; }

    public double ErrorRangeInverse { get; set; }

    public double ProfitRange { get; set; }

    public int UpdateListTime { get; set; }

    public List<Dictionary<string, string>> Arguments { get; set; } = null!;

    public List<int> Blacklist { get; set; } = null!;

    public static void Load()
    {
        if (!File.Exists(PATH))
        {
            throw new FileNotFoundException($"{PATH} is missing.");
        }

        if (!File.Exists(BLACKLIST))
        {
            throw new FileNotFoundException($"{BLACKLIST} is missing.");
        }

        var options = new JsonSerializerOptions
        {
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true,
        };

        Instance = JsonSerializer.Deserialize<Config>(File.ReadAllText(PATH), options);

        if (Instance == null)
        {
            throw new Exception("Couldn't load config");
        }

        Instance.ErrorRangeInverse = 1 + 1 - Instance.ErrorRange;
        Instance.Blacklist = JsonSerializer.Deserialize<List<int>>(File.ReadAllText(BLACKLIST), options)!;
    }
}
