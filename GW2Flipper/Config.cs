namespace GW2Flipper;

using System.Text.Json;

internal class Config
{
    private const string PATH = "./config.json";

    public static Config? Instance { get; private set; }

    public static string ApiKey { get; set; } = string.Empty;

    public static void Load()
    {
        if (!File.Exists(PATH))
        {
            throw new FileNotFoundException($"{PATH} is missing.");
        }

        Instance = JsonSerializer.Deserialize<Config>(File.ReadAllText(PATH));
    }
}
