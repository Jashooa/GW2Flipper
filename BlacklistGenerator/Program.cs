namespace BlacklistGenerator;

using System.Text.Json;
using System.Text.Json.Serialization;

internal static class Program
{
    private static async Task Main()
    {
        var httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromMinutes(10)
        };

        var items = new List<ApiItem>();
        if (!File.Exists("items.json"))
        {
            var allIdsJson = await httpClient.GetStringAsync("https://api.guildwars2.com/v2/items");
            var allIds = JsonSerializer.Deserialize<List<int>>(allIdsJson);
            if (allIds == null)
            {
                return;
            }

            var itemJson = string.Empty;
            const int maxSize = 200;
            for (var i = 0; i < allIds.Count; i += maxSize)
            {
                if (i > allIds.Count)
                {
                    break;
                }

                var max = maxSize;
                if (i + maxSize > allIds.Count)
                {
                    max = allIds.Count % maxSize;
                }

                var currentIds = allIds.GetRange(i, max);
                Console.WriteLine($"Retrieving {currentIds.Min()} to {currentIds.Max()}");
                var currentIdsString = string.Join(',', currentIds);

                var newJson = await httpClient.GetStringAsync($"https://api.guildwars2.com/v2/items?ids={currentIdsString}");

                if (string.IsNullOrEmpty(itemJson))
                {
                    itemJson = newJson;
                }
                else
                {
                    itemJson = itemJson.Remove(itemJson.Length - 2);
                    itemJson += ',';
                    itemJson += newJson.Remove(0, 1);
                }
            }

            // itemJson = await httpClient.GetStringAsync("https://killproof.me/static/json/list_items.txt");

            await File.WriteAllTextAsync("items.json", itemJson);

            var newItems = JsonSerializer.Deserialize<List<ApiItem>>(itemJson);
            if (newItems == null)
            {
                return;
            }

            items.AddRange(newItems);
        }
        else
        {
            var itemJson = await File.ReadAllTextAsync("items.json");
            items.AddRange(JsonSerializer.Deserialize<List<ApiItem>>(itemJson)!);
        }

        var tradeJson = string.Empty;
        if (!File.Exists("trade_ids.json"))
        {
            tradeJson = await httpClient.GetStringAsync("https://api.guildwars2.com/v2/commerce/prices");
            await File.WriteAllTextAsync("trade_ids.json", tradeJson);
        }
        else
        {
            tradeJson = await File.ReadAllTextAsync("trade_ids.json");
        }

        var tradeIds = JsonSerializer.Deserialize<List<int>>(tradeJson);
        if (tradeIds == null)
        {
            return;
        }

        Console.WriteLine($"Total: {items.Count}");
        var tradeable = items.Where(x => !x.Flags.Contains("AccountBound") && tradeIds.Contains(x.Id)).ToList();
        Console.WriteLine($"Tradeable: {tradeable.Count}");
        var duplicates = tradeable.GroupBy(x => new { x.Name, x.Level, x.Rarity }).Where(x => x.Count() > 1).SelectMany(x => x).ToList();
        Console.WriteLine($"Duplicates: {duplicates.Count}");

        using StreamWriter file = new("duplicates.txt");
        foreach (var duplicate in duplicates)
        {
            //Console.WriteLine(duplicate);
            await file.WriteLineAsync(duplicate.ToString());
        }

        var idList = duplicates.ConvertAll(x => x.Id);
        var idString = JsonSerializer.Serialize(idList);
        await File.WriteAllTextAsync("blacklist.json", idString);
    }
}

internal class ApiItem
{
    [JsonPropertyName("id")]
    public int Id { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = null!;

    [JsonPropertyName("rarity")]
    public string Rarity { get; set; } = null!;

    [JsonPropertyName("level")]
    public int Level { get; set; }

    [JsonPropertyName("flags")]
    public string[] Flags { get; set; } = null!;

    [JsonPropertyName("game_types")]
    public string[] GameTypes { get; set; } = null!;

    public override string ToString()
    {
        var flagstring = "";
        foreach (var flag in Flags)
        {
            flagstring += " " + flag;
        }

        var gametypestring = "";
        foreach (var type in GameTypes)
        {
            gametypestring += " " + type;
        }

        return $"[{Id}] {Name}, Level: {Level}, Rarity: {Rarity}, Flags: [{flagstring} ], Game Types: [{gametypestring} ]";
    }
}
