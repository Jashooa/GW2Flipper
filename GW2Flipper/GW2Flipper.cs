namespace GW2Flipper;

using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;

using F23.StringSimilarity;

using global::GW2Flipper.Extensions;
using global::GW2Flipper.Models;
using global::GW2Flipper.Native;
using global::GW2Flipper.Properties;
using global::GW2Flipper.Utility;

using Gw2Sharp;

using HtmlAgilityPack;

using NLog;

internal static class GW2Flipper
{
    private const string ApiKey = "2A39D94B-053F-D84C-8D1C-2A9DC4E2C0917D50A41E-04D9-48B9-8133-FE968904C987";

    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

    private static readonly Dictionary<string, string> LowCostArguments = new()
    {
        { "sort", "profit" },
        { "ipg", "200" },
        { "sell-min", "100" },
        // { "sell-max", "0" },
        { "buy-min", "1" },
        { "buy-max", "500" },
        { "profit-min", "100" },
        // { "profit-max", "0" },
        // { "profit-pct-min", "0" },
        // { "profit-pct-max", "30" },
        // { "supply-min", "0" },
        // { "supply-max", "0" },
        { "demand-min", "100" },
        // { "demand-max", "0" },
        { "sold-day-min", "100" },
        // { "sold-day-max", "0" },
        { "bought-day-min", "50" },
        // { "bought-day-max", "0" },
        // { "offers-day-min", "0" },
        // { "offers-day-max", "0" },
        // { "bids-day-min", "0" },
        // { "bids-day-max", "0" },
    };

    private static readonly Dictionary<string, string> VeryFastArguments = new()
    {
        { "sort", "profit" },
        { "ipg", "200" },
        { "sell-min", "1" },
        // { "sell-max", "0" },
        { "buy-min", "1" },
        // { "buy-max", "500" },
        { "profit-min", "10" },
        // { "profit-max", "0" },
        { "profit-pct-min", "10" },
        // { "profit-pct-max", "30" },
        { "supply-min", "5000" },
        // { "supply-max", "0" },
        { "demand-min", "5000" },
        // { "demand-max", "0" },
        { "sold-day-min", "5000" },
        // { "sold-day-max", "0" },
        { "bought-day-min", "5000" },
        // { "bought-day-max", "0" },
        // { "offers-day-min", "0" },
        // { "offers-day-max", "0" },
        // { "bids-day-min", "0" },
        // { "bids-day-max", "0" },
    };

    private static readonly Dictionary<string, string> MediumArguments = new()
    {
        { "sort", "profit" },
        { "ipg", "200" },
        { "sell-min", "100" },
        // { "sell-max", "0" },
        { "buy-min", "1" },
        { "buy-max", "50000" },
        { "profit-min", "500" },
        // { "profit-max", "0" },
        { "profit-pct-min", "10" },
        { "profit-pct-max", "50" },
        // { "supply-min", "0" },
        // { "supply-max", "0" },
        // { "demand-min", "0" },
        // { "demand-max", "0" },
        { "sold-day-min", "200" },
        // { "sold-day-max", "0" },
        { "bought-day-min", "50" },
        // { "bought-day-max", "0" },
        // { "offers-day-min", "0" },
        // { "offers-day-max", "0" },
        // { "bids-day-min", "0" },
        // { "bids-day-max", "0" },
    };

    private static readonly Dictionary<string, Color> RarityColors = new()
    {
        { "Basic", Color.FromArgb(255, 255, 255) },
        { "Fine", Color.FromArgb(79, 157, 254) },
        { "Masterwork", Color.FromArgb(45, 197, 14) },
        { "Rare", Color.FromArgb(250, 223, 31) },
        { "Exotic", Color.FromArgb(238, 155, 3) },
        { "Ascended", Color.FromArgb(233, 67, 125) },
        { "Legendary", Color.FromArgb(160, 46, 247) },
        { "Gold", Color.FromArgb(253, 200, 78) },
        { "Silver", Color.FromArgb(180, 180, 180) },
        { "Copper", Color.FromArgb(202, 121, 66) },
    };

    private static readonly Dictionary<string, int> RarityOrder = new()
    {
        { "Junk", 0 },
        { "Basic", 1 },
        { "Fine", 2 },
        { "Masterwork", 3 },
        { "Rare", 4 },
        { "Exotic", 5 },
        { "Ascended", 6 },
        { "Legendary", 7 },
    };

    private enum TradingPostScreen
    {
        None,
        Buy,
        Sell,
        Transactions,
    }

    private static readonly Dictionary<string, string> WhichArguments = MediumArguments;
    private static readonly bool DoUndercut = true;
    private static readonly double Quantity = 0.01;
    private static readonly int MaxSpend = 20000;

    private static readonly List<BuyItem> BuyItemsList = new();

    private static readonly Connection ApiConnection = new(ApiKey);

    private static readonly Random Random = new();

    private static Process? process;

    private static Point? tradingPostPoint;

    private static DateTime timeSinceLastAfkCheck = DateTime.MinValue;

    private static DateTime timeSinceLastBuyListGenerate = DateTime.MinValue;

    public static async Task Run()
    {
        process = Process.GetProcessesByName("Gw2-64").SingleOrDefault();
        if (process == null)
        {
            Logger.Error("Couldn't find process");
            return;
        }

        Input.EnsureForegroundWindow(process);

        Logger.Info($"[{process!.Id}] {process!.MainWindowTitle} {process!.MainWindowHandle:x16}");

        tradingPostPoint = await GetTradingPostPoint();
        if (tradingPostPoint == null)
        {
            Logger.Error("Couldn't open trading post");
            return;
        }

        Logger.Info($"Found trading post at: {tradingPostPoint}");

        while (true)
        {
            await GenerateBuyList();

            try
            {
                await SellItems();
            }
            catch (Exception e)
            {
                Logger.Error(e);
            }

            Input.KeyPress(process!, VirtualKeyCode.ESCAPE);

            try
            {
                await CancelUndercut();
            }
            catch (Exception e)
            {
                Logger.Error(e);
            }

            Input.KeyPress(process!, VirtualKeyCode.ESCAPE);

            try
            {
                await BuyItems();
            }
            catch (Exception e)
            {
                Logger.Error(e);
            }

            Input.KeyPress(process!, VirtualKeyCode.ESCAPE);

            AntiAfk();
            CheckNewMap();

            var delay = Random.Next((int)TimeSpan.FromMinutes(3).TotalMilliseconds, (int)TimeSpan.FromMinutes(5).TotalMilliseconds);

            Logger.Info($"Sleeping for {TimeSpan.FromMilliseconds(delay).TotalMinutes} minutes");
            await Task.Delay(delay);
        }
    }

    private static async Task GenerateBuyList()
    {
        if (DateTime.Now - timeSinceLastBuyListGenerate < TimeSpan.FromMinutes(60))
        {
            return;
        }

        Logger.Info("Updating buy list");

        timeSinceLastBuyListGenerate = DateTime.Now;

        for (var page = 1; page < 10; page++)
        {
            var url = "https://www.gw2bltc.com/en/tp/search?";
            foreach (var argument in WhichArguments)
            {
                url += $"&{argument.Key}={argument.Value}";
            }

            url += $"&page={page}";

            HtmlWeb web = new();

            var htmlDoc = await web.LoadFromWebAsync(url);

            var itemTable = htmlDoc.DocumentNode.SelectSingleNode("//table[contains(@class, 'table-result')]/body");
            if (itemTable == null)
            {
                break;
            }

            foreach (var item in itemTable.SelectNodes(".//tr[td]"))
            {
                // var itemLink = item.SelectSingleNode("td[contains(@class, 'td-name')]/a").Attributes["href"].Value;
                // var itemId = Int32.Parse(Regex.Match(itemLink, @"/en/item/([\d]+)-[\w\d-]*").Groups[1].Value);
                var itemId = int.Parse(item.SelectSingleNode("td[1]/img").Attributes["data-id"].Value);
                var itemName = HtmlEntity.DeEntitize(item.SelectSingleNode("td[2]/a").InnerText);
                var sellPrice = int.Parse(item.SelectSingleNode("td[3]").InnerText);
                var buyPrice = int.Parse(item.SelectSingleNode("td[4]").InnerText);
                var profit = int.Parse(item.SelectSingleNode("td[5]").InnerText);
                var sold = int.Parse(item.SelectSingleNode("td[9]").InnerText, NumberStyles.AllowThousands);
                var bought = int.Parse(item.SelectSingleNode("td[11]").InnerText, NumberStyles.AllowThousands);

                var newItem = new BuyItem(itemId, itemName, buyPrice, sellPrice, profit, sold, bought);

                // Add or update here
                var findIndex = BuyItemsList.FindIndex(x => x.Id == itemId);
                if (findIndex != -1)
                {
                    BuyItemsList[findIndex] = newItem;
                }
                else
                {
                    BuyItemsList.Add(newItem);
                }
            }
        }

        BuyItemsList.Sort((x, y) => y.Profit.CompareTo(x.Profit));
    }

    private static async Task<Point?> GetTradingPostPoint()
    {
        var point = ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostLion);

        if (point == null)
        {
            Input.KeyPress(process!, VirtualKeyCode.VK_F);

            try
            {
                await WaitWhile(() => ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostLion) == null, 100, 10000);
            }
            catch (TimeoutException)
            {
                Logger.Error("Trading Post Lion not found");
                LogImage();
                return null;
            }

            point = ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostLion);
            Input.KeyPress(process!, VirtualKeyCode.ESCAPE);
        }

        return point;
    }

    private static async Task OpenTradingPost()
    {
        var point = ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostLion);

        if (point == null)
        {
            Input.KeyPress(process!, VirtualKeyCode.VK_F);

            try
            {
                await WaitWhile(() => ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostHome, 0.9) == null, 100, 10000);
            }
            catch (TimeoutException)
            {
                Logger.Error("Trading Post Home not found");
                LogImage();
            }

            point = ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostLion);
        }
    }

    private static async Task SellItems()
    {
        Logger.Info("====================");
        Logger.Info("Selling items");
        Logger.Info("====================");
        await OpenTradingPost();

        // Click take all
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 179, tradingPostPoint!.Value.Y + 685);
        await Task.Delay(100);

        // Click sell items
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 732, tradingPostPoint!.Value.Y + 90);

        // Wait for it to be activated
        try
        {
            static bool FindSellItemsActive()
            {
                var find = ImageSearch.FindImageInWindow(process!, Resources.SellItemsActive, tradingPostPoint!.Value.X + 805, tradingPostPoint!.Value.Y + 130, Resources.SellItemsActive.Width, Resources.SellItemsActive.Height, 0.99);
                return find == null;
            }

            await WaitWhile(FindSellItemsActive, 100, 10000);
        }
        catch (TimeoutException)
        {
            Logger.Error("Couldn't click sell items");
            LogImage();
            return;
        }

        // No items found
        if (ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.9) != null)
        {
            return;
        }

        using var apiClient = new Gw2Client(ApiConnection);

        apiClient.Mumble.Update();
        var character = apiClient.Mumble.CharacterName;

        var backpack = await apiClient.WebApi.V2.Characters[character].Inventory.GetAsync();
        var inventoryItems = backpack.Bags.SelectMany(x => x.Inventory.Select(y => y?.Id));

        var currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(0);
        var numSellingPages = currentSelling.HttpResponseInfo?.PageTotal;
        var currentSellingIds = currentSelling.Select(x => x.ItemId);
        for (var i = 1; i < numSellingPages; i++)
        {
            currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(i);
            currentSellingIds = currentSellingIds.Concat(currentSelling.Select(x => x.ItemId));
        }

        foreach (var item in BuyItemsList)
        {
            if (!inventoryItems.Contains(item.Id))
            {
                continue;
            }

            // Get item info
            var itemInfo = await apiClient.WebApi.V2.Items.GetAsync(item.Id);
            if (itemInfo == null)
            {
                Logger.Warn("itemInfo null");
                continue;
            }

            Logger.Info($"Selling [{item.Id}] {item.Name}");
            Logger.Info($"Buy: {item.BuyPrice} Sell: {item.SellPrice} Profit: {item.Profit} Bought: {item.Bought} Sold: {item.Sold}");

            // Search
            try
            {
                await SearchForItem(itemInfo, TradingPostScreen.Sell);
            }
            catch (TimeoutException)
            {
                Logger.Warn("Filter didn't appear");
                continue;
            }

            // Find position of Qty to determine position of search results
            var qtyPoint = ImageSearch.FindImageInWindow(process!, Resources.Qty, tradingPostPoint.Value.X + 324, tradingPostPoint.Value.Y + 155, Resources.Qty.Width, Resources.Qty.Height + 72, 0.9);
            if (qtyPoint == null)
            {
                Logger.Warn("qtyPoint null");
                LogImage();
                continue;
            }

            // Check if no item found
            if (ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.9) != null)
            {
                continue;
            }

            var resultsPoint = Point.Add(qtyPoint.Value, new Size(-2, 22));

            for (var y = 0; y < 7; y++)
            {
                try
                {
                    // Find the item border of the result
                    var resultPoint = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, resultsPoint.X, resultsPoint.Y + (y * 61), Resources.ResultCorner.Width, Resources.ResultCorner.Height);
                    if (resultPoint == null)
                    {
                        Logger.Debug($"No. {y} result not found at {resultsPoint.X}, {resultsPoint.Y + (y * 61)}");
                        break;
                    }

                    // Click the result
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 28, resultsPoint.Y + 27 + (y * 61));

                    // Wait for the item border to show up
                    try
                    {
                        static bool FindResultBorder()
                        {
                            var find = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, tradingPostPoint!.Value.X + 271, tradingPostPoint!.Value.Y + 136, Resources.ResultCorner.Width, Resources.ResultCorner.Height);
                            return find == null;
                        }

                        await WaitWhile(FindResultBorder, 100, 10000);
                    }
                    catch (TimeoutException)
                    {
                        Logger.Warn("Item border didn't appear");
                        LogImage();
                        continue;
                    }

                    await Task.Delay(500);

                    // Wait for gold coin to show up
                    try
                    {
                        static bool FindGoldCoin()
                        {
                            var find = ImageSearch.FindImageInWindow(process!, Resources.GoldCoin, tradingPostPoint!.Value.X + 462, tradingPostPoint!.Value.Y + 280, Resources.GoldCoin.Width, Resources.GoldCoin.Height + 24, 0.9);
                            return find == null;
                        }

                        await WaitWhile(FindGoldCoin, 100, 10000);
                    }
                    catch (TimeoutException)
                    {
                        Logger.Warn("Gold coin didn't appear");
                        LogImage();
                        break;
                    }

                    // Wait for available to show up
                    try
                    {
                        static bool FindAvailable()
                        {
                            var find1 = ImageSearch.FindImageInWindow(process!, Resources.Available, tradingPostPoint!.Value.X + 531, tradingPostPoint!.Value.Y + 469, Resources.Available.Width, Resources.Available.Height);
                            var find2 = ImageSearch.FindImageInWindow(process!, Resources.Available, tradingPostPoint!.Value.X + 531, tradingPostPoint!.Value.Y + 493, Resources.Available.Width, Resources.Available.Height, 0.9);
                            return find1 == null && find2 == null;
                        }

                        await WaitWhile(FindAvailable, 100, 10000);
                    }
                    catch (TimeoutException)
                    {
                        Logger.Warn("Available didn't appear");
                        LogImage();
                        break;
                    }

                    await Task.Delay(500);

                    // Find position of min. price gold coin image to set the offset for other UI positions
                    var goldCoinPos = ImageSearch.FindImageInWindow(process!, Resources.GoldCoin, tradingPostPoint!.Value.X + 462, tradingPostPoint!.Value.Y + 280, Resources.GoldCoin.Width, Resources.GoldCoin.Height + 24, 0.9);
                    if (goldCoinPos == null)
                    {
                        Logger.Warn("goldCoinPos null");
                        LogImage();
                        continue;
                    }

                    var offset = 0;
                    var nameSize = 34;
                    if (goldCoinPos.Value.Y == tradingPostPoint!.Value.Y + 304)
                    {
                        offset = 24;
                        nameSize = 64;
                    }

                    // Get OCR of item name and check if it matches API name
                    var nameImage = ImageSearch.CaptureWindow(process!, tradingPostPoint!.Value.X + 334, tradingPostPoint.Value.Y + 136, 428, nameSize);
                    if (nameImage == null)
                    {
                        Logger.Warn("nameImage null");
                        LogImage();
                        continue;
                    }

                    var capturedName = OCR.ReadName(nameImage, RarityColors[itemInfo.Rarity.ToString()!]);
                    if (string.IsNullOrEmpty(capturedName))
                    {
                        Logger.Debug("Empty string returned");
                        LogImage();
                        continue;
                    }

                    var jw = new JaroWinkler();
                    var nameDistance = 1.0 - jw.Distance(itemInfo.Name.RemoveDiacritics(), capturedName);
                    Logger.Debug($"Item name: {itemInfo.Name} Captured name: {capturedName} Distance: {nameDistance}");
                    if (nameDistance < 0.98)
                    {
                        await CloseItemWindow();
                        continue;
                    }

                    // Click highest current seller
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint.Value.X + 542, tradingPostPoint.Value.Y + 499 + offset);
                    await Task.Delay(1000);

                    // Set quantity to 1
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 448, tradingPostPoint.Value.Y + 240 + offset);
                    await Task.Delay(1000);

                    // Get sell price
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 402, tradingPostPoint.Value.Y + 236 + offset);
                    await Task.Delay(100);
                    Input.KeyPress(process!, VirtualKeyCode.TAB);
                    await Task.Delay(100);
                    var goldAmount = Convert.ToInt32(Input.GetSelectedText(process!));
                    await Task.Delay(100);
                    Input.KeyPress(process!, VirtualKeyCode.TAB);
                    await Task.Delay(100);
                    var silverAmount = Convert.ToInt32(Input.GetSelectedText(process!));
                    await Task.Delay(100);
                    Input.KeyPress(process!, VirtualKeyCode.TAB);
                    await Task.Delay(100);
                    var copperAmount = Convert.ToInt32(Input.GetSelectedText(process!));

                    var sellCurrencyAmount = ToCurrency(goldAmount, silverAmount, copperAmount);
                    Logger.Info($"Highest seller: {sellCurrencyAmount}");

                    var sellPrice = sellCurrencyAmount;

                    if (DoUndercut)
                    {
                        sellPrice--;
                    }

                    if (DoUndercut && !(currentSellingIds.Contains(item.Id) && currentSelling.First(x => x.ItemId == item.Id).Price == sellCurrencyAmount))
                    {
                        // Click to -1 copper
                        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint.Value.X + 593, tradingPostPoint.Value.Y + 283 + offset);
                        await Task.Delay(1000);
                        sellCurrencyAmount--;
                    }

                    // Check if sell price is within range
                    if (sellCurrencyAmount < item.SellPrice * 0.9)
                    {
                        Logger.Debug($"Sell price {sellCurrencyAmount} below range {item.SellPrice * 0.9}");
                        await CloseItemWindow();
                        continue;
                    }

                    // Set quantity to max
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 625, tradingPostPoint.Value.Y + 240 + offset);
                    await Task.Delay(1000);

                    Logger.Info($"Selling for {sellPrice} coins.");

                    // Sell item
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 440, tradingPostPoint.Value.Y + 373 + offset);
                    await Task.Delay(1000);
                    await CloseItemWindow();
                    break;
                }
                catch (FormatException e)
                {
                    Logger.Error(e);
                    LogImage();
                    await CloseItemWindow();
                }
            }
        }

        // Click take all
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 179, tradingPostPoint!.Value.Y + 685);
        await Task.Delay(100);
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 179, tradingPostPoint!.Value.Y + 685);
        await Task.Delay(100);
    }

    private static async Task BuyItems()
    {
        Logger.Info("====================");
        Logger.Info("Buying items");
        Logger.Info("====================");
        Logger.Info($"Items: {BuyItemsList.Count}");
        await OpenTradingPost();

        // Click buy items
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 566, tradingPostPoint!.Value.Y + 90);

        // Wait for it to be activated
        try
        {
            static bool FindBuyItemsActive()
            {
                var find = ImageSearch.FindImageInWindow(process!, Resources.BuyItemsActive, tradingPostPoint!.Value.X + 638, tradingPostPoint!.Value.Y + 130, Resources.BuyItemsActive.Width, Resources.BuyItemsActive.Height, 0.99);
                return find == null;
            }

            await WaitWhile(FindBuyItemsActive, 100, 10000);
        }
        catch (TimeoutException)
        {
            Logger.Error("Couldn't click buy items.");
            LogImage();
            return;
        }

        // Click take all
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 179, tradingPostPoint!.Value.Y + 685);
        await Task.Delay(100);

        using var apiClient = new Gw2Client(ApiConnection);

        apiClient.Mumble.Update();
        var character = apiClient.Mumble.CharacterName;

        var currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(0);
        var numSellingPages = currentSelling.HttpResponseInfo?.PageTotal;
        var currentSellingIds = currentSelling.Select(x => x.ItemId);
        for (var i = 1; i < numSellingPages; i++)
        {
            currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(i);
            currentSellingIds = currentSellingIds.Concat(currentSelling.Select(x => x.ItemId));
        }

        var currentBuying = await apiClient.WebApi.V2.Commerce.Transactions.Current.Buys.PageAsync(0);
        var numBuyingPages = currentBuying.HttpResponseInfo?.PageTotal;
        var currentBuyingIds = currentBuying.Select(x => x.ItemId);
        for (var i = 1; i < numBuyingPages; i++)
        {
            currentBuying = await apiClient.WebApi.V2.Commerce.Transactions.Current.Buys.PageAsync(i);
            currentBuyingIds = currentBuyingIds.Concat(currentBuying.Select(x => x.ItemId));
        }

        foreach (var item in BuyItemsList)
        {
            // Check if currently selling item
            if (currentSellingIds.Contains(item.Id))
            {
                continue;
            }

            // Check if already buying item
            if (currentBuyingIds.Contains(item.Id))
            {
                continue;
            }

            Logger.Info($"Buying [{item.Id}] {item.Name}");
            Logger.Info($"Buy: {item.BuyPrice} Sell: {item.SellPrice} Profit: {item.Profit} Bought: {item.Bought} Sold: {item.Sold}");

            // Get item info
            var itemInfo = await apiClient.WebApi.V2.Items.GetAsync(item.Id);
            if (itemInfo == null)
            {
                Logger.Warn("itemInfo null");
                continue;
            }

            // Search
            try
            {
                await SearchForItem(itemInfo, TradingPostScreen.Buy);
            }
            catch (TimeoutException)
            {
                Logger.Warn("Filter didn't appear");
                continue;
            }

            // Find position of Qty to determine position of search results
            var qtyPoint = ImageSearch.FindImageInWindow(process!, Resources.Qty, tradingPostPoint.Value.X + 324, tradingPostPoint.Value.Y + 155, Resources.Qty.Width, Resources.Qty.Height + 72, 0.9);
            if (qtyPoint == null)
            {
                Logger.Warn("qtyPoint null");
                LogImage();
                continue;
            }

            var resultsPoint = Point.Add(qtyPoint.Value, new Size(-2, 22));

            // Wait for search results to load
            try
            {
                static bool FindItemsLoaded()
                {
                    var noItems = ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.9);
                    var resultBorder = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, tradingPostPoint!.Value.X + 322, tradingPostPoint!.Value.Y + 177, Resources.ResultCorner.Width, Resources.ResultCorner.Height + 72);
                    return noItems == null && resultBorder == null;
                }

                await WaitWhile(() => FindItemsLoaded(), 100, 10000);
            }
            catch (TimeoutException)
            {
                Logger.Error("Didn't load any search results.");
                LogImage();
                continue;
            }

            // Check if no item found
            if (ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.9) != null)
            {
                continue;
            }

            await Task.Delay(2000);

            for (var y = 0; y < 7; y++)
            {
                try
                {
                    // Find the item border of the result
                    var resultPoint = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, resultsPoint.X, resultsPoint.Y + (y * 61), Resources.ResultCorner.Width, Resources.ResultCorner.Height);
                    if (resultPoint == null)
                    {
                        Logger.Debug($"No. {y} result not found at {resultsPoint.X}, {resultsPoint.Y + (y * 61)}");
                        break;
                    }

                    // Click the result
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 28, resultsPoint.Y + 27 + (y * 61));

                    // Wait for the item border to show up
                    try
                    {
                        static bool FindResultBorder()
                        {
                            var find = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, tradingPostPoint!.Value.X + 271, tradingPostPoint!.Value.Y + 136, Resources.ResultCorner.Width, Resources.ResultCorner.Height);
                            return find == null;
                        }

                        await WaitWhile(FindResultBorder, 100, 10000);
                    }
                    catch (TimeoutException)
                    {
                        Logger.Warn("Item border didn't appear");
                        LogImage();
                        break;
                    }

                    await Task.Delay(500);

                    // Wait for gold coin to show up
                    try
                    {
                        static bool FindGoldCoin()
                        {
                            var find = ImageSearch.FindImageInWindow(process!, Resources.GoldCoin, tradingPostPoint!.Value.X + 462, tradingPostPoint!.Value.Y + 280, Resources.GoldCoin.Width, Resources.GoldCoin.Height + 24, 0.9);
                            return find == null;
                        }

                        await WaitWhile(FindGoldCoin, 100, 10000);
                    }
                    catch (TimeoutException)
                    {
                        Logger.Warn("Gold coin didn't appear");
                        LogImage();
                        break;
                    }

                    // Wait for available to show up
                    try
                    {
                        static bool FindAvailable()
                        {
                            var find1 = ImageSearch.FindImageInWindow(process!, Resources.Available, tradingPostPoint!.Value.X + 531, tradingPostPoint!.Value.Y + 469, Resources.Available.Width, Resources.Available.Height);
                            var find2 = ImageSearch.FindImageInWindow(process!, Resources.Available, tradingPostPoint!.Value.X + 531, tradingPostPoint!.Value.Y + 493, Resources.Available.Width, Resources.Available.Height, 0.9);
                            return find1 == null && find2 == null;
                        }

                        await WaitWhile(FindAvailable, 100, 10000);
                    }
                    catch (TimeoutException)
                    {
                        Logger.Warn("Available didn't appear");
                        LogImage();
                        break;
                    }

                    await Task.Delay(500);

                    // Find position of min. price gold coin image to set the offset for other UI positions
                    var goldCoinPos = ImageSearch.FindImageInWindow(process!, Resources.GoldCoin, tradingPostPoint!.Value.X + 462, tradingPostPoint!.Value.Y + 280, Resources.GoldCoin.Width, Resources.GoldCoin.Height + 24, 0.9);
                    if (goldCoinPos == null)
                    {
                        Logger.Warn("goldCoinPos null");
                        LogImage();
                        break;
                    }

                    var offset = 0;
                    var nameSize = 34;
                    if (goldCoinPos.Value.Y == tradingPostPoint!.Value.Y + 304)
                    {
                        offset = 24;
                        nameSize = 64;
                    }

                    // Get OCR of item name and check if it matches API name
                    var nameImage = ImageSearch.CaptureWindow(process!, tradingPostPoint!.Value.X + 334, tradingPostPoint.Value.Y + 136, 428, nameSize);
                    if (nameImage == null)
                    {
                        Logger.Warn("nameImage null");
                        LogImage();
                        continue;
                    }

                    var capturedName = OCR.ReadName(nameImage, RarityColors[itemInfo.Rarity.ToString()!]);
                    if (string.IsNullOrEmpty(capturedName))
                    {
                        Logger.Debug("Empty string returned");
                        LogImage();
                        continue;
                    }

                    var jw = new JaroWinkler();
                    var nameDistance = 1.0 - jw.Distance(itemInfo.Name.RemoveDiacritics(), capturedName);
                    Logger.Debug($"Item name: {itemInfo.Name} Captured name: {capturedName} Distance: {nameDistance}");
                    if (nameDistance < 0.98)
                    {
                        await CloseItemWindow();
                        continue;
                    }

                    await Task.Delay(1000);

                    // Get sell price
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 402, tradingPostPoint.Value.Y + 236 + offset);
                    await Task.Delay(100);
                    Input.KeyPress(process!, VirtualKeyCode.TAB);
                    await Task.Delay(100);
                    var goldAmount = Convert.ToInt32(Input.GetSelectedText(process!));
                    await Task.Delay(100);
                    Input.KeyPress(process!, VirtualKeyCode.TAB);
                    await Task.Delay(100);
                    var silverAmount = Convert.ToInt32(Input.GetSelectedText(process!));
                    await Task.Delay(100);
                    Input.KeyPress(process!, VirtualKeyCode.TAB);
                    await Task.Delay(100);
                    var copperAmount = Convert.ToInt32(Input.GetSelectedText(process!));

                    var sellCurrencyAmount = ToCurrency(goldAmount, silverAmount, copperAmount);
                    Logger.Info($"Highest seller: {sellCurrencyAmount}");

                    // Click highest current buyer
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint.Value.X + 254, tradingPostPoint.Value.Y + 498 + offset);
                    await Task.Delay(1000);

                    // Set quantity to 1
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 448, tradingPostPoint.Value.Y + 240 + offset);
                    await Task.Delay(1000);

                    // Get buy price
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 402, tradingPostPoint.Value.Y + 236 + offset);
                    await Task.Delay(100);
                    Input.KeyPress(process!, VirtualKeyCode.TAB);
                    await Task.Delay(100);
                    goldAmount = Convert.ToInt32(Input.GetSelectedText(process!));
                    await Task.Delay(100);
                    Input.KeyPress(process!, VirtualKeyCode.TAB);
                    await Task.Delay(100);
                    silverAmount = Convert.ToInt32(Input.GetSelectedText(process!));
                    await Task.Delay(100);
                    Input.KeyPress(process!, VirtualKeyCode.TAB);
                    await Task.Delay(100);
                    copperAmount = Convert.ToInt32(Input.GetSelectedText(process!));

                    var buyCurrencyAmount = ToCurrency(goldAmount, silverAmount, copperAmount);
                    Logger.Info($"Highest buyer: {buyCurrencyAmount}");

                    // Check if prices are wrong
                    if ((sellCurrencyAmount > item.SellPrice * 1.1 || sellCurrencyAmount < item.SellPrice * 0.9)
                        && (buyCurrencyAmount > item.BuyPrice * 0.9 || buyCurrencyAmount < item.BuyPrice * 1.1))
                    {
                        Logger.Debug("Sell & buy price too different.");
                        await CloseItemWindow();
                        continue;
                    }

                    // Check if sell price is within range
                    if (sellCurrencyAmount < item.SellPrice * 0.9)
                    {
                        Logger.Debug("Sell price below range.");
                        await CloseItemWindow();
                        continue;
                    }

                    // Check if buy price is within range
                    if (buyCurrencyAmount > item.BuyPrice * 1.1)
                    {
                        Logger.Debug("Buy price above range.");
                        await CloseItemWindow();
                        continue;
                    }

                    var buyPrice = buyCurrencyAmount;
                    var sellprice = sellCurrencyAmount;

                    if (DoUndercut)
                    {
                        buyPrice++;
                        sellprice--;
                    }

                    // Check profit is within range
                    if (GetProfit(buyPrice, sellprice) < (item.Profit * 0.9))
                    {
                        Logger.Debug("Profit outside of range.");
                        break;
                    }

                    if (DoUndercut)
                    {
                        // Click to +1 copper
                        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint.Value.X + 593, tradingPostPoint.Value.Y + 272 + offset);
                        await Task.Delay(1000);
                    }

                    var quantity = Math.Min(Math.Min((int)(MaxSpend / buyPrice), (int)(item.Sold * Quantity)), 250);
                    quantity = quantity < 1 ? 1 : quantity;

                    Logger.Info($"Buying x{quantity} for {buyPrice} coins each.");

                    if (quantity > 1)
                    {
                        // Set quantity
                        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 402, tradingPostPoint.Value.Y + 236 + offset);
                        await Task.Delay(100);
                        Input.KeyStringSend(process!, quantity.ToString());
                        await Task.Delay(500);
                    }

                    // Buy item
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 440, tradingPostPoint.Value.Y + 373 + offset);
                    await Task.Delay(1000);
                    await CloseItemWindow();
                    break;
                }
                catch (FormatException e)
                {
                    Logger.Error(e);
                    LogImage();
                    await CloseItemWindow();
                }
            }

            // Click take all
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 179, tradingPostPoint!.Value.Y + 685);
            await Task.Delay(100);
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 179, tradingPostPoint!.Value.Y + 685);
            await Task.Delay(100);
        }
    }

    private static async Task CancelUndercut()
    {
        Logger.Info("====================");
        Logger.Info("Cancelling undercut buys");
        Logger.Info("====================");
        await OpenTradingPost();

        // Click transactions
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 899, tradingPostPoint!.Value.Y + 90);

        // Wait for it to be activated
        try
        {
            static bool FindTransactionsActive()
            {
                var find = ImageSearch.FindImageInWindow(process!, Resources.TransactionsActive, tradingPostPoint!.Value.X + 817, tradingPostPoint!.Value.Y + 130, Resources.TransactionsActive.Width, Resources.TransactionsActive.Height, 0.99);
                return find == null;
            }

            await WaitWhile(FindTransactionsActive, 100, 10000);
        }
        catch (TimeoutException)
        {
            Logger.Error("Couldn't click transactions.");
            LogImage();
            return;
        }

        // Wait for search results to load
        try
        {
            static bool FindItemsLoaded()
            {
                // var noItems = ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.9);
                var resultBorder = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, tradingPostPoint!.Value.X + 322, tradingPostPoint!.Value.Y + 177, Resources.ResultCorner.Width, Resources.ResultCorner.Height + 72);
                // return noItems == null && resultBorder == null;
                return resultBorder == null;
            }

            await WaitWhile(() => FindItemsLoaded(), 100, 10000);
        }
        catch (TimeoutException)
        {
            Logger.Error("Didn't load any search results.");
            LogImage();
            return;
        }

        // Check if no item found
        if (ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.9) != null)
        {
            return;
        }

        // Click buying
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 174, tradingPostPoint!.Value.Y + 240);

        // Wait for search results to load
        try
        {
            static bool FindItemsLoaded()
            {
                // var noItems = ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.9);
                var resultBorder = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, tradingPostPoint!.Value.X + 322, tradingPostPoint!.Value.Y + 177, Resources.ResultCorner.Width, Resources.ResultCorner.Height + 72);
                // return noItems == null && resultBorder == null;
                return resultBorder == null;
            }

            await WaitWhile(() => FindItemsLoaded(), 100, 10000);
        }
        catch (TimeoutException)
        {
            Logger.Error("Didn't load any search results.");
            LogImage();
            return;
        }

        // Check if no item found
        if (ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.9) != null)
        {
            return;
        }

        using var apiClient = new Gw2Client(ApiConnection);

        var currentBuying = await apiClient.WebApi.V2.Commerce.Transactions.Current.Buys.PageAsync(0);
        var numBuyingPages = currentBuying.HttpResponseInfo?.PageTotal;
        var currentBuyingIds = currentBuying.Select(x => x.ItemId);
        for (var i = 1; i < numBuyingPages; i++)
        {
            currentBuying = await apiClient.WebApi.V2.Commerce.Transactions.Current.Buys.PageAsync(i);
            currentBuyingIds = currentBuyingIds.Concat(currentBuying.Select(x => x.ItemId));
        }

        foreach (var item in BuyItemsList)
        {
            // Check if we're currently buying item
            if (currentBuyingIds.Contains(item.Id))
            {
                var itemPrices = await apiClient.WebApi.V2.Commerce.Prices.GetAsync(item.Id);
                var highestPrice = itemPrices.Buys.UnitPrice;
                var price = currentBuying.First(x => x.ItemId == item.Id).Price;

                // If we haven't been undercut
                if (highestPrice <= price)
                {
                    Logger.Debug($"{item.Name} Not undercut");
                    continue;
                }
            }
            else
            {
                Logger.Debug($"{item.Name} Not buying");
                continue;
            }

            Logger.Info($"Cancelling {item.Name}");

            // Get item info
            var itemInfo = await apiClient.WebApi.V2.Items.GetAsync(item.Id);
            if (itemInfo == null)
            {
                Logger.Warn("itemInfo null");
                continue;
            }

            // Search
            try
            {
                await SearchForItem(itemInfo, TradingPostScreen.Transactions);
            }
            catch (TimeoutException)
            {
                Logger.Warn("Filter didn't appear");
                continue;
            }

            // Find position of Qty to determine position of search results
            var qtyPoint = ImageSearch.FindImageInWindow(process!, Resources.Qty, tradingPostPoint.Value.X + 324, tradingPostPoint.Value.Y + 155, Resources.Qty.Width, Resources.Qty.Height + 72, 0.9);
            if (qtyPoint == null)
            {
                Logger.Warn("qtyPoint null");
                LogImage();
                continue;
            }

            var resultsPoint = Point.Add(qtyPoint.Value, new Size(-2, 22));

            // Wait for search results to load
            try
            {
                static bool FindItemsLoaded()
                {
                    var noItems = ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.9);
                    var resultBorder = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, tradingPostPoint!.Value.X + 322, tradingPostPoint!.Value.Y + 177, Resources.ResultCorner.Width, Resources.ResultCorner.Height + 72);
                    return noItems == null && resultBorder == null;
                }

                await WaitWhile(() => FindItemsLoaded(), 100, 10000);
            }
            catch (TimeoutException)
            {
                Logger.Error("Didn't load any search results.");
                LogImage();
                continue;
            }

            // Check if no item found
            if (ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.9) != null)
            {
                continue;
            }

            await Task.Delay(2000);

            for (var y = 0; y < 7; y++)
            {
                // Find the item border of the result
                var resultPoint = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, resultsPoint.X, resultsPoint.Y + (y * 61), Resources.ResultCorner.Width, Resources.ResultCorner.Height);
                if (resultPoint == null)
                {
                    Logger.Debug($"No. {y} result not found at {resultsPoint.X}, {resultsPoint.Y + (y * 61)}");
                    break;
                }

                // Click the result
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 28, resultsPoint.Y + 27 + (y * 61));

                // Wait for the item border to show up
                try
                {
                    static bool FindResultBorder()
                    {
                        var find = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, tradingPostPoint!.Value.X + 271, tradingPostPoint!.Value.Y + 136, Resources.ResultCorner.Width, Resources.ResultCorner.Height);
                        return find == null;
                    }

                    await WaitWhile(FindResultBorder, 100, 10000);
                }
                catch (TimeoutException)
                {
                    Logger.Warn("Item border didn't appear");
                    LogImage();
                    break;
                }

                await Task.Delay(500);

                // Wait for gold coin to show up
                try
                {
                    static bool FindGoldCoin()
                    {
                        var find = ImageSearch.FindImageInWindow(process!, Resources.GoldCoin, tradingPostPoint!.Value.X + 462, tradingPostPoint!.Value.Y + 280, Resources.GoldCoin.Width, Resources.GoldCoin.Height + 24, 0.9);
                        return find == null;
                    }

                    await WaitWhile(FindGoldCoin, 100, 10000);
                }
                catch (TimeoutException)
                {
                    Logger.Warn("Gold coin didn't appear");
                    LogImage();
                    break;
                }

                // Wait for available to show up
                /*try
                {
                    static bool FindAvailable()
                    {
                        var find1 = ImageSearch.FindImageInWindow(process!, Resources.Available, tradingPostPoint!.Value.X + 531, tradingPostPoint!.Value.Y + 469, Resources.Available.Width, Resources.Available.Height);
                        var find2 = ImageSearch.FindImageInWindow(process!, Resources.Available, tradingPostPoint!.Value.X + 531, tradingPostPoint!.Value.Y + 493, Resources.Available.Width, Resources.Available.Height, 0.9);
                        return find1 == null && find2 == null;
                    }

                    await WaitWhile(FindAvailable, 100, 10000);
                }
                catch (TimeoutException)
                {
                    Logger.Warn("Available didn't appear");
                    LogImage();
                    break;
                }

                await Task.Delay(500);*/

                // Find position of min. price gold coin image to set the offset for other UI positions
                var goldCoinPos = ImageSearch.FindImageInWindow(process!, Resources.GoldCoin, tradingPostPoint!.Value.X + 462, tradingPostPoint!.Value.Y + 280, Resources.GoldCoin.Width, Resources.GoldCoin.Height + 24, 0.9);
                if (goldCoinPos == null)
                {
                    Logger.Warn("goldCoinPos null");
                    LogImage();
                    break;
                }

                var nameSize = 34;
                if (goldCoinPos.Value.Y == tradingPostPoint!.Value.Y + 304)
                {
                    nameSize = 64;
                }

                // Get OCR of item name and check if it matches API name
                var nameImage = ImageSearch.CaptureWindow(process!, tradingPostPoint!.Value.X + 334, tradingPostPoint.Value.Y + 136, 428, nameSize);
                if (nameImage == null)
                {
                    Logger.Warn("nameImage null");
                    LogImage();
                    continue;
                }

                var capturedName = OCR.ReadName(nameImage, RarityColors[itemInfo.Rarity.ToString()!]);
                if (string.IsNullOrEmpty(capturedName))
                {
                    Logger.Debug("Empty string returned");
                    LogImage();
                    continue;
                }

                var jw = new JaroWinkler();
                var nameDistance = 1.0 - jw.Distance(itemInfo.Name.RemoveDiacritics(), capturedName);
                Logger.Debug($"Item name: {itemInfo.Name} Captured name: {capturedName} Distance: {nameDistance}");
                if (nameDistance < 0.98)
                {
                    await CloseItemWindow();
                    continue;
                }

                // Press X to close
                await CloseItemWindow();

                // Click cancel and confirm
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 600, resultsPoint.Y + 28 + (y * 61));
                await Task.Delay(500);
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 600, resultsPoint.Y + 28 + (y * 61));
                await Task.Delay(500);

                y--;
            }
        }
    }

    private static async Task CloseItemWindow()
    {
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 777, tradingPostPoint!.Value.Y + 125);
        await Task.Delay(1000);
    }

    private static async Task SearchForItem(Gw2Sharp.WebApi.V2.Models.Item itemInfo, TradingPostScreen screen)
    {
        // Click search bar
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 163, tradingPostPoint!.Value.Y + 165);
        await Task.Delay(500);

        // Check if filter isn't open
        if (ImageSearch.FindImageInWindow(process!, Resources.FilterOpen, tradingPostPoint!.Value.X + 50, tradingPostPoint!.Value.Y + 182, Resources.FilterOpen.Width, Resources.FilterOpen.Height, 0.9) == null)
        {
            // Click filter cog
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 287, tradingPostPoint!.Value.Y + 165);
            await Task.Delay(100);

            // Wait for filter to show up
            try
            {
                static bool FindFilter()
                {
                    var find = ImageSearch.FindImageInWindow(process!, Resources.FilterOpen, tradingPostPoint!.Value.X + 50, tradingPostPoint!.Value.Y + 182, Resources.FilterOpen.Width, Resources.FilterOpen.Height, 0.9);
                    return find == null;
                }

                await WaitWhile(FindFilter, 100, 10000);
            }
            catch (TimeoutException)
            {
                LogImage();
                throw new TimeoutException();
            }

            await Task.Delay(100);
        }

        // Click reset filters
        if (screen == TradingPostScreen.Buy)
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 173, tradingPostPoint!.Value.Y + 356);
        }
        else
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 173, tradingPostPoint!.Value.Y + 309);
        }

        await Task.Delay(500);

        // Input level
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 77, tradingPostPoint!.Value.Y + 270);
        await Task.Delay(100);
        Input.KeyStringSend(process!, itemInfo.Level.ToString());
        await Task.Delay(100);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 138, tradingPostPoint!.Value.Y + 270);
        await Task.Delay(100);
        Input.KeyStringSend(process!, itemInfo.Level.ToString());
        await Task.Delay(100);

        // Set rarity
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 173, tradingPostPoint!.Value.Y + 205);
        await Task.Delay(500);
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 173, tradingPostPoint!.Value.Y + 230 + (25 * RarityOrder[itemInfo.Rarity!]));
        await Task.Delay(500);

        // Click search bar
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 163, tradingPostPoint!.Value.Y + 165);
        await Task.Delay(100);

        // Select all
        // Input.KeyPressWithModifier(process!, VirtualKeyCode.VK_A, false, true, false);
        // await Task.Delay(500);

        // Backspace
        // Input.KeyPress(process!, VirtualKeyCode.BACK);
        // await Task.Delay(500);

        // Type in item name
        // Input.KeyStringSend(process!, itemInfo.Name.MaxSize(30));
        // Input.KeyPress(process!, VirtualKeyCode.RETURN);
        // await Task.Delay(500);

        // Paste in name
        Input.KeyStringSendClipboard(process!, itemInfo.Name.MaxSize(30));
        Input.KeyPress(process!, VirtualKeyCode.RETURN);
        await Task.Delay(500);

        // Click filter cog to close
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 287, tradingPostPoint!.Value.Y + 165);
        await Task.Delay(1000);
    }

    private static async void CheckNewMap()
    {
        Logger.Info("Checking if new map available");
        await Task.Delay(2000);
        var shroomPoint = ImageSearch.FindImageInFullWindow(process!, Resources.NewMapShroom);

        if (shroomPoint != null)
        {
            Logger.Info("Found shroom");
            var yesPoint = ImageSearch.FindImageInWindow(process!, Resources.NewMapYes, shroomPoint.Value.X - 118, shroomPoint.Value.Y + 42, Resources.NewMapYes.Width, Resources.NewMapYes.Height, 0.9);

            if (yesPoint != null)
            {
                Logger.Info("Accepting new map");

                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, yesPoint.Value.X + 26, yesPoint.Value.Y + 9);
                Thread.Sleep(10000);
            }
        }
    }

    private static void AntiAfk()
    {
        if (DateTime.Now - timeSinceLastAfkCheck < TimeSpan.FromMinutes(30))
        {
            return;
        }

        Logger.Info("Moving for anti-afk");
        Input.KeyPress(process!, VirtualKeyCode.VK_S);
        Input.KeyPress(process!, VirtualKeyCode.VK_W);
        timeSinceLastAfkCheck = DateTime.Now;
    }

    private static int ToCurrency(int gold, int silver, int copper) => (gold * 10000) + (silver * 100) + copper;

    private static int GetProfit(int buy, int sell)
    {
        var listingFee = (int)Math.Ceiling(sell * 0.05);
        listingFee = listingFee < 1 ? 1 : listingFee;

        var exchangeFee = (int)Math.Ceiling(sell * 0.1);
        exchangeFee = exchangeFee < 1 ? 1 : exchangeFee;

        return sell - (buy + listingFee + exchangeFee);
    }

    private static bool IsStackable(string type) => type is "Consumable"
    or "Container"
    or "CraftingMaterial"
    or "Trophy"
    or "UpgradeComponent";

    private static async Task WaitWhile(Func<bool> condition, int frequency, int timeout)
    {
        var waitTask = Task.Run(async () =>
        {
            while (condition())
            {
                await Task.Delay(frequency);
            }
        });

        if (waitTask != await Task.WhenAny(waitTask, Task.Delay(timeout)))
        {
            throw new TimeoutException();
        }
    }

    private static void LogImage()
    {
        var image = ImageSearch.CaptureFullWindow(process!);
        if (image != null)
        {
            var directory = Path.Combine("logs", "images");
            if (!Directory.Exists(directory))
            {
                _ = Directory.CreateDirectory(directory);
            }

            var fileName = Path.Combine(directory, $"{DateTime.Now.ToString("s").Replace(":", string.Empty)}.png");
            Logger.Error($"File saved as {fileName}");
            image.Save(fileName, ImageFormat.Png);
        }
    }

    private static void LogTPImage()
    {
        if (tradingPostPoint == null)
        {
            return;
        }

        var image = ImageSearch.CaptureWindow(process!, tradingPostPoint.Value.X, tradingPostPoint.Value.Y, 992, 732);
        if (image != null)
        {
            var directory = Path.Combine("TPImages");
            if (!Directory.Exists(directory))
            {
                _ = Directory.CreateDirectory(directory);
            }

            var fileName = Path.Combine(directory, $"{DateTime.Now.ToString("s").Replace(":", string.Empty)}.bmp");
            Logger.Error($"File saved as {fileName}");
            image.Save(fileName, ImageFormat.Bmp);
        }
    }
}
