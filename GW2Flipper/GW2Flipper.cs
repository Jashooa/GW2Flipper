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
        { "sell-min", "100" },
        // { "sell-max", "0" },
        { "buy-min", "1" },
        { "buy-max", "500" },
        { "profit-min", "100" },
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

    private static readonly Dictionary<string, string> VeryFastArguments = new()
    {
        { "sell-min", "1" },
        // { "sell-max", "0" },
        { "buy-min", "1" },
        // { "buy-max", "500" },
        { "profit-min", "20" },
        // { "profit-max", "0" },
        { "profit-pct-min", "10" },
        { "profit-pct-max", "50" },
        // { "supply-min", "0" },
        // { "supply-max", "0" },
        // { "demand-min", "0" },
        // { "demand-max", "0" },
        { "sold-day-min", "5000" },
        // { "sold-day-max", "0" },
        { "bought-day-min", "1000" },
        // { "bought-day-max", "0" },
        // { "offers-day-min", "0" },
        // { "offers-day-max", "0" },
        // { "bids-day-min", "0" },
        // { "bids-day-max", "0" },
    };

    private static readonly Dictionary<string, string> MediumArguments = new()
    {
        { "sell-min", "100" },
        // { "sell-max", "0" },
        { "buy-min", "1" },
        { "buy-max", "100000" },
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

    private static readonly List<Dictionary<string, string>> WhichArguments = new() { MediumArguments };
    private static readonly bool UndercutBuys = true;
    private static readonly bool UndercutSells = true;
    private static readonly double Quantity = 0.02;
    private static readonly int MaxSpend = 100000;
    private static readonly double ErrorRange = 0.8;
    private static readonly double ErrorRangeInverse = 1 + 1 - ErrorRange;
    private static readonly double ProfitRange = 0.5;
    private static readonly bool BuyIfSelling = true;
    private static readonly bool CancelUndercutDuringBuy = true;
    private static readonly bool SellForBestIfUnderRange = true;
    private static readonly int BuysPerSellLoop = 40;

    private static readonly List<BuyItem> BuyItemsList = new();
    private static readonly Connection ApiConnection = new(ApiKey);
    private static readonly Random Random = new();
    private static Process? process;
    private static Point? tradingPostPoint;
    private static DateTime timeSinceLastAfkCheck = DateTime.MinValue;
    private static DateTime timeSinceLastBuyListGenerate = DateTime.MinValue;
    private static DateTime timeSinceLastSell = DateTime.MinValue;
    private static int currentLoopIndex = 0;

    private enum TradingPostScreen
    {
        Buy,
        Sell,
        Transactions,
    }

    public static async Task Run()
    {
        process = Array.Find(Process.GetProcessesByName("Gw2-64"), p => p.SessionId == Process.GetCurrentProcess().SessionId);
        if (process == null)
        {
            Logger.Error("Couldn't find process");
            return;
        }

        Logger.Info($"Found process: [{process!.Id}] {process!.MainWindowTitle}");

        process.PriorityClass = ProcessPriorityClass.High;

        foreach (var coherentProcess in Process.GetProcessesByName("CoherentUI_Host").Where(p => p.SessionId == Process.GetCurrentProcess().SessionId))
        {
            coherentProcess.PriorityClass = ProcessPriorityClass.High;
        }

        _ = User32.MoveWindow(process.MainWindowHandle, 0, 5, 1080, 850, true);
        _ = User32.MoveWindow(Process.GetCurrentProcess().MainWindowHandle, 0, 855, 1080, 360, true);

        Input.EnsureForegroundWindow(process);

        tradingPostPoint = await GetTradingPostPoint();
        if (tradingPostPoint == null)
        {
            Logger.Error("Couldn't open trading post");
            return;
        }

        Logger.Info($"Found trading post at: {tradingPostPoint}");

        using var apiClient = new Gw2Client(ApiConnection);

        /*while (true)
        {
            try
            {
                var amount = GetItemWindowPrice(0);
                Logger.Info(amount);
            }
            catch (Exception e)
            {
                Logger.Error(e);
                return;
            }
        }*/

        while (true)
        {
            if (Array.Find(Process.GetProcessesByName("Gw2-64"), p => p.SessionId == Process.GetCurrentProcess().SessionId) == null)
            {
                Logger.Error("Process gone");
                return;
            }

            apiClient.Mumble.Update();
            if (string.IsNullOrEmpty(apiClient.Mumble.CharacterName))
            {
                Logger.Error("Not logged in");
                return;
            }

            tradingPostPoint = await GetTradingPostPoint();
            if (tradingPostPoint == null)
            {
                Logger.Error("Couldn't open trading post");
                return;
            }

            if (currentLoopIndex >= BuyItemsList.Count)
            {
                try
                {
                    await UpdateBuyList();
                }
                catch (Exception e)
                {
                    Logger.Error(e);
                }

                currentLoopIndex = 0;
            }

            if (!CancelUndercutDuringBuy)
            {
                try
                {
                    await CancelUndercut();
                }
                catch (Exception e)
                {
                    Logger.Error(e);
                }
            }

            try
            {
                await BuyItems();
            }
            catch (Exception e)
            {
                Logger.Error(e);
            }

            try
            {
                var howLongLastSell = DateTime.Now - timeSinceLastSell;
                if (howLongLastSell < TimeSpan.FromMinutes(6))
                {
                    var howLongWait = TimeSpan.FromMinutes(6) - howLongLastSell;
                    Logger.Info($"Sleeping for {howLongWait.TotalMinutes} minutes");
                    await Task.Delay(howLongWait);
                }

                timeSinceLastSell = DateTime.Now;
                await SellItems();
            }
            catch (Exception e)
            {
                Logger.Error(e);
            }

            if (ResetUI())
            {
                AntiAfk();
                CheckNewMap();
            }

            /*if (currentLoopIndex >= BuyItemsList.Count)
            {
                var delay = Random.Next((int)TimeSpan.FromMinutes(2).TotalMilliseconds, (int)TimeSpan.FromMinutes(4).TotalMilliseconds);

                Logger.Info($"Sleeping for {TimeSpan.FromMilliseconds(delay).TotalMinutes} minutes");
                await Task.Delay(delay);
            }*/
        }
    }

    private static async Task UpdateBuyList()
    {
        if (DateTime.Now - timeSinceLastBuyListGenerate < TimeSpan.FromMinutes(15))
        {
            return;
        }

        Logger.Info("Updating buy list");

        timeSinceLastBuyListGenerate = DateTime.Now;

        foreach (var arguments in WhichArguments)
        {
            for (var page = 1; page < 10; page++)
            {
                var url = $"https://www.gw2bltc.com/en/tp/search?page={page}&sort=profit&ipg=200";
                foreach (var argument in arguments)
                {
                    url += $"&{argument.Key}={argument.Value}";
                }

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
        }

        BuyItemsList.Sort((x, y) => y.Profit.CompareTo(x.Profit));

        Logger.Info($"Items: {BuyItemsList.Count}");
    }

    private static async Task<Point?> GetTradingPostPoint()
    {
        var point = ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostLion);

        if (point == null)
        {
            Input.KeyPress(process!, VirtualKeyCode.VK_F);

            try
            {
                await WaitWhile(() => ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostLion) == null, 500, 5000);
            }
            catch (TimeoutException)
            {
                if (ResetUI())
                {
                    Input.KeyPress(process!, VirtualKeyCode.VK_F);

                    try
                    {
                        await WaitWhile(() => ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostLion) == null, 500, 5000);
                    }
                    catch (TimeoutException)
                    {
                        Logger.Error("Trading Post Lion not found");
                        LogImage();
                        return null;
                    }
                }
                else
                {
                    Logger.Error("Trading Post Lion not found");
                    LogImage();
                    return null;
                }
            }

            point = ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostLion);
            Input.KeyPress(process!, VirtualKeyCode.ESCAPE);
            await Task.Delay(2000);
        }

        return point;
    }

    private static async Task OpenTradingPost()
    {
        var point = ImageSearch.FindImageInWindow(process!, Resources.TradingPostLion, tradingPostPoint!.Value.X, tradingPostPoint!.Value.Y, Resources.TradingPostLion.Width, Resources.TradingPostLion.Height);

        if (point == null)
        {
            Logger.Info("Lost sight of trading post");
            Input.KeyPress(process!, VirtualKeyCode.VK_F);

            try
            {
                await WaitWhile(() => ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostHome, 0.9) == null, 500, 5000);
            }
            catch (TimeoutException)
            {
                if (ResetUI())
                {
                    Input.KeyPress(process!, VirtualKeyCode.VK_F);

                    try
                    {
                        await WaitWhile(() => ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostHome, 0.9) == null, 500, 5000);
                    }
                    catch (TimeoutException)
                    {
                        throw new TimeoutException();
                    }
                }
                else
                {
                    throw new TimeoutException();
                }
            }

            tradingPostPoint = ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostLion);
            Logger.Info($"Found trading post at: {tradingPostPoint}");

            await Task.Delay(2000);
        }
    }

    private static bool ResetUI()
    {
        Logger.Info("Attempting reset UI");

        for (var i = 0; i < 10; i++)
        {
            Input.KeyPress(process!, VirtualKeyCode.ESCAPE);
            Thread.Sleep(2000);

            if (ImageSearch.FindImageInFullWindow(process!, Resources.ReturnToGame, 0.8) != null)
            {
                Logger.Info("Found escape menu");
                Input.KeyPress(process!, VirtualKeyCode.ESCAPE);
                Thread.Sleep(2000);
                return true;
            }
        }

        Logger.Debug("Wasn't able to reset UI");
        LogImage();
        return false;
    }

    private static async Task GoToScreen(TradingPostScreen screen)
    {
        await OpenTradingPost();

        // Click take all
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 179, tradingPostPoint!.Value.Y + 685);
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 179, tradingPostPoint!.Value.Y + 685);

        Func<bool>? findFunc = null;
        switch (screen)
        {
            case TradingPostScreen.Sell:
                // Check if already open
                if (ImageSearch.FindImageInWindow(process!, Resources.SellItemsActive, tradingPostPoint!.Value.X + 805, tradingPostPoint!.Value.Y + 130, Resources.SellItemsActive.Width, Resources.SellItemsActive.Height, 0.99) != null)
                {
                    return;
                }

                // Click sell items
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 732, tradingPostPoint!.Value.Y + 90);

                static bool FindSellItemsActive()
                {
                    var find = ImageSearch.FindImageInWindow(process!, Resources.SellItemsActive, tradingPostPoint!.Value.X + 805, tradingPostPoint!.Value.Y + 130, Resources.SellItemsActive.Width, Resources.SellItemsActive.Height, 0.99);
                    return find == null;
                }

                findFunc = FindSellItemsActive;
                break;
            case TradingPostScreen.Buy:
                // Check if already open
                if (ImageSearch.FindImageInWindow(process!, Resources.BuyItemsActive, tradingPostPoint!.Value.X + 638, tradingPostPoint!.Value.Y + 130, Resources.BuyItemsActive.Width, Resources.BuyItemsActive.Height, 0.99) != null)
                {
                    return;
                }

                // Click buy items
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 566, tradingPostPoint!.Value.Y + 90);

                static bool FindBuyItemsActive()
                {
                    var find = ImageSearch.FindImageInWindow(process!, Resources.BuyItemsActive, tradingPostPoint!.Value.X + 638, tradingPostPoint!.Value.Y + 130, Resources.BuyItemsActive.Width, Resources.BuyItemsActive.Height, 0.99);
                    return find == null;
                }

                findFunc = FindBuyItemsActive;
                break;
            case TradingPostScreen.Transactions:
                // Check if already open
                if (ImageSearch.FindImageInWindow(process!, Resources.TransactionsActive, tradingPostPoint!.Value.X + 817, tradingPostPoint!.Value.Y + 130, Resources.TransactionsActive.Width, Resources.TransactionsActive.Height, 0.99) != null)
                {
                    return;
                }

                // Click transactions
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 899, tradingPostPoint!.Value.Y + 90);

                static bool FindTransactionsActive()
                {
                    var find = ImageSearch.FindImageInWindow(process!, Resources.TransactionsActive, tradingPostPoint!.Value.X + 817, tradingPostPoint!.Value.Y + 130, Resources.TransactionsActive.Width, Resources.TransactionsActive.Height, 0.99);
                    return find == null;
                }

                findFunc = FindTransactionsActive;
                break;
        }

        // Wait for it to be activated
        try
        {
            await WaitWhile(findFunc!, 500, 5000);
        }
        catch (TimeoutException)
        {
            LogImage();
            throw new TimeoutException();
        }

        await Task.Delay(1000);
    }

    private static async Task CloseItemWindow()
    {
        if (ImageSearch.FindImageInWindow(process!, Resources.Exit, tradingPostPoint!.Value.X + 769, tradingPostPoint!.Value.Y + 117, Resources.Exit.Width, Resources.Exit.Height, 0.9) != null)
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 777, tradingPostPoint!.Value.Y + 125);
        }

        await Task.Delay(500);
    }

    private static async Task SearchForItem(Gw2Sharp.WebApi.V2.Models.Item itemInfo, TradingPostScreen screen)
    {
        // Click search bar
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 163, tradingPostPoint!.Value.Y + 165);

        if (screen == TradingPostScreen.Buy)
        {
            // Check if filter isn't open
            if (ImageSearch.FindImageInWindow(process!, Resources.FilterOpen, tradingPostPoint!.Value.X + 50, tradingPostPoint!.Value.Y + 182, Resources.FilterOpen.Width, Resources.FilterOpen.Height, 0.9) == null)
            {
                // Click filter cog
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 287, tradingPostPoint!.Value.Y + 165);

                // Wait for filter to show up
                try
                {
                    await WaitWhile(FindFilter, 500, 5000);
                }
                catch (TimeoutException)
                {
                    LogImage();
                    throw new TimeoutException();
                }
            }

            // Click reset filters
            if (screen == TradingPostScreen.Buy)
            {
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 173, tradingPostPoint!.Value.Y + 356);
                Thread.Sleep(1000);
            }
            else
            {
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 173, tradingPostPoint!.Value.Y + 309);
            }
        }

        // Click search bar
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 163, tradingPostPoint!.Value.Y + 165);

        if (screen != TradingPostScreen.Buy)
        {
            // Select all
            Input.KeyPressWithModifier(process!, VirtualKeyCode.VK_A, false, true, false);

            // Backspace
            Input.KeyPress(process!, VirtualKeyCode.BACK);
        }

        // Paste in name
        Input.KeyStringSendClipboard(process!, itemInfo.Name.MaxSize(30));
        Input.KeyPress(process!, VirtualKeyCode.RETURN);

        if (screen == TradingPostScreen.Buy)
        {
            // Set rarity
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 173, tradingPostPoint!.Value.Y + 205);
            await Task.Delay(500);
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 173, tradingPostPoint!.Value.Y + 230 + (25 * RarityOrder[itemInfo.Rarity!]));
            await Task.Delay(500);

            // Input level
            Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 77, tradingPostPoint!.Value.Y + 270);
            Input.KeyStringSend(process!, itemInfo.Level.ToString());
            Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 138, tradingPostPoint!.Value.Y + 270);
            Input.KeyStringSend(process!, itemInfo.Level.ToString());

            // Click filter cog to close
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 287, tradingPostPoint!.Value.Y + 165);
        }
    }

    private static int GetItemWindowPrice(int offset)
    {
        /*Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 402, tradingPostPoint.Value.Y + 236 + offset);
        Input.KeyPress(process!, VirtualKeyCode.TAB);
        var goldAmount = Convert.ToInt32(Input.GetSelectedText(process!));
        Input.KeyPress(process!, VirtualKeyCode.TAB);
        var silverAmount = Convert.ToInt32(Input.GetSelectedText(process!));
        Input.KeyPress(process!, VirtualKeyCode.TAB);
        var copperAmount = Convert.ToInt32(Input.GetSelectedText(process!));*/

        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 412, tradingPostPoint.Value.Y + 279 + offset);
        var goldAmount = Convert.ToInt32(Input.GetSelectedText(process!));
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 502, tradingPostPoint.Value.Y + 279 + offset);
        var silverAmount = Convert.ToInt32(Input.GetSelectedText(process!));
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 576, tradingPostPoint.Value.Y + 279 + offset);
        var copperAmount = Convert.ToInt32(Input.GetSelectedText(process!));

        return ToCurrency(goldAmount, silverAmount, copperAmount);
    }

    private static int GetItemWindowPriceWithDismiss(int offset)
    {
        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 412, tradingPostPoint.Value.Y + 279 + offset);
        var goldAmount = Convert.ToInt32(Input.GetSelectedText(process!));
        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 502, tradingPostPoint.Value.Y + 279 + offset);
        var silverAmount = Convert.ToInt32(Input.GetSelectedText(process!));
        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 576, tradingPostPoint.Value.Y + 279 + offset);
        var copperAmount = Convert.ToInt32(Input.GetSelectedText(process!));

        return ToCurrency(goldAmount, silverAmount, copperAmount);
    }

    private static void SetItemWindowPrice(int offset, int coins)
    {
        // Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 402, tradingPostPoint.Value.Y + 236 + offset);
        // Input.KeyPress(process!, VirtualKeyCode.TAB);
        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 412, tradingPostPoint.Value.Y + 279 + offset);
        Input.KeyStringSend(process!, ToGold(coins).ToString());
        // Input.KeyPress(process!, VirtualKeyCode.TAB);
        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 502, tradingPostPoint.Value.Y + 279 + offset);
        Input.KeyStringSend(process!, ToSilver(coins).ToString());
        // Input.KeyPress(process!, VirtualKeyCode.TAB);
        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 576, tradingPostPoint.Value.Y + 279 + offset);
        Input.KeyStringSend(process!, ToCopper(coins).ToString());
    }

    private static async Task SellItems()
    {
        Logger.Info("====================");
        Logger.Info("Selling items");

        using var apiClient = new Gw2Client(ApiConnection);

        apiClient.Mumble.Update();
        var character = apiClient.Mumble.CharacterName;

        var backpack = await apiClient.WebApi.V2.Characters[character].Inventory.GetAsync();
        var inventoryItems = backpack.Bags.SelectMany(x => x.Inventory.Select(y => y?.Id));

        var currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(0);
        var currentSellingList = currentSelling.ToList();
        var numSellingPages = currentSelling.HttpResponseInfo?.PageTotal;
        for (var i = 1; i < numSellingPages; i++)
        {
            currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(i);
            currentSellingList = currentSellingList.Concat(currentSelling.ToList()).ToList();
        }

        foreach (var item in BuyItemsList)
        {
            if (!inventoryItems.Contains(item.Id))
            {
                continue;
            }

            // Get item info
            Gw2Sharp.WebApi.V2.Models.Item? itemInfo = null;
            try
            {
                itemInfo = await apiClient.WebApi.V2.Items.GetAsync(item.Id);
            }
            catch (Exception e)
            {
                Logger.Error(e);
                continue;
            }

            if (itemInfo == null)
            {
                Logger.Warn("itemInfo null");
                continue;
            }

            Logger.Info("--------------------");
            Logger.Info($"Selling [{item.Id}] {item.Name}");
            Logger.Info($"Buy: {item.BuyPrice} Sell: {item.SellPrice} Profit: {item.Profit} Bought: {item.Bought} Sold: {item.Sold}");

            try
            {
                await GoToScreen(TradingPostScreen.Sell);
            }
            catch (TimeoutException)
            {
                Logger.Debug("Timeout on GoToScreen");
                continue;
            }

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

            await Task.Delay(500);

            // Wait for search results to load
            try
            {
                await WaitWhile(FindSearchResultsLoaded, 500, 5000);
            }
            catch (TimeoutException)
            {
                Logger.Error("Didn't load any search results");
                LogImage();
                continue;
            }

            // Check if no item found
            if (ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.8) != null)
            {
                Logger.Debug("No item found");
                continue;
            }

            // Find position of Qty to determine position of search results
            var qtyPoint = ImageSearch.FindImageInWindow(process!, Resources.Qty, tradingPostPoint!.Value.X + 324, tradingPostPoint!.Value.Y + 155, Resources.Qty.Width, Resources.Qty.Height + 72, 0.9);
            if (qtyPoint == null)
            {
                Logger.Warn("qtyPoint null");
                LogImage();
                continue;
            }

            var resultsPoint = Point.Add(qtyPoint.Value, new Size(-2, 22));

            for (var y = 0; y < 7; y++)
            {
                var attempts = 0;
                SellItem:
                if (attempts >= 5)
                {
                    break;
                }

                attempts++;

                // Find the item border of the result
                var resultPoint = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, resultsPoint.X, resultsPoint.Y + (y * 61), Resources.ResultCorner.Width, Resources.ResultCorner.Height);
                if (resultPoint == null)
                {
                    Logger.Debug("Result border not found");
                    LogImage();
                    // break;
                    goto SellItem;
                }

                // Click the result
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 28, resultsPoint.Y + 27 + (y * 61));

                // Wait for the item border to show up
                try
                {
                    await WaitWhile(FindItemBorder, 500, 5000);
                }
                catch (TimeoutException)
                {
                    Logger.Warn("Item border didn't appear");
                    LogImage();
                    await CloseItemWindow();
                    // break;
                    goto SellItem;
                }

                DismissSuccess();

                // Wait for "vendor" to show up
                try
                {
                    await WaitWhile(FindVendorE, 500, 5000);
                }
                catch (TimeoutException)
                {
                    Logger.Warn("Vendor E didn't appear");
                    LogImage();
                    await CloseItemWindow();
                    // break;
                    goto SellItem;
                }

                DismissSuccess();

                // Find position of 'e' in Vendor to set the offset for other UI positions
                var vendorEPos = ImageSearch.FindImageInWindow(process!, Resources.VendorE, tradingPostPoint!.Value.X + 345, tradingPostPoint!.Value.Y + 176, Resources.VendorE.Width, Resources.VendorE.Height + 30, 0.9);
                if (vendorEPos == null)
                {
                    Logger.Warn("vendorEPos null");
                    LogImage();
                    await CloseItemWindow();
                    // break;
                    goto SellItem;
                }

                var offset = 0;
                var nameSize = 34;
                if (vendorEPos.Value.Y == tradingPostPoint!.Value.Y + 206)
                {
                    offset = 24;
                    nameSize = 64;
                }

                // Get OCR of item name and check if it matches API name
                var nameImage = ImageSearch.CaptureWindow(process!, tradingPostPoint!.Value.X + 334, tradingPostPoint.Value.Y + 136, 428, nameSize);
                var capturedName = OCR.ReadName(nameImage!, RarityColors[itemInfo.Rarity.ToString()!]);
                if (string.IsNullOrEmpty(capturedName))
                {
                    Logger.Debug("Empty string returned");
                    LogImage();
                    await CloseItemWindow();
                    continue;
                }

                var jw = new JaroWinkler();
                var nameDistance = 1.0 - jw.Distance(itemInfo.Name.RemoveDiacritics().ToLower(), capturedName.ToLower());
                Logger.Debug($"Item name: {itemInfo.Name} Captured name: {capturedName} Distance: {nameDistance}");
                if (nameDistance < 0.99)
                {
                    await CloseItemWindow();
                    continue;
                }

                DismissSuccessWithOffset(offset);

                // Wait for available to show up
                try
                {
                    await WaitWhile(FindAvailable, 500, 10000);
                }
                catch (TimeoutException)
                {
                    Logger.Warn("Available didn't appear");
                    await CloseItemWindow();
                    // break;
                    goto SellItem;
                }

                // Get buy price
                int? buyCurrencyAmount = null;
                try
                {
                    buyCurrencyAmount = GetItemWindowPriceWithDismiss(offset);
                }
                catch (FormatException e)
                {
                    Logger.Error(e);
                    LogImage();
                    await CloseItemWindow();
                    // break;
                    goto SellItem;
                }

                Logger.Info($"Highest buyer: {buyCurrencyAmount}");

                DismissSuccessWithOffset(offset);

                // Click highest current seller
                await Task.Delay(100);
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint.Value.X + 542, tradingPostPoint.Value.Y + 499 + offset);

                // Get sell price
                int? sellCurrencyAmount = null;
                try
                {
                    sellCurrencyAmount = GetItemWindowPriceWithDismiss(offset);
                }
                catch (FormatException e)
                {
                    Logger.Error(e);
                    LogImage();
                    await CloseItemWindow();
                    // break;
                    goto SellItem;
                }

                Logger.Info($"Lowest seller: {sellCurrencyAmount}");

                // Check if prices are wrong
                if ((sellCurrencyAmount > item.SellPrice * ErrorRangeInverse || sellCurrencyAmount < item.SellPrice * ErrorRange)
                    && (buyCurrencyAmount > item.BuyPrice * ErrorRangeInverse || buyCurrencyAmount < item.BuyPrice * ErrorRange))
                {
                    Logger.Info("Sell & buy price too different");
                    await CloseItemWindow();
                    continue;
                }

                // Check if sell price is within range
                /*if (sellCurrencyAmount < item.SellPrice * ErrorRange)
                {
                    Logger.Info($"Sell price {sellCurrencyAmount} below range {item.SellPrice * ErrorRange}");
                    await CloseItemWindow();
                    break;
                }*/

                var buyPrice = buyCurrencyAmount!.Value;
                var sellPrice = sellCurrencyAmount!.Value;

                if (UndercutBuys)
                {
                    buyPrice++;
                }

                DismissSuccessWithOffset(offset);

                // Check profit is within range
                var profit = GetProfit(item.BuyPrice, sellPrice);
                var lowestProfit = (int)(item.Profit * ProfitRange);
                if (profit < lowestProfit)
                {
                    Logger.Info($"Profit of {profit} outside of range");

                    // Find best price in range and undercut that instead
                    if (SellForBestIfUnderRange)
                    {
                        Logger.Info($"Selling for best near profit range {lowestProfit}");
                        var listings = await apiClient.WebApi.V2.Commerce.Listings.GetAsync(item.Id);

                        var profitableSell = GetSellPriceForProfit(item.BuyPrice, lowestProfit);
                        sellPrice = listings.Sells.Where(x => x.UnitPrice >= profitableSell).Min(x => x.UnitPrice);

                        SetItemWindowPrice(offset, sellPrice);

                        // Get buy price
                        int? enteredAmount = null;
                        try
                        {
                            enteredAmount = GetItemWindowPriceWithDismiss(offset);
                        }
                        catch (FormatException e)
                        {
                            Logger.Error(e);
                            LogImage();
                            await CloseItemWindow();
                            // break;
                            goto SellItem;
                        }

                        if (enteredAmount != sellPrice)
                        {
                            Logger.Debug("Didn't enter price correctly");
                            await CloseItemWindow();
                            // break;
                            goto SellItem;
                        }
                    }
                    else
                    {
                        await CloseItemWindow();
                        break;
                    }
                }

                // Check if not my listing
                if (UndercutSells && !currentSellingList.Any(x => x.ItemId == item.Id && x.Price == sellCurrencyAmount))
                {
                    // Click to -1 copper
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint.Value.X + 593, tradingPostPoint.Value.Y + 283 + offset);
                    sellPrice--;
                }

                // Set quantity to max
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 625, tradingPostPoint.Value.Y + 240 + offset);
                await Task.Delay(100);
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 625, tradingPostPoint.Value.Y + 240 + offset);

                Logger.Info($"Selling for {sellPrice} coins");

                // Sell item
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 440, tradingPostPoint.Value.Y + 373 + offset);
                await CloseItemWindow();
                break;
            }
        }
    }

    private static async Task BuyItems()
    {
        Logger.Info("====================");
        Logger.Info("Buying items");

        using var apiClient = new Gw2Client(ApiConnection);

        apiClient.Mumble.Update();
        var character = apiClient.Mumble.CharacterName;

        List<Gw2Sharp.WebApi.V2.Models.CommerceTransactionCurrent>? currentSellingList = null;
        if (!BuyIfSelling)
        {
            var currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(0);
            currentSellingList = currentSelling.ToList();
            var numSellingPages = currentSelling.HttpResponseInfo?.PageTotal;
            for (var i = 1; i < numSellingPages; i++)
            {
                currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(i);
                currentSellingList = currentSellingList.Concat(currentSelling.ToList()).ToList();
            }
        }

        var currentBuying = await apiClient.WebApi.V2.Commerce.Transactions.Current.Buys.PageAsync(0);
        var currentBuyingList = currentBuying.ToList();
        var numBuyingPages = currentBuying.HttpResponseInfo?.PageTotal;
        for (var i = 1; i < numBuyingPages; i++)
        {
            currentBuying = await apiClient.WebApi.V2.Commerce.Transactions.Current.Buys.PageAsync(i);
            currentBuyingList = currentBuyingList.Concat(currentBuying.ToList()).ToList();
        }

        // foreach (var item in BuyItemsList)
        // {

        for (var endLoopIndex = currentLoopIndex + BuysPerSellLoop; currentLoopIndex < endLoopIndex; currentLoopIndex++)
        {
            if (currentLoopIndex >= BuyItemsList.Count)
            {
                Logger.Info("Finished current loop of buys");
                return;
            }

            var item = BuyItemsList[currentLoopIndex];

            Logger.Info("--------------------");
            Logger.Info($"Buying Item {currentLoopIndex + 1}/{BuyItemsList.Count} [{item.Id}] {item.Name}");
            Logger.Info($"Buy: {item.BuyPrice} Sell: {item.SellPrice} Profit: {item.Profit} Bought: {item.Bought} Sold: {item.Sold}");

            // Check if currently selling item
            if (!BuyIfSelling && currentSellingList!.Any(x => x.ItemId == item.Id))
            {
                Logger.Info("Skipping, already selling");
                continue;
            }

            // Check if already buying item
            if (currentBuyingList.Any(x => x.ItemId == item.Id))
            {
                if (CancelUndercutDuringBuy)
                {
                    try
                    {
                        var itemPrices = await apiClient.WebApi.V2.Commerce.Prices.GetAsync(item.Id);
                        var highestPrice = itemPrices.Buys.UnitPrice;
                        var price = currentBuyingList.Where(x => x.ItemId == item.Id).Max(x => x.Price);

                        // If we haven't been undercut
                        if (highestPrice <= price)
                        {
                            Logger.Info("Skipping, already buying");
                            continue;
                        }
                    }
                    catch (Exception e)
                    {
                        Logger.Error(e);
                        continue;
                    }

                    Logger.Info("Cancelling undercut");
                    await CancelItem(item);
                }
                else
                {
                    Logger.Info("Skipping, already buying");
                    continue;
                }
            }

            try
            {
                await GoToScreen(TradingPostScreen.Buy);
            }
            catch (TimeoutException)
            {
                Logger.Debug("Timeout on GoToScreen");
                continue;
            }

            // Get item info
            Gw2Sharp.WebApi.V2.Models.Item? itemInfo = null;
            try
            {
                itemInfo = await apiClient.WebApi.V2.Items.GetAsync(item.Id);
            }
            catch (Exception e)
            {
                Logger.Error(e);
                continue;
            }

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

            await Task.Delay(2500);

            // Wait for search results to load
            try
            {
                await WaitWhile(FindSearchResultsLoaded, 500, 5000);
            }
            catch (TimeoutException)
            {
                Logger.Error("Didn't load any search results");
                LogImage();
                continue;
            }

            // Check if no item found
            if (ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.8) != null)
            {
                Logger.Debug("No item found");
                continue;
            }

            // Find position of Qty to determine position of search results
            var qtyPoint = ImageSearch.FindImageInWindow(process!, Resources.Qty, tradingPostPoint!.Value.X + 324, tradingPostPoint!.Value.Y + 155, Resources.Qty.Width, Resources.Qty.Height + 72, 0.9);
            if (qtyPoint == null)
            {
                Logger.Warn("qtyPoint null");
                LogImage();
                continue;
            }

            var resultsPoint = Point.Add(qtyPoint.Value, new Size(-2, 22));

            for (var y = 0; y < 7; y++)
            {
                var attempts = 0;
                BuyItem:
                if (attempts >= 5)
                {
                    break;
                }

                attempts++;

                // Find the item border of the result
                var resultPoint = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, resultsPoint.X, resultsPoint.Y + (y * 61), Resources.ResultCorner.Width, Resources.ResultCorner.Height);
                if (resultPoint == null)
                {
                    Logger.Debug("Result border not found");
                    LogImage();
                    break;
                }

                // Click the result
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 28, resultsPoint.Y + 27 + (y * 61));

                // Wait for the item border to show up
                try
                {
                    await WaitWhile(FindItemBorder, 500, 5000);
                }
                catch (TimeoutException)
                {
                    Logger.Warn("Item border didn't appear");
                    await CloseItemWindow();
                    // break;
                    goto BuyItem;
                }

                // Wait for "vendor" to show up
                try
                {
                    await WaitWhile(FindVendorE, 500, 5000);
                }
                catch (TimeoutException)
                {
                    Logger.Warn("Vendor E didn't appear");
                    LogImage();
                    await CloseItemWindow();
                    // break;
                    goto BuyItem;
                }

                // Find position of 'e' in Vendor to set the offset for other UI positions
                var vendorEPos = ImageSearch.FindImageInWindow(process!, Resources.VendorE, tradingPostPoint!.Value.X + 345, tradingPostPoint!.Value.Y + 176, Resources.VendorE.Width, Resources.VendorE.Height + 30, 0.9);
                if (vendorEPos == null)
                {
                    Logger.Warn("vendorEPos null");
                    LogImage();
                    await CloseItemWindow();
                    // break;
                    goto BuyItem;
                }

                var offset = 0;
                var nameSize = 34;
                if (vendorEPos.Value.Y == tradingPostPoint!.Value.Y + 206)
                {
                    offset = 24;
                    nameSize = 64;
                }

                // Get OCR of item name and check if it matches API name
                var nameImage = ImageSearch.CaptureWindow(process!, tradingPostPoint!.Value.X + 334, tradingPostPoint.Value.Y + 136, 428, nameSize);
                var capturedName = OCR.ReadName(nameImage!, RarityColors[itemInfo.Rarity.ToString()!]);
                if (string.IsNullOrEmpty(capturedName))
                {
                    Logger.Debug("Empty string returned");
                    LogImage();
                    await CloseItemWindow();
                    continue;
                }

                var jw = new JaroWinkler();
                var nameDistance = 1.0 - jw.Distance(itemInfo.Name.RemoveDiacritics().ToLower(), capturedName.ToLower());
                Logger.Debug($"Item name: {itemInfo.Name} Captured name: {capturedName} Distance: {nameDistance}");
                if (nameDistance < 0.99)
                {
                    await CloseItemWindow();
                    continue;
                }

                // Wait for available to show up
                try
                {
                    await WaitWhile(FindAvailable, 500, 10000);
                }
                catch (TimeoutException)
                {
                    Logger.Warn("Available didn't appear");
                    LogImage();
                    await CloseItemWindow();
                    break;
                }

                // Get sell price
                int? sellCurrencyAmount = null;
                try
                {
                    sellCurrencyAmount = GetItemWindowPrice(offset);
                }
                catch (FormatException e)
                {
                    Logger.Error(e);
                    LogImage();
                    await CloseItemWindow();
                    // break;
                    goto BuyItem;
                }

                Logger.Info($"Lowest seller: {sellCurrencyAmount}");

                // Click highest current buyer
                await Task.Delay(100);
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint.Value.X + 254, tradingPostPoint.Value.Y + 499 + offset);

                // Get buy price
                int? buyCurrencyAmount = null;
                try
                {
                    buyCurrencyAmount = GetItemWindowPrice(offset);
                }
                catch (FormatException e)
                {
                    Logger.Error(e);
                    LogImage();
                    await CloseItemWindow();
                    // break;
                    goto BuyItem;
                }

                Logger.Info($"Highest buyer: {buyCurrencyAmount}");

                // Check if prices are wrong
                if ((sellCurrencyAmount > item.SellPrice * ErrorRangeInverse || sellCurrencyAmount < item.SellPrice * ErrorRange)
                    && (buyCurrencyAmount > item.BuyPrice * ErrorRangeInverse || buyCurrencyAmount < item.BuyPrice * ErrorRange))
                {
                    Logger.Info("Sell & buy price too different");
                    await CloseItemWindow();
                    continue;
                }

                // Check if sell price is within range
                if (sellCurrencyAmount < item.SellPrice * ErrorRange)
                {
                    Logger.Info("Sell price below range");
                    await CloseItemWindow();
                    break;
                }

                // Check if buy price is within range
                if (buyCurrencyAmount > item.BuyPrice * ErrorRangeInverse)
                {
                    Logger.Info("Buy price above range");
                    await CloseItemWindow();
                    break;
                }

                var buyPrice = buyCurrencyAmount!.Value;
                var sellPrice = sellCurrencyAmount!.Value;

                if (UndercutBuys)
                {
                    buyPrice++;
                }

                if (UndercutSells)
                {
                    sellPrice--;
                }

                // Check profit is within range
                var profit = GetProfit(buyPrice, sellPrice);
                if (profit < (item.Profit * ErrorRange))
                {
                    Logger.Info($"Profit of {profit} outside of range");
                    await CloseItemWindow();
                    break;
                }

                if (UndercutBuys)
                {
                    // Click to +1 copper
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint.Value.X + 593, tradingPostPoint.Value.Y + 272 + offset);
                }

                var quantity = Math.Min(Math.Min((int)(MaxSpend / buyPrice), (int)(item.Sold * Quantity)), 250);
                quantity = quantity < 1 ? 1 : quantity;

                Logger.Info($"Buying {quantity} for {buyPrice} coins each");

                if (quantity > 1)
                {
                    // Set quantity
                    Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 402, tradingPostPoint.Value.Y + 236 + offset);
                    Input.KeyStringSend(process!, quantity.ToString());
                }
                else
                {
                    // Set quantity to 1
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 448, tradingPostPoint.Value.Y + 240 + offset);
                }

                // Buy item
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 440, tradingPostPoint.Value.Y + 373 + offset);
                await CloseItemWindow();
                break;
            }
        }
    }

    private static async Task CancelItem(BuyItem item)
    {
        try
        {
            await GoToScreen(TradingPostScreen.Transactions);
        }
        catch (TimeoutException)
        {
            Logger.Debug("Timeout on GoToScreen");
            return;
        }

        // Get item info
        Gw2Sharp.WebApi.V2.Models.Item? itemInfo = null;
        try
        {
            using var apiClient = new Gw2Client(ApiConnection);
            itemInfo = await apiClient.WebApi.V2.Items.GetAsync(item.Id);
        }
        catch (Exception e)
        {
            Logger.Error(e);
            return;
        }

        if (itemInfo == null)
        {
            Logger.Warn("itemInfo null");
            return;
        }

        // Click buying if not selected
        if (ImageSearch.FindImageInWindow(process!, Resources.CurrentBuying, tradingPostPoint!.Value.X + 60, tradingPostPoint!.Value.Y + 224, Resources.CurrentBuying.Width, Resources.CurrentBuying.Height, 0.9) != null)
        {
            // Wait for search results to load
            try
            {
                await WaitWhile(FindSearchResultsLoaded, 500, 10000);
            }
            catch (TimeoutException)
            {
                Input.KeyPress(process!, VirtualKeyCode.ESCAPE);
                await Task.Delay(2000);
                return;
            }

            // Click load more
            for (var i = 0; i < 5; i++)
            {
                if (ImageSearch.FindImageInWindow(process!, Resources.LoadMore, tradingPostPoint!.Value.X + 610, tradingPostPoint!.Value.Y + 679, Resources.LoadMore.Width, Resources.LoadMore.Height, 0.5) == null)
                {
                    break;
                }

                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 650, tradingPostPoint!.Value.Y + 688);
                await Task.Delay(5000);
            }

            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 174, tradingPostPoint!.Value.Y + 240);
            await Task.Delay(1000);

            // Wait for search results to load
            try
            {
                await WaitWhile(FindSearchResultsLoaded, 500, 10000);
            }
            catch (TimeoutException)
            {
                Input.KeyPress(process!, VirtualKeyCode.ESCAPE);
                await Task.Delay(2000);
                return;
            }
        }

        // Click load more
        for (var i = 0; i < 5; i++)
        {
            if (ImageSearch.FindImageInWindow(process!, Resources.LoadMore, tradingPostPoint!.Value.X + 610, tradingPostPoint!.Value.Y + 679, Resources.LoadMore.Width, Resources.LoadMore.Height, 0.5) == null)
            {
                break;
            }

            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 650, tradingPostPoint!.Value.Y + 688);
            await Task.Delay(5000);
        }

        // Search
        try
        {
            await SearchForItem(itemInfo, TradingPostScreen.Transactions);
        }
        catch (TimeoutException)
        {
            Logger.Warn("Filter didn't appear");
            return;
        }

        await Task.Delay(500);

        // Wait for search results to load
        try
        {
            await WaitWhile(FindSearchResultsLoaded, 500, 5000);
        }
        catch (TimeoutException)
        {
            Logger.Error("Didn't load any search results");
            LogImage();
            return;
        }

        // Check if no item found
        if (ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.8) != null)
        {
            Logger.Debug("No item found");
            return;
        }

        // Find position of Qty to determine position of search results
        var qtyPoint = ImageSearch.FindImageInWindow(process!, Resources.Qty, tradingPostPoint.Value.X + 324, tradingPostPoint.Value.Y + 155, Resources.Qty.Width, Resources.Qty.Height + 72, 0.9);
        if (qtyPoint == null)
        {
            Logger.Warn("qtyPoint null");
            LogImage();
            return;
        }

        var resultsPoint = Point.Add(qtyPoint.Value, new Size(-2, 22));

        for (var y = 0; y < 7; y++)
        {
            var attempts = 0;
            CancelItem:
            if (attempts >= 5)
            {
                break;
            }

            attempts++;

            // Find the item border of the result
            var resultPoint = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, resultsPoint.X, resultsPoint.Y + (y * 61), Resources.ResultCorner.Width, Resources.ResultCorner.Height);
            if (resultPoint == null)
            {
                break;
            }

            // Click the result
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 28, resultsPoint.Y + 27 + (y * 61));

            // Wait for the item border to show up
            try
            {
                await WaitWhile(FindItemBorder, 500, 5000);
            }
            catch (TimeoutException)
            {
                Logger.Warn("Item border didn't appear");
                LogImage();
                await CloseItemWindow();
                // break;
                goto CancelItem;
            }

            // Wait for "vendor" to show up
            try
            {
                await WaitWhile(FindVendorE, 500, 5000);
            }
            catch (TimeoutException)
            {
                Logger.Warn("Vendor E didn't appear");
                LogImage();
                await CloseItemWindow();
                // break;
                goto CancelItem;
            }

            // Find position of 'e' in Vendor to set the offset for other UI positions
            var vendorEPos = ImageSearch.FindImageInWindow(process!, Resources.VendorE, tradingPostPoint!.Value.X + 345, tradingPostPoint!.Value.Y + 176, Resources.VendorE.Width, Resources.VendorE.Height + 30, 0.9);
            if (vendorEPos == null)
            {
                Logger.Warn("vendorEPos null");
                LogImage();
                await CloseItemWindow();
                // break;
                goto CancelItem;
            }

            var nameSize = 34;
            if (vendorEPos.Value.Y == tradingPostPoint!.Value.Y + 206)
            {
                nameSize = 64;
            }

            // Get OCR of item name and check if it matches API name
            var nameImage = ImageSearch.CaptureWindow(process!, tradingPostPoint!.Value.X + 334, tradingPostPoint.Value.Y + 136, 428, nameSize);
            var capturedName = OCR.ReadName(nameImage!, RarityColors[itemInfo.Rarity.ToString()!]);
            if (string.IsNullOrEmpty(capturedName))
            {
                Logger.Debug("Empty string returned");
                LogImage();
                await CloseItemWindow();
                continue;
            }

            var jw = new JaroWinkler();
            var nameDistance = 1.0 - jw.Distance(itemInfo.Name.RemoveDiacritics().ToLower(), capturedName.ToLower());
            Logger.Debug($"Item name: {itemInfo.Name} Captured name: {capturedName} Distance: {nameDistance}");
            if (nameDistance < 0.99)
            {
                await CloseItemWindow();
                continue;
            }

            // Press X to close
            await CloseItemWindow();

            // Click cancel and confirm
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 600, resultsPoint.Y + 28 + (y * 61));
            await Task.Delay(200);
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 600, resultsPoint.Y + 28 + (y * 61));
            await Task.Delay(200);

            y--;
        }
    }

    private static async Task CancelUndercut()
    {
        Logger.Info("====================");
        Logger.Info("Cancelling undercut buys");

        try
        {
            await GoToScreen(TradingPostScreen.Transactions);
        }
        catch (TimeoutException)
        {
            return;
        }

        // Click buying if not selected
        if (ImageSearch.FindImageInWindow(process!, Resources.CurrentBuying, tradingPostPoint!.Value.X + 60, tradingPostPoint!.Value.Y + 224, Resources.CurrentBuying.Width, Resources.CurrentBuying.Height, 0.9) != null)
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 174, tradingPostPoint!.Value.Y + 240);
        }

        // Wait for search results to load
        try
        {
            await WaitWhile(FindSearchResultsLoaded, 500, 5000);
        }
        catch (TimeoutException)
        {
            Logger.Error("Didn't load any search results");
            LogImage();
            return;
        }

        // Check if no item found
        if (ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.8) != null)
        {
            return;
        }

        using var apiClient = new Gw2Client(ApiConnection);

        var currentBuying = await apiClient.WebApi.V2.Commerce.Transactions.Current.Buys.PageAsync(0);
        var currentBuyingList = currentBuying.ToList();
        var numBuyingPages = currentBuying.HttpResponseInfo?.PageTotal;
        for (var i = 1; i < numBuyingPages; i++)
        {
            currentBuying = await apiClient.WebApi.V2.Commerce.Transactions.Current.Buys.PageAsync(i);
            currentBuyingList = currentBuyingList.Concat(currentBuying.ToList()).ToList();
        }

        foreach (var item in BuyItemsList)
        {
            // Check if we're currently buying item
            if (currentBuyingList.Any(x => x.ItemId == item.Id))
            {
                try
                {
                    var itemPrices = await apiClient.WebApi.V2.Commerce.Prices.GetAsync(item.Id);
                    var highestPrice = itemPrices.Buys.UnitPrice;
                    var price = currentBuyingList.Where(x => x.ItemId == item.Id).Max(x => x.Price);

                    // If we haven't been undercut
                    if (highestPrice <= price)
                    {
                        Logger.Debug($"{item.Name} Not undercut");
                        continue;
                    }
                }
                catch (Exception e)
                {
                    Logger.Error(e);
                    continue;
                }
            }
            else
            {
                Logger.Debug($"{item.Name} Not buying");
                continue;
            }

            Logger.Info($"Cancelling [{item.Id}] {item.Name}");

            await CancelItem(item);
        }
    }

    private static void CheckNewMap()
    {
        Logger.Info("Checking if new map available");
        Thread.Sleep(2000);
        var shroomPoint = ImageSearch.FindImageInFullWindow(process!, Resources.NewMapShroom);

        if (shroomPoint != null)
        {
            Logger.Info($"Found shroom at {shroomPoint}");
            var yesPoint = ImageSearch.FindImageInWindow(process!, Resources.NewMapYes, shroomPoint.Value.X - 118, shroomPoint.Value.Y + 42, Resources.NewMapYes.Width, Resources.NewMapYes.Height, 0.8);

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
        if (DateTime.Now - timeSinceLastAfkCheck < TimeSpan.FromMinutes(20))
        {
            return;
        }

        Logger.Info("Moving for anti-afk");
        Input.KeyPress(process!, VirtualKeyCode.VK_S);
        Input.KeyPress(process!, VirtualKeyCode.VK_W);
        timeSinceLastAfkCheck = DateTime.Now;
    }

    private static void CheckPosition()
    {
        using var apiClient = new Gw2Client(ApiConnection);

        apiClient.Mumble.Update();

        var position = apiClient.Mumble.AvatarPosition;
    }

    private static int ToCurrency(int gold, int silver, int copper) => (gold * 10000) + (silver * 100) + copper;

    private static int ToGold(int coins) => coins / 10000;

    private static int ToSilver(int coins) => (coins - (ToGold(coins) * 10000)) / 100;

    private static int ToCopper(int coins) => coins - (ToSilver(coins) * 100) - (ToGold(coins) * 10000);

    private static int GetProfit(int buy, int sell)
    {
        var listingFee = (int)Math.Ceiling(sell * 0.05);

        var exchangeFee = (int)Math.Ceiling(sell * 0.1);

        return sell - (buy + listingFee + exchangeFee);
    }

    private static int GetSellPriceForProfit(int buy, int profit) => (int)Math.Floor((profit + buy) / 0.85);

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

    private static bool FindSearchResultsLoaded()
    {
        var noItems = ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.8);
        var resultBorder = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, tradingPostPoint!.Value.X + 322, tradingPostPoint!.Value.Y + 177, Resources.ResultCorner.Width, Resources.ResultCorner.Height + 72);
        return noItems == null && resultBorder == null;
    }

    private static bool FindFilter()
    {
        var find = ImageSearch.FindImageInWindow(process!, Resources.FilterOpen, tradingPostPoint!.Value.X + 50, tradingPostPoint!.Value.Y + 182, Resources.FilterOpen.Width, Resources.FilterOpen.Height, 0.9);
        return find == null;
    }

    private static bool FindItemBorder()
    {
        var find = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, tradingPostPoint!.Value.X + 271, tradingPostPoint!.Value.Y + 136, Resources.ResultCorner.Width, Resources.ResultCorner.Height);
        return find == null;
    }

    private static bool FindGoldCoin()
    {
        var find = ImageSearch.FindImageInWindow(process!, Resources.GoldCoin, tradingPostPoint!.Value.X + 462, tradingPostPoint!.Value.Y + 280, Resources.GoldCoin.Width, Resources.GoldCoin.Height + 24, 0.9);
        return find == null;
    }

    private static bool FindVendorE()
    {
        var find = ImageSearch.FindImageInWindow(process!, Resources.VendorE, tradingPostPoint!.Value.X + 345, tradingPostPoint!.Value.Y + 176, Resources.VendorE.Width, Resources.VendorE.Height + 30, 0.9);
        return find == null;
    }

    private static bool FindAvailable()
    {
        var find1 = ImageSearch.FindImageInWindow(process!, Resources.Available, tradingPostPoint!.Value.X + 531, tradingPostPoint!.Value.Y + 469, Resources.Available.Width, Resources.Available.Height, 0.9);
        var find2 = ImageSearch.FindImageInWindow(process!, Resources.Available, tradingPostPoint!.Value.X + 531, tradingPostPoint!.Value.Y + 493, Resources.Available.Width, Resources.Available.Height, 0.9);
        return find1 == null && find2 == null;
    }

    private static void DismissSuccess()
    {
        if (ImageSearch.FindImageInWindow(process!, Resources.SuccessOK, tradingPostPoint!.Value.X + 392, tradingPostPoint!.Value.Y + 365, Resources.SuccessOK.Width, Resources.SuccessOK.Height) != null)
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 450, tradingPostPoint!.Value.Y + 374);
            Thread.Sleep(500);
        }
        else if (ImageSearch.FindImageInWindow(process!, Resources.SuccessOK, tradingPostPoint!.Value.X + 392, tradingPostPoint!.Value.Y + 365 + 24, Resources.SuccessOK.Width, Resources.SuccessOK.Height) != null)
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 450, tradingPostPoint!.Value.Y + 374 + 24);
            Thread.Sleep(500);
        }
    }

    private static void DismissSuccessWithOffset(int offset)
    {
        if (ImageSearch.FindImageInWindow(process!, Resources.SuccessOK, tradingPostPoint!.Value.X + 392, tradingPostPoint!.Value.Y + 365 + offset, Resources.SuccessOK.Width, Resources.SuccessOK.Height) != null)
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 450, tradingPostPoint!.Value.Y + 374 + offset);
            Thread.Sleep(500);
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
