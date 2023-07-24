namespace GW2Flipper;

using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

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
    private const bool Debug = false;

    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
    private static readonly Config Config = Config.Instance!;

    private static readonly List<BuyItem> BuyItemsList = new();
    private static readonly List<BuyItem> RemoveItemsList = new();
    private static readonly Connection ApiConnection = new(Config.ApiKey);

    private static readonly Dictionary<string, Color> RarityColors = new()
    {
        { "Junk", Color.FromArgb(170, 170, 170) },
        { "Basic", Color.FromArgb(255, 255, 255) },
        { "Fine", Color.FromArgb(79, 157, 254) },
        { "Masterwork", Color.FromArgb(45, 197, 14) },
        { "Rare", Color.FromArgb(250, 223, 31) },
        { "Exotic", Color.FromArgb(238, 155, 3) },
        { "Ascended", Color.FromArgb(233, 67, 125) },
        { "Legendary", Color.FromArgb(160, 46, 247) },
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

    private static Process? process;
    private static Point? tradingPostPoint;
    private static DateTime timeSinceLastAfkCheck = DateTime.MinValue;
    private static DateTime timeSinceLastNewMapCheck = DateTime.MinValue;
    private static DateTime timeSinceLastBuyListGenerate = DateTime.MinValue;
    private static DateTime timeSinceLastSell = DateTime.MinValue;
    private static DateTime timeSinceLastBuy = DateTime.MinValue;
    private static int currentLoopIndex;

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

        foreach (var cefProcess in Process.GetProcessesByName("CefHost").Where(p => p.SessionId == Process.GetCurrentProcess().SessionId))
        {
            cefProcess.PriorityClass = ProcessPriorityClass.High;
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

        while (true)
        {
            if (Array.Find(Process.GetProcessesByName("Gw2-64"), p => p.SessionId == Process.GetCurrentProcess().SessionId) == null)
            {
                Logger.Error("Process gone");
                return;
            }

            tradingPostPoint = await GetTradingPostPoint();
            if (tradingPostPoint == null)
            {
                Logger.Error("Couldn't open trading post");
                return;
            }

            try
            {
                if (ResetUI())
                {
                    AntiAfk();
                    CheckNewMap();
                }
            }
            catch (Exception e)
            {
                Logger.Error(e);
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

                if (Config.RemoveUnprofitable && RemoveItemsList.Count > 0)
                {
                    try
                    {
                        await CancelUnprofitable();
                    }
                    catch (Exception e)
                    {
                        Logger.Error(e);
                    }

                    RemoveItemsList.Clear();
                    _ = ResetUI();
                }
            }

            try
            {
                var howLongLastBuy = DateTime.Now - timeSinceLastBuy;
                if (howLongLastBuy < TimeSpan.FromMinutes(6))
                {
                    _ = ResetUI();
                    var howLongWait = TimeSpan.FromMinutes(6) - howLongLastBuy;
                    Logger.Info($"Sleeping for {howLongWait.TotalMinutes} minutes");
                    await Task.Delay(howLongWait);
                }

                timeSinceLastBuy = DateTime.Now;
                await BuyItems();
            }
            catch (Exception e)
            {
                Logger.Error("BuyItems exception");
                Logger.Error(e);
            }

            try
            {
                var howLongLastSell = DateTime.Now - timeSinceLastSell;
                if (howLongLastSell < TimeSpan.FromMinutes(6))
                {
                    _ = ResetUI();
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
        }
    }

    public static async Task RunCancelAndSell()
    {
        process = Array.Find(Process.GetProcessesByName("Gw2-64"), p => p.SessionId == Process.GetCurrentProcess().SessionId);
        if (process == null)
        {
            Logger.Error("Couldn't find process");
            return;
        }

        Logger.Info($"Found process: [{process!.Id}] {process!.MainWindowTitle}");

        process.PriorityClass = ProcessPriorityClass.High;

        foreach (var cefProcess in Process.GetProcessesByName("CefHost").Where(p => p.SessionId == Process.GetCurrentProcess().SessionId))
        {
            cefProcess.PriorityClass = ProcessPriorityClass.High;
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

        _ = ResetUI();

        try
        {
            await UpdateBuyList();
        }
        catch (Exception e)
        {
            Logger.Error(e);
        }

        try
        {
            await CancelAllBuying();
        }
        catch (Exception e)
        {
            Logger.Error(e);
        }

        try
        {
            await SellItems();
        }
        catch (Exception e)
        {
            Logger.Error(e);
        }
    }

    public static async Task RunSellAll()
    {
        process = Array.Find(Process.GetProcessesByName("Gw2-64"), p => p.SessionId == Process.GetCurrentProcess().SessionId);
        if (process == null)
        {
            Logger.Error("Couldn't find process");
            return;
        }

        Logger.Info($"Found process: [{process!.Id}] {process!.MainWindowTitle}");

        process.PriorityClass = ProcessPriorityClass.High;

        foreach (var cefProcess in Process.GetProcessesByName("CoherentUI_Host").Where(p => p.SessionId == Process.GetCurrentProcess().SessionId))
        {
            cefProcess.PriorityClass = ProcessPriorityClass.High;
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

        try
        {
            await CancelAllBuying();
        }
        catch (Exception e)
        {
            Logger.Error(e);
        }
    }

    public static async Task RunImageSave()
    {
        process = Array.Find(Process.GetProcessesByName("Gw2-64"), p => p.SessionId == Process.GetCurrentProcess().SessionId);
        if (process == null)
        {
            Logger.Error("Couldn't find process");
            return;
        }

        Logger.Info($"Found process: [{process!.Id}] {process!.MainWindowTitle}");

        process.PriorityClass = ProcessPriorityClass.High;

        foreach (var cefProcess in Process.GetProcessesByName("CefHost").Where(p => p.SessionId == Process.GetCurrentProcess().SessionId))
        {
            cefProcess.PriorityClass = ProcessPriorityClass.High;
        }

        _ = User32.MoveWindow(process.MainWindowHandle, 0, 5, 1080, 850, true);
        _ = User32.MoveWindow(Process.GetCurrentProcess().MainWindowHandle, 0, 855, 1080, 360, true);

        Input.EnsureForegroundWindow(process);

        LogFullImage();

        /*// Play
        Input.MouseMove(process!, 540, 0);
        await Task.Delay(2500);
        LogSpecificImage("Play", 861, 782, 25, 10);*/

        /*// NewMapShroom
        LogSpecificImage("NewMapShroom", 148, 418, 24, 24);

        // NewMapYes
        LogSpecificImage("NewMapYes", 30, 460, 52, 19);

        return;*/

        tradingPostPoint = ImageSearch.FindImageInFullWindow(process!, Resources.TradingPostLion);

        if (tradingPostPoint == null)
        {
            Logger.Error("Couldn't open trading post");
            return;
        }

        Logger.Info($"Found trading post at: {tradingPostPoint}");

        LogTPImage();

        // SuccessOK
        // LogSpecificImage("SuccessOK", tradingPostPoint!.Value.X + 392, tradingPostPoint!.Value.Y + 365, 116, 18);

        // NewDelivery
        // LogSpecificImage("NewDelivery", tradingPostPoint!.Value.X + 70, tradingPostPoint!.Value.Y + 445, 14, 14);

        // Buying
        // LogSpecificImage("Buying", tradingPostPoint!.Value.X + 672, tradingPostPoint!.Value.Y + 238, 40, 7);

        // return;

        // Home
        LogSpecificImage("Home", tradingPostPoint!.Value.X + 377, tradingPostPoint!.Value.Y + 58, 44, 18);

        // FilterOpen
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 287, tradingPostPoint!.Value.Y + 165);
        Input.MouseMove(process!, 540, 0);
        await Task.Delay(2500);
        LogSpecificImage("FilterOpen", tradingPostPoint!.Value.X + 50, tradingPostPoint!.Value.Y + 182, 3, 3);

        // BuyItemsActive
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 566, tradingPostPoint!.Value.Y + 90);
        Input.MouseMove(process!, 540, 0);
        await Task.Delay(2500);
        LogSpecificImage("BuyItemsActive", tradingPostPoint!.Value.X + 638, tradingPostPoint!.Value.Y + 130, 10, 10);

        // Qty
        LogSpecificImage("Qty", tradingPostPoint!.Value.X + 324, tradingPostPoint!.Value.Y + 191, 20, 10);

        // ResultCorner
        LogSpecificImage("ResultCorner", tradingPostPoint!.Value.X + 322, tradingPostPoint!.Value.Y + 213, 2, 2);

        // VendorE
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 350, tradingPostPoint!.Value.Y + 240);
        Input.MouseMove(process!, 540, 0);
        await Task.Delay(2500);
        LogSpecificImage("VendorE", tradingPostPoint!.Value.X + 345, tradingPostPoint!.Value.Y + 176, 6, 2);

        // Available
        LogSpecificImage("Available", tradingPostPoint!.Value.X + 531, tradingPostPoint!.Value.Y + 469, 15, 10);

        // BuyInstantly
        LogSpecificImage("BuyInstantly", tradingPostPoint!.Value.X + 382, tradingPostPoint!.Value.Y + 364, 116, 18);

        // PlaceOrder
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 254, tradingPostPoint!.Value.Y + 499);
        Input.MouseMove(process!, 540, 0);
        await Task.Delay(2500);
        LogSpecificImage("PlaceOrder", tradingPostPoint!.Value.X + 382, tradingPostPoint!.Value.Y + 364, 116, 18);

        // Exit
        LogSpecificImage("Exit", tradingPostPoint!.Value.X + 769, tradingPostPoint!.Value.Y + 117, 16, 16);
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 777, tradingPostPoint!.Value.Y + 125);
        await Task.Delay(2500);

        // SellItemsActive
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 732, tradingPostPoint!.Value.Y + 90);
        Input.MouseMove(process!, 540, 0);
        await Task.Delay(2500);
        LogSpecificImage("SellItemsActive", tradingPostPoint!.Value.X + 805, tradingPostPoint!.Value.Y + 130, 10, 10);

        // SellInstantly
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 350, tradingPostPoint!.Value.Y + 214);
        Input.MouseMove(process!, 540, 0);
        await Task.Delay(2500);
        LogSpecificImage("SellInstantly", tradingPostPoint!.Value.X + 382, tradingPostPoint!.Value.Y + 364, 116, 18);

        // ListItem
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 542, tradingPostPoint!.Value.Y + 499);
        Input.MouseMove(process!, 540, 0);
        await Task.Delay(2500);
        LogSpecificImage("ListItem", tradingPostPoint!.Value.X + 382, tradingPostPoint!.Value.Y + 364, 116, 18);
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 777, tradingPostPoint!.Value.Y + 125);
        await Task.Delay(2500);

        // TransactionsActive
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 899, tradingPostPoint!.Value.Y + 90);
        Input.MouseMove(process!, 540, 0);
        await Task.Delay(2500);
        LogSpecificImage("TransactionsActive", tradingPostPoint!.Value.X + 817, tradingPostPoint!.Value.Y + 130, 10, 10);

        // CurrentBuying
        LogSpecificImage("CurrentBuying", tradingPostPoint!.Value.X + 60, tradingPostPoint!.Value.Y + 224, 34, 24);

        // LoadMore
        LogSpecificImage("LoadMore", tradingPostPoint!.Value.X + 610, tradingPostPoint!.Value.Y + 679, 80, 18);

        // LoadMoreLight
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 650, tradingPostPoint!.Value.Y + 688);
        Input.MouseMove(process!, 540, 0);
        await Task.Delay(2500);
        LogSpecificImage("LoadMoreLight", tradingPostPoint!.Value.X + 610, tradingPostPoint!.Value.Y + 679, 80, 18);

        // NoItems
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 174, tradingPostPoint!.Value.Y + 240);
        Input.MouseMove(process!, 540, 0);
        await Task.Delay(2500);
        LogSpecificImage("NoItems", tradingPostPoint!.Value.X + 494, tradingPostPoint!.Value.Y + 412, 9, 15);
    }

    private static async Task UpdateBuyList()
    {
        if (DateTime.Now - timeSinceLastBuyListGenerate < TimeSpan.FromMinutes(Config.UpdateListTime))
        {
            return;
        }

        Logger.Info("Updating buy list");

        timeSinceLastBuyListGenerate = DateTime.Now;

        List<BuyItem> newItemsList = new();

        foreach (var arguments in Config.Arguments)
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

                    if (Config.IgnoreSellsLessThanBuys && (sold * Config.SellsLessThanBuysRange) < bought)
                    {
                        continue;
                    }

                    if (Config.Blacklist.Contains(itemId))
                    {
                        Logger.Debug($"Removing blacklisted item: [{itemId}] {itemName}");
                        continue;
                    }

                    var newItem = new BuyItem(itemId, itemName, buyPrice, sellPrice, profit, sold, bought);

                    if (Config.RemoveUnprofitable)
                    {
                        var findIndex = newItemsList.FindIndex(x => x.Id == itemId);
                        if (findIndex == -1)
                        {
                            newItemsList.Add(newItem);
                        }
                    }
                    else
                    {
                        // Add or update here
                        var findIndex = BuyItemsList.FindIndex(x => x.Id == itemId);
                        if (findIndex == -1)
                        {
                            BuyItemsList.Add(newItem);
                        }
                        else
                        {
                            BuyItemsList[findIndex] = newItem;
                        }
                    }
                }
            }
        }

        if (Config.RemoveUnprofitable)
        {
            foreach (var item in BuyItemsList.ToList())
            {
                if (newItemsList.FindIndex(x => x.Id == item.Id) == -1)
                {
                    RemoveItemsList.Add(item);
                }
            }

            BuyItemsList.Clear();
            BuyItemsList.AddRange(newItemsList);
        }

        BuyItemsList.Sort((x, y) => y.Profit.CompareTo(x.Profit));
        // BuyItemsList.Shuffle();

        Logger.Info($"Items: {BuyItemsList.Count}");
    }

    private static async Task UpdateBuyList2()
    {
        if (DateTime.Now - timeSinceLastBuyListGenerate < TimeSpan.FromMinutes(Config.UpdateListTime))
        {
            return;
        }

        Logger.Info("Updating buy list");

        timeSinceLastBuyListGenerate = DateTime.Now;

        List<string> urls = new()
        {
            "http://gw2profits.com/spread.php?order=buy_combo2&price_choice=current&search_name=&min_coin=500&min_coin_day=0&min_percent=10&min_percent_day=0&velocity_min=200",
            "http://gw2profits.com/spread.php?order=buy_combo2&price_choice=current&search_name=&min_coin=100&min_coin_day=0&min_percent=10&min_percent_day=0&velocity_min=1000",
            "http://gw2profits.com/spread.php?order=buy_combo2&price_choice=current&search_name=&min_coin=10&min_coin_day=0&min_percent=10&min_percent_day=0&velocity_min=4000",
        };

        List<BuyItem> newItemsList = new();

        foreach (var url in urls)
        {
            HtmlWeb web = new();

            var htmlDoc = await web.LoadFromWebAsync(url);

            var itemTable = htmlDoc.DocumentNode.SelectSingleNode("//table[contains(@class, 'lessfade')]");
            if (itemTable == null)
            {
                break;
            }

            foreach (var item in itemTable.SelectNodes(".//tr[td]"))
            {
                var itemLink = item.SelectSingleNode("td[2]/a").Attributes["href"].Value;
                var itemIdMatch = Regex.Match(itemLink, @"items.php\?iid=(\d+)").Groups[1].Value;
                var itemId = int.Parse(itemIdMatch);
                var itemName = HtmlEntity.DeEntitize(item.SelectSingleNode("td[2]/a").InnerText);

                var pricesText = item.SelectSingleNode("td[3]").InnerHtml;
                var pricesMatch = Regex.Match(pricesText, @"(?:(\d+)<img src=""images\/Gold_coin.png"">)?(?:(\d+)<img src=""images\/Silver_coin.png"">)?(\d+)<img src=""images\/Copper_coin.png""><br><br>(?:(\d+)<img src=""images\/Gold_coin.png"">)?(?:(\d+)<img src=""images\/Silver_coin.png"">)?(\d+)<img src=""images\/Copper_coin.png"">");
                var buyGoldMatch = pricesMatch.Groups[1].Value;
                buyGoldMatch = string.IsNullOrEmpty(buyGoldMatch) ? "0" : buyGoldMatch;
                var buySilverMatch = pricesMatch.Groups[2].Value;
                buySilverMatch = string.IsNullOrEmpty(buySilverMatch) ? "0" : buySilverMatch;
                var buyCopperMatch = pricesMatch.Groups[3].Value;
                buyCopperMatch = string.IsNullOrEmpty(buyCopperMatch) ? "0" : buyCopperMatch;
                var buyPrice = ToCurrency(int.Parse(buyGoldMatch), int.Parse(buySilverMatch), int.Parse(buyCopperMatch));
                var sellGoldMatch = pricesMatch.Groups[4].Value;
                sellGoldMatch = string.IsNullOrEmpty(sellGoldMatch) ? "0" : sellGoldMatch;
                var sellSilverMatch = pricesMatch.Groups[5].Value;
                sellSilverMatch = string.IsNullOrEmpty(sellSilverMatch) ? "0" : sellSilverMatch;
                var sellCopperMatch = pricesMatch.Groups[6].Value;
                sellCopperMatch = string.IsNullOrEmpty(sellCopperMatch) ? "0" : sellCopperMatch;
                var sellPrice = ToCurrency(int.Parse(sellGoldMatch), int.Parse(sellSilverMatch), int.Parse(sellCopperMatch));

                var profitText = item.SelectSingleNode("td[4]").InnerText;
                var profitMatch = Regex.Match(profitText, @"(\d+) \(\d+%\)").Groups[1].Value;
                var profit = int.Parse(profitMatch);

                var soldText = item.SelectSingleNode("td[2]").InnerHtml;
                var soldMatch = Regex.Match(soldText, @"<br>([\d,.]+)[-+] \/ day").Groups[1].Value;
                var sold = (int)decimal.Parse(soldMatch);

                if (Config.Blacklist.Contains(itemId))
                {
                    Logger.Debug($"Removing blacklisted item: [{itemId}] {itemName}");
                    continue;
                }

                var newItem = new BuyItem(itemId, itemName, buyPrice, sellPrice, profit, sold, 0);

                if (Config.RemoveUnprofitable)
                {
                    var findIndex = newItemsList.FindIndex(x => x.Id == itemId);
                    if (findIndex == -1)
                    {
                        newItemsList.Add(newItem);
                    }
                }
                else
                {
                    // Add or update here
                    var findIndex = BuyItemsList.FindIndex(x => x.Id == itemId);
                    if (findIndex == -1)
                    {
                        BuyItemsList.Add(newItem);
                    }
                    else
                    {
                        BuyItemsList[findIndex] = newItem;
                    }
                }
            }
        }

        if (Config.RemoveUnprofitable)
        {
            foreach (var item in BuyItemsList.ToList())
            {
                if (newItemsList.FindIndex(x => x.Id == item.Id) == -1)
                {
                    RemoveItemsList.Add(item);
                }
            }

            BuyItemsList.Clear();
            BuyItemsList.AddRange(newItemsList);
        }

        // BuyItemsList.Sort((x, y) => y.Profit.CompareTo(x.Profit));
        BuyItemsList.Shuffle();

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
                PlayCharacterCheck();

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
        var lionPoint = ImageSearch.FindImageInWindow(process!, Resources.TradingPostLion, tradingPostPoint!.Value.X, tradingPostPoint!.Value.Y, Resources.TradingPostLion.Width, Resources.TradingPostLion.Height);
        var homePoint = ImageSearch.FindImageInWindow(process!, Resources.Home, tradingPostPoint!.Value.X + 377, tradingPostPoint!.Value.Y + 58, Resources.Home.Width, Resources.Home.Height, 0.9);

        if (lionPoint == null || homePoint == null)
        {
            Logger.Info("Lost sight of trading post");
            Input.KeyPress(process!, VirtualKeyCode.VK_F);

            try
            {
                await WaitWhile(() => ImageSearch.FindImageInFullWindow(process!, Resources.Home, 0.9) == null, 500, 5000);
            }
            catch (TimeoutException)
            {
                if (ResetUI())
                {
                    Input.KeyPress(process!, VirtualKeyCode.VK_F);

                    try
                    {
                        await WaitWhile(() => ImageSearch.FindImageInFullWindow(process!, Resources.Home, 0.9) == null, 500, 5000);
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

            Input.KeyPress(process!, VirtualKeyCode.VK_M);

            await Task.Delay(2000);
        }
    }

    private static bool ResetUI()
    {
        Logger.Info("Attempting reset UI");
        Input.MouseMove(process!, 540, 0);

        for (var i = 0; i < 10; i++)
        {
            if (ImageSearch.FindImageInWindow(process!, Resources.ErrorOK, 661, 465, Resources.ErrorOK.Width, Resources.ErrorOK.Height, 0.9) != null)
            {
                Logger.Info("Found error message box");
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, 671, 472);
                Thread.Sleep(500);
            }

            Input.KeyPress(process!, VirtualKeyCode.ESCAPE);
            Thread.Sleep(500);

            if (ImageSearch.FindImageInFullWindow(process!, Resources.ReturnToGame, 0.9) != null)
            {
                Logger.Info("Found escape menu");
                Input.KeyPress(process!, VirtualKeyCode.ESCAPE);
                Thread.Sleep(500);
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
        TakeNewDelivery();

        Func<bool>? findFunc = null;
        switch (screen)
        {
            case TradingPostScreen.Sell:
                // Check if already open
                if (ImageSearch.FindImageInWindow(process!, Resources.SellItemsActive, tradingPostPoint!.Value.X + 805, tradingPostPoint!.Value.Y + 130, Resources.SellItemsActive.Width, Resources.SellItemsActive.Height, 0.9) != null)
                {
                    return;
                }

                // Click sell items
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 732, tradingPostPoint!.Value.Y + 90);

                static bool FindSellItemsActive()
                {
                    var find = ImageSearch.FindImageInWindow(process!, Resources.SellItemsActive, tradingPostPoint!.Value.X + 805, tradingPostPoint!.Value.Y + 130, Resources.SellItemsActive.Width, Resources.SellItemsActive.Height, 0.9);
                    return find == null;
                }

                findFunc = FindSellItemsActive;
                break;
            case TradingPostScreen.Buy:
                // Check if already open
                if (ImageSearch.FindImageInWindow(process!, Resources.BuyItemsActive, tradingPostPoint!.Value.X + 638, tradingPostPoint!.Value.Y + 130, Resources.BuyItemsActive.Width, Resources.BuyItemsActive.Height, 0.9) != null)
                {
                    return;
                }

                // Click buy items
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 566, tradingPostPoint!.Value.Y + 90);

                static bool FindBuyItemsActive()
                {
                    var find = ImageSearch.FindImageInWindow(process!, Resources.BuyItemsActive, tradingPostPoint!.Value.X + 638, tradingPostPoint!.Value.Y + 130, Resources.BuyItemsActive.Width, Resources.BuyItemsActive.Height, 0.9);
                    return find == null;
                }

                findFunc = FindBuyItemsActive;
                break;
            case TradingPostScreen.Transactions:
                // Check if already open
                if (ImageSearch.FindImageInWindow(process!, Resources.TransactionsActive, tradingPostPoint!.Value.X + 817, tradingPostPoint!.Value.Y + 130, Resources.TransactionsActive.Width, Resources.TransactionsActive.Height, 0.9) != null)
                {
                    return;
                }

                // Click transactions
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 899, tradingPostPoint!.Value.Y + 90);

                static bool FindTransactionsActive()
                {
                    var find = ImageSearch.FindImageInWindow(process!, Resources.TransactionsActive, tradingPostPoint!.Value.X + 817, tradingPostPoint!.Value.Y + 130, Resources.TransactionsActive.Width, Resources.TransactionsActive.Height, 0.9);
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
            PlayCharacterCheck();
            throw new TimeoutException();
        }

        await Task.Delay(1000);
    }

    private static void TakeNewDelivery()
    {
        if (ImageSearch.FindImageInWindow(process!, Resources.NewDelivery, tradingPostPoint!.Value.X + 70, tradingPostPoint!.Value.Y + 445, Resources.NewDelivery.Width, Resources.NewDelivery.Height, 0.9) != null)
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 179, tradingPostPoint!.Value.Y + 685);
        }
    }

    private static async Task CloseItemWindow()
    {
        Input.MouseMove(process!, tradingPostPoint!.Value.X + 777, tradingPostPoint!.Value.Y + 140);
        await Task.Delay(100);
        if (ImageSearch.FindImageInWindow(process!, Resources.Exit, tradingPostPoint!.Value.X + 769, tradingPostPoint!.Value.Y + 117, Resources.Exit.Width, Resources.Exit.Height, 0.5) != null)
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 777, tradingPostPoint!.Value.Y + 125);
        }

        await Task.Delay(500);
    }

    /* private static async Task SearchForItem(Gw2Sharp.WebApi.V2.Models.Item itemInfo, TradingPostScreen screen)
    {
        await Task.Delay(500);

        // Click search bar
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 163, tradingPostPoint!.Value.Y + 165);

        if (screen == TradingPostScreen.Buy)
        {
            // Check if filter isn't open
            if (FindFilter())
            {
                await Task.Delay(200);

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

        if (ImageSearch.FindImageInWindow(process!, Resources.TradingPostLion, tradingPostPoint!.Value.X, tradingPostPoint!.Value.Y, Resources.TradingPostLion.Width, Resources.TradingPostLion.Height) != null)
        {
            if (screen != TradingPostScreen.Buy)
            {
                // Select all
                Input.KeyPressWithModifier(process!, VirtualKeyCode.VK_A, false, true, false);

                // Backspace
                Input.KeyPress(process!, VirtualKeyCode.BACK);
            }

            // Paste in name
            // Input.KeyStringSend(process!, itemInfo.Name.MaxSize(30));
            Input.KeyStringSendClipboard(process!, itemInfo.Name.MaxSize(30));
            Input.KeyPress(process!, VirtualKeyCode.RETURN);
        }

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

        await Task.Delay(1000);
    }*/

    private static async Task SearchForItem(Gw2Sharp.WebApi.V2.Models.Item itemInfo, TradingPostScreen screen)
    {
        await Task.Delay(500);

        // Click search bar
        // Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 163, tradingPostPoint!.Value.Y + 165);

        // Check if filter isn't open
        if (FindFilter())
        {
            await Task.Delay(200);

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
            Thread.Sleep(200);
        }
        else
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 173, tradingPostPoint!.Value.Y + 309);
        }

        // Click search bar
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 163, tradingPostPoint!.Value.Y + 165);

        if (ImageSearch.FindImageInWindow(process!, Resources.TradingPostLion, tradingPostPoint!.Value.X, tradingPostPoint!.Value.Y, Resources.TradingPostLion.Width, Resources.TradingPostLion.Height) != null)
        {
            // Paste in name
            // Input.KeyStringSend(process!, itemInfo.Name.MaxSize(30));
            Input.KeyStringSendClipboard(process!, itemInfo.Name.MaxSize(30));
            Input.KeyPress(process!, VirtualKeyCode.RETURN);
        }

        if (screen == TradingPostScreen.Buy)
        {
            // Set rarity
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 173, tradingPostPoint!.Value.Y + 205);
            await Task.Delay(500);
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 173, tradingPostPoint!.Value.Y + 230 + (25 * RarityOrder[itemInfo.Rarity!]));
            await Task.Delay(500);
        }

        // Input level
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 63, tradingPostPoint!.Value.Y + 270);
        Input.KeyStringSend(process!, itemInfo.Level.ToString());
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 124, tradingPostPoint!.Value.Y + 270);
        Input.KeyStringSend(process!, itemInfo.Level.ToString());

        // Click filter cog to close
        Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 287, tradingPostPoint!.Value.Y + 165);

        await Task.Delay(1000);
    }

    private static int GetItemWindowPrice(int offset)
    {
        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 412, tradingPostPoint.Value.Y + 279 + offset, 100);
        var goldAmount = Convert.ToInt32(Input.GetSelectedText(process!));

        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 502, tradingPostPoint.Value.Y + 279 + offset, 100);
        var silverAmount = Convert.ToInt32(Input.GetSelectedText(process!));

        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 576, tradingPostPoint.Value.Y + 279 + offset, 100);
        var copperAmount = Convert.ToInt32(Input.GetSelectedText(process!));

        return ToCurrency(goldAmount, silverAmount, copperAmount);
    }

    private static void SetItemWindowPrice(int offset, int coins)
    {
        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 412, tradingPostPoint.Value.Y + 279 + offset, 100);
        Input.KeyStringSend(process!, ToGold(coins).ToString());

        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 502, tradingPostPoint.Value.Y + 279 + offset, 100);
        Input.KeyStringSend(process!, ToSilver(coins).ToString());

        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 576, tradingPostPoint.Value.Y + 279 + offset, 100);
        Input.KeyStringSend(process!, ToCopper(coins).ToString());
    }

    private static int GetItemWindowQuantity(int offset)
    {
        DismissSuccessWithOffset(offset);
        Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 402, tradingPostPoint.Value.Y + 236 + offset);
        DismissSuccessWithOffset(offset);

        return Convert.ToInt32(Input.GetSelectedText(process!));
    }

    private static async Task SellItem(BuyItem item, List<Gw2Sharp.WebApi.V2.Models.CommerceTransactionCurrent>? currentSellingList = null)
    {
        using var apiClient = new Gw2Client(ApiConnection);

        var restartAttempts = 0;
        RestartSell:
        if (restartAttempts >= 5)
        {
            return;
        }

        restartAttempts++;

        // Get item info
        Gw2Sharp.WebApi.V2.Models.Item? itemInfo = null;
        try
        {
            itemInfo = await apiClient.WebApi.V2.Items.GetAsync(item.Id);
        }
        catch (JsonException)
        {
            return;
        }
        catch (Exception e)
        {
            Logger.Error(e);
            if (!ResetUI())
            {
                return;
            }

            goto RestartSell;
        }

        if (itemInfo == null)
        {
            Logger.Warn("itemInfo null");
            return;
        }

        // Get item prices
        Gw2Sharp.WebApi.V2.Models.CommercePrices? itemPrices = null;
        try
        {
            itemPrices = await apiClient.WebApi.V2.Commerce.Prices.GetAsync(item.Id);
        }
        catch (Exception e)
        {
            Logger.Error(e);
            if (!ResetUI())
            {
                return;
            }

            goto RestartSell;
        }

        if (itemPrices == null)
        {
            Logger.Warn("itemPrices null");
            return;
        }

        var buyCurrencyAmount = itemPrices.Buys.UnitPrice;
        var sellCurrencyAmount = itemPrices.Sells.UnitPrice;

        Logger.Info($"Highest buyer: {buyCurrencyAmount}");
        Logger.Info($"Lowest seller: {sellCurrencyAmount}");

        // Check if sell price is within range
        /*if (sellCurrencyAmount < item.SellPrice * ErrorRange)
        {
            Logger.Info($"Sell price {sellCurrencyAmount} below range {item.SellPrice * ErrorRange}");
            await CloseItemWindow();
            break;
        }*/

        try
        {
            await GoToScreen(TradingPostScreen.Sell);
        }
        catch (TimeoutException)
        {
            Logger.Debug("Timeout on GoToScreen");
            if (!ResetUI())
            {
                return;
            }

            goto RestartSell;
        }

        // Search
        try
        {
            await SearchForItem(itemInfo, TradingPostScreen.Sell);
        }
        catch (TimeoutException)
        {
            Logger.Warn("Filter didn't appear");
            if (!ResetUI())
            {
                return;
            }

            goto RestartSell;
        }

        // Wait for search results to load
        try
        {
            await WaitWhile(FindSearchResultsLoaded, 500, 5000);
        }
        catch (TimeoutException)
        {
            Logger.Error("Didn't load any search results");
            if (!ResetUI())
            {
                return;
            }

            goto RestartSell;
        }

        // Check if no item found
        if (ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.8) != null)
        {
            Logger.Debug("No item found");
            return;
        }

        // Find position of Qty to determine position of search results
        var qtyPoint = ImageSearch.FindImageInWindow(process!, Resources.Qty, tradingPostPoint!.Value.X + 324, tradingPostPoint!.Value.Y + 155, Resources.Qty.Width, Resources.Qty.Height + 72, 0.9);
        if (qtyPoint == null)
        {
            Logger.Debug("qtyPoint null");
            if (!ResetUI())
            {
                return;
            }

            goto RestartSell;
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
                Logger.Debug("Result point null");
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
                Logger.Debug("Item border didn't appear");
                await CloseItemWindow();
                goto SellItem;
            }

            // Wait for "vendor" to show up
            try
            {
                await WaitWhile(FindVendorE, 500, 5000);
            }
            catch (TimeoutException)
            {
                Logger.Debug("Vendor E didn't appear");
                await CloseItemWindow();
                goto SellItem;
            }

            // Find position of 'e' in Vendor to set the offset for other UI positions
            var vendorEPos = ImageSearch.FindImageInWindow(process!, Resources.VendorE, tradingPostPoint!.Value.X + 345, tradingPostPoint!.Value.Y + 176, Resources.VendorE.Width, Resources.VendorE.Height + 31, 0.8);
            if (vendorEPos == null)
            {
                Logger.Debug("vendorEPos null");
                await CloseItemWindow();
                goto SellItem;
            }

            var offset = 0;
            var nameSize = 34;
            if (vendorEPos.Value.Y == tradingPostPoint!.Value.Y + 207)
            {
                offset = 25;
                nameSize = 64;
            }

            // Get OCR of item name and check if it matches API name
            var nameImage = ImageSearch.CaptureWindow(process!, tradingPostPoint!.Value.X + 334, tradingPostPoint.Value.Y + 136, 428, nameSize);
            var capturedName = OCR.ReadName(nameImage!, RarityColors[itemInfo.Rarity.ToString()!]);

            // Check captured name
            if (string.IsNullOrEmpty(capturedName) || !OCR.NameCompare(itemInfo.Name, capturedName, Config.OcrFixes))
            {
                Logger.Debug("Item name different");
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
                Logger.Debug("Available didn't appear");
                await CloseItemWindow();
                goto SellItem;
            }

            var sellPrice = sellCurrencyAmount;

            // Check profit is within range
            var profit = GetProfit(item.BuyPrice, sellPrice);
            var lowestProfit = (int)(item.Profit * Config.ProfitRange);
            if (profit < lowestProfit)
            {
                Logger.Info($"Profit of {profit} outside of range");

                // Find best price in range and undercut that instead
                if (Config.SellForBestIfUnderRange)
                {
                    Logger.Info($"Selling for best near profit range {lowestProfit}");

                    Gw2Sharp.WebApi.V2.Models.CommerceListings? listings = null;
                    try
                    {
                        listings = await apiClient.WebApi.V2.Commerce.Listings.GetAsync(item.Id);
                    }
                    catch (Exception e)
                    {
                        Logger.Error(e);
                        await CloseItemWindow();
                        goto SellItem;
                    }

                    if (listings == null)
                    {
                        Logger.Warn("listings null");
                        await CloseItemWindow();
                        goto SellItem;
                    }

                    var profitableSell = GetSellPriceForProfit(item.BuyPrice, lowestProfit);
                    sellPrice = listings.Sells.Where(x => x.UnitPrice >= profitableSell).Min(x => x.UnitPrice);
                }
                else
                {
                    await CloseItemWindow();
                    break;
                }
            }

            DismissSuccessWithOffset(offset);

            // Click highest current seller
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint.Value.X + 542, tradingPostPoint.Value.Y + 499 + offset);

            // Check if not my listing
            if ((Config.UndercutSells && currentSellingList == null) || (Config.UndercutSells && currentSellingList?.Any(x => x.ItemId == item.Id && x.Price == sellCurrencyAmount) == false))
            {
                sellPrice--;
            }

            SetItemWindowPrice(offset, sellPrice);

            DismissSuccessWithOffset(offset);

            // Set quantity to max
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 625, tradingPostPoint.Value.Y + 240 + offset);
            await Task.Delay(100);
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 625, tradingPostPoint.Value.Y + 240 + offset);

            Logger.Info($"Selling for {sellPrice} coins");

            DismissSuccessWithOffset(offset);

            // Sell item
            if (ImageSearch.FindImageInWindow(process!, Resources.ListItem, tradingPostPoint!.Value.X + 382, tradingPostPoint!.Value.Y + 364 + offset, Resources.ListItem.Width, Resources.ListItem.Height, 0.5) != null)
            {
                if (!Debug)
                {
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 440, tradingPostPoint.Value.Y + 373 + offset, 100);
                    await Task.Delay(100);
                }
            }
            else if (ImageSearch.FindImageInWindow(process!, Resources.SellInstantly, tradingPostPoint!.Value.X + 382, tradingPostPoint!.Value.Y + 364 + offset, Resources.SellInstantly.Width, Resources.SellInstantly.Height, 0.5) != null)
            {
                Logger.Warn("Wrong sell button");

                // Check entered price
                /*int? enteredAmount = null;
                try
                {
                    enteredAmount = GetItemWindowPrice(offset);
                }
                catch (FormatException e)
                {
                    Logger.Error(e);
                    await CloseItemWindow();
                    goto SellItem;
                }

                if (enteredAmount != sellPrice)
                {
                    Logger.Debug("Didn't enter price correctly");
                    await CloseItemWindow();
                    goto SellItem;
                }

                LogImage();
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 440, tradingPostPoint.Value.Y + 373 + offset, 100);
                await Task.Delay(100);

                y--;

                await CloseItemWindow();
                await Task.Delay(1000);
                continue;*/

                await CloseItemWindow();
                goto SellItem;
            }
            else
            {
                Logger.Debug("Couldn't find sell button");
                await CloseItemWindow();
                goto SellItem;
            }

            await CloseItemWindow();
            break;
        }
    }

    private static async Task SellItems()
    {
        Logger.Info("====================");
        Logger.Info("Selling items");

        using var apiClient = new Gw2Client(ApiConnection);

        var restartAttempts = 0;
        RestartSellItems:
        if (restartAttempts >= 5)
        {
            return;
        }

        restartAttempts++;

        Gw2Sharp.WebApi.V2.Models.CharactersInventory? backpack = null;
        try
        {
            backpack = await apiClient.WebApi.V2.Characters[Config.CharacterName].Inventory.GetAsync();
        }
        catch (Exception e)
        {
            Logger.Error(e);

            goto RestartSellItems;
        }

        if (backpack == null)
        {
            Logger.Debug("backpack null");
            return;
        }

        var inventoryBags = backpack.Bags.Where(x => x != null).ToList();
        var inventoryItems = inventoryBags.SelectMany(x => x.Inventory.Select(y => y?.Id)).ToList();

        if (!inventoryItems.Any())
        {
            Logger.Info("Inventory empty");
            return;
        }

        List<Gw2Sharp.WebApi.V2.Models.CommerceTransactionCurrent>? currentSellingList = null;
        try
        {
            var currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(0);
            currentSellingList = currentSelling.ToList();
            var numSellingPages = currentSelling.HttpResponseInfo?.PageTotal;
            for (var i = 1; i < numSellingPages; i++)
            {
                currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(i);
                numSellingPages = currentSelling.HttpResponseInfo?.PageTotal;
                currentSellingList = currentSellingList.Concat(currentSelling.ToList()).ToList();
            }
        }
        catch (Exception e)
        {
            Logger.Error(e);
            // return;

            goto RestartSellItems;
        }

        if (currentSellingList == null)
        {
            Logger.Debug("currentSellingList null");
            return;
        }

        foreach (var item in BuyItemsList)
        {
            if (!inventoryItems.Contains(item.Id))
            {
                continue;
            }

            Logger.Info("--------------------");
            Logger.Info($"Selling [{item.Id}] {item.Name}");
            Logger.Info($"Buy: {item.BuyPrice} Sell: {item.SellPrice} Profit: {item.Profit} Sold: {item.Sold} Bought: {item.Bought}");

            try
            {
                await SellItem(item, currentSellingList);
            }
            catch (Exception e)
            {
                Logger.Error(e);
                return;
            }
        }
    }

    private static async Task SellAllItems()
    {
        Logger.Info("====================");
        Logger.Info("Selling all items");

        using var apiClient = new Gw2Client(ApiConnection);

        var restartAttempts = 0;
        RestartSellAllItems:
        if (restartAttempts >= 5)
        {
            return;
        }

        restartAttempts++;

        Gw2Sharp.WebApi.V2.Models.CharactersInventory? backpack = null;
        try
        {
            backpack = await apiClient.WebApi.V2.Characters[Config.CharacterName].Inventory.GetAsync();
        }
        catch (Exception e)
        {
            Logger.Error(e);

            goto RestartSellAllItems;
        }

        if (backpack == null)
        {
            Logger.Debug("backpack null");
            return;
        }

        var inventoryBags = backpack.Bags.Where(x => x != null).ToList();
        var inventoryItems = inventoryBags.SelectMany(x => x.Inventory.Select(y => y?.Id)).ToList();

        if (!inventoryItems.Any())
        {
            Logger.Info("Inventory empty");
            return;
        }

        List<Gw2Sharp.WebApi.V2.Models.CommerceTransactionCurrent>? currentSellingList = null;
        try
        {
            var currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(0);
            currentSellingList = currentSelling.ToList();
            var numSellingPages = currentSelling.HttpResponseInfo?.PageTotal;
            for (var i = 1; i < numSellingPages; i++)
            {
                currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(i);
                numSellingPages = currentSelling.HttpResponseInfo?.PageTotal;
                currentSellingList = currentSellingList.Concat(currentSelling.ToList()).ToList();
            }
        }
        catch (Exception e)
        {
            Logger.Error(e);

            goto RestartSellAllItems;
        }

        if (currentSellingList == null)
        {
            Logger.Debug("currentSellingList null");
            return;
        }

        foreach (var item in BuyItemsList)
        {
            if (!inventoryItems.Contains(item.Id))
            {
                continue;
            }

            Logger.Info("--------------------");
            Logger.Info($"Selling [{item.Id}] {item.Name}");
            Logger.Info($"Buy: {item.BuyPrice} Sell: {item.SellPrice} Profit: {item.Profit} Sold: {item.Sold} Bought: {item.Bought}");

            await SellItem(item, currentSellingList);
        }
    }

    private static async Task BuyItem(BuyItem item)
    {
        using var apiClient = new Gw2Client(ApiConnection);

        var restartAttempts = 0;
        RestartBuy:
        if (restartAttempts >= 5)
        {
            return;
        }

        restartAttempts++;

        // Get item info
        Gw2Sharp.WebApi.V2.Models.Item? itemInfo = null;
        try
        {
            itemInfo = await apiClient.WebApi.V2.Items.GetAsync(item.Id);
        }
        catch (JsonException)
        {
            return;
        }
        catch (Exception e)
        {
            Logger.Error(e);
            if (!ResetUI())
            {
                return;
            }

            goto RestartBuy;
        }

        if (itemInfo == null)
        {
            Logger.Warn("itemInfo null");
            return;
        }

        // Get item prices
        Gw2Sharp.WebApi.V2.Models.CommercePrices? itemPrices = null;
        try
        {
            itemPrices = await apiClient.WebApi.V2.Commerce.Prices.GetAsync(item.Id);
        }
        catch (Exception e)
        {
            Logger.Error(e);
            if (!ResetUI())
            {
                return;
            }

            goto RestartBuy;
        }

        if (itemPrices == null)
        {
            Logger.Warn("itemPrices null");
            return;
        }

        var buyCurrencyAmount = itemPrices.Buys.UnitPrice;
        var sellCurrencyAmount = itemPrices.Sells.UnitPrice;

        Logger.Info($"Highest buyer: {buyCurrencyAmount}");
        Logger.Info($"Lowest seller: {sellCurrencyAmount}");

        // Check if sell price is within range
        if (sellCurrencyAmount < item.SellPrice * Config.ErrorRange)
        {
            Logger.Info("Sell price below range");
            return;
        }

        // Check if buy price is within range
        if (buyCurrencyAmount > item.BuyPrice * Config.ErrorRangeInverse)
        {
            Logger.Info("Buy price above range");
            return;
        }

        // Check profit is within range
        var profit = GetProfit(buyCurrencyAmount, sellCurrencyAmount);
        if (profit < (item.Profit * Config.ErrorRange) || profit < 5)
        {
            Logger.Info($"Profit of {profit} outside of range");
            return;
        }

        var buyPrice = buyCurrencyAmount;

        if (Config.UndercutBuys)
        {
            buyPrice++;
        }

        // Calculate quantity
        var quantity = Math.Min(Math.Min((int)(Config.MaxSpend / buyPrice), (int)(item.Sold * Config.Quantity)), 250);
        quantity = quantity < 1 ? 1 : quantity;

        // Go to buy screen
        try
        {
            await GoToScreen(TradingPostScreen.Buy);
        }
        catch (TimeoutException)
        {
            Logger.Debug("Timeout on GoToScreen");
            if (!ResetUI())
            {
                return;
            }

            goto RestartBuy;
        }

        // Search
        try
        {
            await SearchForItem(itemInfo, TradingPostScreen.Buy);
        }
        catch (TimeoutException)
        {
            Logger.Debug("Filter didn't appear");
            if (!ResetUI())
            {
                return;
            }

            goto RestartBuy;
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
            if (!ResetUI())
            {
                return;
            }

            goto RestartBuy;
        }

        // Check if no item found
        if (ImageSearch.FindImageInWindow(process!, Resources.NoItems, tradingPostPoint!.Value.X + 494, tradingPostPoint.Value.Y + 376, Resources.NoItems.Width, Resources.NoItems.Height + 72, 0.8) != null)
        {
            Logger.Debug("No item found");
            return;
        }

        // Find position of Qty to determine position of search results
        var qtyPoint = ImageSearch.FindImageInWindow(process!, Resources.Qty, tradingPostPoint!.Value.X + 324, tradingPostPoint!.Value.Y + 155, Resources.Qty.Width, Resources.Qty.Height + 72, 0.9);
        if (qtyPoint == null)
        {
            Logger.Debug("qtyPoint null");

            if (!ResetUI())
            {
                return;
            }

            goto RestartBuy;
        }

        var resultsPoint = Point.Add(qtyPoint.Value, new Size(-2, 22));

        var secondTry = false;

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
                Logger.Debug("Result point null");
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
                Logger.Debug("Item border didn't appear");
                await CloseItemWindow();
                goto BuyItem;
            }

            // Wait for "vendor" to show up
            try
            {
                await WaitWhile(FindVendorE, 500, 5000);
            }
            catch (TimeoutException)
            {
                Logger.Debug("Vendor E didn't appear");
                await CloseItemWindow();
                goto BuyItem;
            }

            // Find position of 'e' in Vendor to set the offset for other UI positions
            var vendorEPos = ImageSearch.FindImageInWindow(process!, Resources.VendorE, tradingPostPoint!.Value.X + 345, tradingPostPoint!.Value.Y + 176, Resources.VendorE.Width, Resources.VendorE.Height + 31, 0.8);
            if (vendorEPos == null)
            {
                Logger.Debug("vendorEPos null");
                await CloseItemWindow();
                goto BuyItem;
            }

            var offset = 0;
            var nameSize = 34;
            if (vendorEPos.Value.Y == tradingPostPoint!.Value.Y + 207)
            {
                offset = 25;
                nameSize = 64;
            }

            // Get OCR of item name and check if it matches API name
            var nameImage = ImageSearch.CaptureWindow(process!, tradingPostPoint!.Value.X + 334, tradingPostPoint.Value.Y + 136, 428, nameSize);
            var capturedName = OCR.ReadName(nameImage!, RarityColors[itemInfo.Rarity.ToString()!]);

            // Check captured name
            if (string.IsNullOrEmpty(capturedName) || !OCR.NameCompare(itemInfo.Name, capturedName, Config.OcrFixes))
            {
                Logger.Debug("Item name different");
                await CloseItemWindow();

                if (y == 6)
                {
                    secondTry = true;
                    y = -1;

                    Logger.Debug("Trying first item a second time");
                }

                if (secondTry && y == 0)
                {
                    return;
                }

                continue;
            }

            // Wait for available to show up
            try
            {
                await WaitWhile(FindAvailable, 500, 10000);
            }
            catch (TimeoutException)
            {
                Logger.Debug("Available didn't appear");
                await CloseItemWindow();
                goto BuyItem;
            }

            DismissSuccessWithOffset(offset);

            // Click highest current buyer
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint.Value.X + 254, tradingPostPoint.Value.Y + 499 + offset);

            SetItemWindowPrice(offset, buyPrice);

            DismissSuccessWithOffset(offset);

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

            DismissSuccessWithOffset(offset);

            // Buy item
            if (ImageSearch.FindImageInWindow(process!, Resources.PlaceOrder, tradingPostPoint!.Value.X + 382, tradingPostPoint!.Value.Y + 364 + offset, Resources.PlaceOrder.Width, Resources.PlaceOrder.Height, 0.5) != null)
            {
                Logger.Info($"Buying {quantity} for {buyPrice} coins each");

                if (!Debug)
                {
                    Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 440, tradingPostPoint.Value.Y + 373 + offset, 100);
                    await Task.Delay(100);
                }
            }
            else if (ImageSearch.FindImageInWindow(process!, Resources.BuyInstantly, tradingPostPoint!.Value.X + 382, tradingPostPoint!.Value.Y + 364 + offset, Resources.BuyInstantly.Width, Resources.BuyInstantly.Height, 0.5) != null)
            {
                Logger.Warn("Wrong buy button");

                // Click highest current seller
                /*Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint.Value.X + 542, tradingPostPoint.Value.Y + 499 + offset);
                await Task.Delay(200);

                DismissSuccessWithOffset(offset);
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint.Value.X + 542, tradingPostPoint.Value.Y + 499 + offset);

                // Check entered price
                int? enteredAmount = null;
                try
                {
                    enteredAmount = GetItemWindowPrice(offset);
                }
                catch (FormatException e)
                {
                    Logger.Error(e);
                    await CloseItemWindow();
                    goto BuyItem;
                }

                if (enteredAmount != buyPrice)
                {
                    Logger.Debug("Didn't enter price correctly");
                    await CloseItemWindow();
                    goto BuyItem;
                }

                int? currentQuantity = null;
                try
                {
                    currentQuantity = GetItemWindowQuantity(offset);
                }
                catch (FormatException e)
                {
                    Logger.Error(e);
                    await CloseItemWindow();
                    goto BuyItem;
                }

                if (currentQuantity == null)
                {
                    Logger.Debug("currentQuantity null");
                    await CloseItemWindow();
                    goto BuyItem;
                }

                if (quantity < currentQuantity.Value)
                {
                    DismissSuccessWithOffset(offset);

                    Input.MouseMoveAndDoubleClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 402, tradingPostPoint.Value.Y + 236 + offset);
                    Input.KeyStringSend(process!, quantity.ToString());
                }
                else
                {
                    quantity -= currentQuantity.Value;
                }

                Logger.Info($"Buying {currentQuantity.Value} for {buyPrice} coins each");

                DismissSuccessWithOffset(offset);
                LogImage();
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 440, tradingPostPoint.Value.Y + 373 + offset, 100);
                await Task.Delay(100);

                if (quantity > 0)
                {
                    y--;

                    await CloseItemWindow();
                    continue;
                }*/

                await CloseItemWindow();
                goto BuyItem;
            }
            else
            {
                Logger.Debug("Couldn't find buy button");
                await CloseItemWindow();
                goto BuyItem;
            }

            await CloseItemWindow();
            break;
        }
    }

    private static async Task BuyItems()
    {
        Logger.Info("====================");
        Logger.Info("Buying items");

        using var apiClient = new Gw2Client(ApiConnection);

        var restartAttempts = 0;
        RestartBuyItems:
        if (restartAttempts >= 5)
        {
            return;
        }

        restartAttempts++;

        List<Gw2Sharp.WebApi.V2.Models.CommerceTransactionCurrent>? currentSellingList = null;
        if (!Config.BuyIfSelling)
        {
            try
            {
                var currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(0);
                currentSellingList = currentSelling.ToList();
                var numSellingPages = currentSelling.HttpResponseInfo?.PageTotal;
                for (var i = 1; i < numSellingPages; i++)
                {
                    currentSelling = await apiClient.WebApi.V2.Commerce.Transactions.Current.Sells.PageAsync(i);
                    numSellingPages = currentSelling.HttpResponseInfo?.PageTotal;
                    currentSellingList = currentSellingList.Concat(currentSelling.ToList()).ToList();
                }
            }
            catch (Exception e)
            {
                Logger.Error(e);

                goto RestartBuyItems;
            }

            if (currentSellingList == null)
            {
                Logger.Debug("currentSellingList null");
                return;
            }
        }

        List<Gw2Sharp.WebApi.V2.Models.CommerceTransactionCurrent>? currentBuyingList = null;
        try
        {
            var currentBuying = await apiClient.WebApi.V2.Commerce.Transactions.Current.Buys.PageAsync(0);
            currentBuyingList = currentBuying.ToList();
            var numBuyingPages = currentBuying.HttpResponseInfo?.PageTotal;
            for (var i = 1; i < numBuyingPages; i++)
            {
                currentBuying = await apiClient.WebApi.V2.Commerce.Transactions.Current.Buys.PageAsync(i);
                numBuyingPages = currentBuying.HttpResponseInfo?.PageTotal;
                currentBuyingList = currentBuyingList.Concat(currentBuying.ToList()).ToList();
            }
        }
        catch (Exception e)
        {
            Logger.Error(e);

            goto RestartBuyItems;
        }

        if (currentBuyingList == null)
        {
            Logger.Debug("currentBuyingList null");
            return;
        }

        for (var endLoopIndex = currentLoopIndex + Config.BuysPerSellLoop; currentLoopIndex < endLoopIndex; currentLoopIndex++)
        {
            if (currentLoopIndex >= BuyItemsList.Count)
            {
                Logger.Info("Finished current loop of buys");
                return;
            }

            var item = BuyItemsList[currentLoopIndex];

            Logger.Info("--------------------");
            Logger.Info($"Buying Item {currentLoopIndex + 1}/{BuyItemsList.Count} [{item.Id}] {item.Name}");
            Logger.Info($"Buy: {item.BuyPrice} Sell: {item.SellPrice} Profit: {item.Profit} Sold: {item.Sold} Bought: {item.Bought}");

            // Check if currently selling item
            if (!Config.BuyIfSelling && currentSellingList!.Any(x => x.ItemId == item.Id))
            {
                Logger.Info("Skipping, already selling");
                continue;
            }

            // Check if already buying item
            if (currentBuyingList.Any(x => x.ItemId == item.Id))
            {
                Gw2Sharp.WebApi.V2.Models.CommercePrices? itemPrices = null;
                try
                {
                    itemPrices = await apiClient.WebApi.V2.Commerce.Prices.GetAsync(item.Id);
                }
                catch (Exception e)
                {
                    Logger.Error("Price exception");
                    Logger.Error(e);
                    continue;
                }

                if (itemPrices == null)
                {
                    Logger.Warn("itemPrices null");
                    continue;
                }

                var highestPrice = itemPrices.Buys.UnitPrice;
                var price = currentBuyingList.Where(x => x.ItemId == item.Id).Max(x => x.Price);

                // If we haven't been undercut
                if (highestPrice <= price)
                {
                    Logger.Info("Skipping, already buying");
                    continue;
                }

                Logger.Info("Cancelling undercut");

                try
                {
                    await CancelItem(item);
                }
                catch (Exception e)
                {
                    Logger.Error(e);
                    return;
                }
            }

            try
            {
                await BuyItem(item);
            }
            catch (Exception e)
            {
                Logger.Error(e);
                return;
            }
        }
    }

    private static async Task CancelItem(BuyItem item)
    {
        using var apiClient = new Gw2Client(ApiConnection);

        var restartAttempts = 0;
        RestartCancel:
        if (restartAttempts >= 5)
        {
            return;
        }

        restartAttempts++;

        // Get item info
        Gw2Sharp.WebApi.V2.Models.Item? itemInfo = null;
        try
        {
            itemInfo = await apiClient.WebApi.V2.Items.GetAsync(item.Id);
        }
        catch (JsonException)
        {
            return;
        }
        catch (Exception e)
        {
            Logger.Error(e);
            if (!ResetUI())
            {
                return;
            }

            goto RestartCancel;
        }

        if (itemInfo == null)
        {
            Logger.Warn("itemInfo null");
            return;
        }

        // Go to transactions screen
        try
        {
            await GoToScreen(TradingPostScreen.Transactions);
        }
        catch (TimeoutException)
        {
            Logger.Debug("Timeout on GoToScreen");
            if (!ResetUI())
            {
                return;
            }

            goto RestartCancel;
        }

        // Check if buying is not selected
        if (ImageSearch.FindImageInWindow(process!, Resources.CurrentBuying, tradingPostPoint!.Value.X + 60, tradingPostPoint!.Value.Y + 224, Resources.CurrentBuying.Width, Resources.CurrentBuying.Height, 0.9) != null)
        {
            // Wait for search results to load
            try
            {
                await WaitWhile(FindSearchResultsLoaded, 500, 10000);
            }
            catch (TimeoutException)
            {
                _ = ResetUI();
                return;
            }

            await Task.Delay(1000);

            // Click load more
            for (var i = 0; i < 5; i++)
            {
                if (FindLoadMore())
                {
                    break;
                }

                // Wait for load more button to appear
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 650, tradingPostPoint!.Value.Y + 688);
                Input.MouseMove(process!, 540, 0);
                await Task.Delay(500);
                try
                {
                    await WaitWhile(FindLoadMore, 500, 5000);
                }
                catch (TimeoutException)
                {
                    break;
                }
            }

            // Click buying
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 174, tradingPostPoint!.Value.Y + 240);
            await Task.Delay(1000);

            // Wait for search results to load
            try
            {
                await WaitWhile(FindSearchResultsLoaded, 500, 10000);
            }
            catch (TimeoutException)
            {
                _ = ResetUI();
                return;
            }
        }

        // Click load more
        for (var i = 0; i < 5; i++)
        {
            if (FindLoadMore())
            {
                break;
            }

            // Wait for load more button to appear
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 650, tradingPostPoint!.Value.Y + 688);
            Input.MouseMove(process!, 540, 0);
            await Task.Delay(500);
            try
            {
                await WaitWhile(FindLoadMore, 500, 5000);
            }
            catch (TimeoutException)
            {
                break;
            }
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

        // Wait for search results to load
        try
        {
            await WaitWhile(FindSearchResultsLoaded, 500, 5000);
        }
        catch (TimeoutException)
        {
            Logger.Error("Didn't load any search results");
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
            Logger.Debug("qtyPoint null");
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

            await Task.Delay(500);

            // Find the item border of the result
            var resultPoint = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, resultsPoint.X, resultsPoint.Y + (y * 61), Resources.ResultCorner.Width, Resources.ResultCorner.Height);
            if (resultPoint == null)
            {
                Logger.Debug("Result point null");
                break;
            }

            // Check if buying
            if (ImageSearch.FindImageInWindow(process!, Resources.Buying, resultsPoint.X + 350, resultsPoint.Y + 25 + (y * 61), Resources.Buying.Width, Resources.Buying.Height, 0.7) == null)
            {
                Logger.Debug("Not buying transaction");
                if (!ResetUI())
                {
                    return;
                }

                goto RestartCancel;
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
                Logger.Debug("Item border didn't appear");
                await CloseItemWindow();
                goto CancelItem;
            }

            // Wait for "vendor" to show up
            try
            {
                await WaitWhile(FindVendorE, 500, 5000);
            }
            catch (TimeoutException)
            {
                Logger.Debug("Vendor E didn't appear");
                await CloseItemWindow();
                goto CancelItem;
            }

            // Find position of 'e' in Vendor to set the offset for other UI positions
            var vendorEPos = ImageSearch.FindImageInWindow(process!, Resources.VendorE, tradingPostPoint!.Value.X + 345, tradingPostPoint!.Value.Y + 176, Resources.VendorE.Width, Resources.VendorE.Height + 31, 0.8);
            if (vendorEPos == null)
            {
                Logger.Debug("vendorEPos null");
                await CloseItemWindow();
                goto CancelItem;
            }

            var nameSize = 34;
            if (vendorEPos.Value.Y == tradingPostPoint!.Value.Y + 207)
            {
                nameSize = 64;
            }

            // Get OCR of item name and check if it matches API name
            var nameImage = ImageSearch.CaptureWindow(process!, tradingPostPoint!.Value.X + 334, tradingPostPoint.Value.Y + 136, 428, nameSize);
            var capturedName = OCR.ReadName(nameImage!, RarityColors[itemInfo.Rarity.ToString()!]);

            // Check captured name
            if (string.IsNullOrEmpty(capturedName) || !OCR.NameCompare(itemInfo.Name, capturedName, Config.OcrFixes))
            {
                Logger.Debug("Item name different");
                await CloseItemWindow();
                continue;
            }

            // Press X to close
            await CloseItemWindow();

            // Click cancel and confirm
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 600, resultsPoint.Y + 28 + (y * 61), 100);
            await Task.Delay(250);
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 600, resultsPoint.Y + 28 + (y * 61), 100);
            await Task.Delay(250);

            y--;
        }
    }

    private static async Task CancelAllBuying()
    {
        var restartAttempts = 0;
        RestartCancel:
        if (restartAttempts >= 5)
        {
            return;
        }

        restartAttempts++;

        // Go to transactions screen
        try
        {
            await GoToScreen(TradingPostScreen.Transactions);
        }
        catch (TimeoutException)
        {
            Logger.Debug("Timeout on GoToScreen");
            return;
        }

        // Check if buying is not selected
        if (ImageSearch.FindImageInWindow(process!, Resources.CurrentBuying, tradingPostPoint!.Value.X + 60, tradingPostPoint!.Value.Y + 224, Resources.CurrentBuying.Width, Resources.CurrentBuying.Height, 0.9) != null)
        {
            // Wait for search results to load
            try
            {
                await WaitWhile(FindSearchResultsLoaded, 500, 10000);
            }
            catch (TimeoutException)
            {
                _ = ResetUI();
                return;
            }

            await Task.Delay(1000);

            // Click load more
            for (var i = 0; i < 5; i++)
            {
                if (FindLoadMore())
                {
                    break;
                }

                // Wait for load more button to appear
                Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 650, tradingPostPoint!.Value.Y + 688);
                Input.MouseMove(process!, 540, 0);
                await Task.Delay(500);
                try
                {
                    await WaitWhile(FindLoadMore, 500, 5000);
                }
                catch (TimeoutException)
                {
                    break;
                }
            }

            // Click buying
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 174, tradingPostPoint!.Value.Y + 240);
            await Task.Delay(1000);

            // Wait for search results to load
            try
            {
                await WaitWhile(FindSearchResultsLoaded, 500, 10000);
            }
            catch (TimeoutException)
            {
                _ = ResetUI();
                return;
            }
        }

        // Click load more
        for (var i = 0; i < 5; i++)
        {
            if (FindLoadMore())
            {
                break;
            }

            // Wait for load more button to appear
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 650, tradingPostPoint!.Value.Y + 688);
            Input.MouseMove(process!, 540, 0);
            await Task.Delay(500);
            try
            {
                await WaitWhile(FindLoadMore, 500, 5000);
            }
            catch (TimeoutException)
            {
                break;
            }
        }

        // Wait for search results to load
        try
        {
            await WaitWhile(FindSearchResultsLoaded, 500, 5000);
        }
        catch (TimeoutException)
        {
            Logger.Error("Didn't load any search results");
            // LogImage();
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
            Logger.Debug("qtyPoint null");
            return;
        }

        var resultsPoint = Point.Add(qtyPoint.Value, new Size(-2, 22));

        while (true)
        {
            // Find the item border of the result
            var resultPoint = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, resultsPoint.X, resultsPoint.Y, Resources.ResultCorner.Width, Resources.ResultCorner.Height);
            if (resultPoint == null)
            {
                Logger.Debug("Result point null");
                break;
            }

            // Check if buying
            if (ImageSearch.FindImageInWindow(process!, Resources.Buying, resultsPoint.X + 350, resultsPoint.Y + 25, Resources.Buying.Width, Resources.Buying.Height, 0.7) == null)
            {
                Logger.Debug("Not buying transaction");
                if (!ResetUI())
                {
                    return;
                }

                goto RestartCancel;
            }

            // Click cancel and confirm
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 600, resultsPoint.Y + 28, 100);
            await Task.Delay(150);
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, resultsPoint.X + 600, resultsPoint.Y + 28, 100);
            await Task.Delay(150);
        }
    }

    /*private static async Task CancelUndercut()
    {
        Logger.Info("====================");
        Logger.Info("Cancelling undercut buys");

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
    }*/

    private static async Task CancelUnprofitable()
    {
        Logger.Info("====================");
        Logger.Info("Cancelling unprofitable items");

        using var apiClient = new Gw2Client(ApiConnection);

        var restartAttempts = 0;
        RestartCancelUnprofitable:
        if (restartAttempts >= 5)
        {
            return;
        }

        restartAttempts++;

        List<Gw2Sharp.WebApi.V2.Models.CommerceTransactionCurrent>? currentBuyingList = null;
        try
        {
            var currentBuying = await apiClient.WebApi.V2.Commerce.Transactions.Current.Buys.PageAsync(0);
            currentBuyingList = currentBuying.ToList();
            var numBuyingPages = currentBuying.HttpResponseInfo?.PageTotal;
            for (var i = 1; i < numBuyingPages; i++)
            {
                currentBuying = await apiClient.WebApi.V2.Commerce.Transactions.Current.Buys.PageAsync(i);
                numBuyingPages = currentBuying.HttpResponseInfo?.PageTotal;
                currentBuyingList = currentBuyingList.Concat(currentBuying.ToList()).ToList();
            }
        }
        catch (Exception e)
        {
            Logger.Error(e);

            goto RestartCancelUnprofitable;
        }

        if (currentBuyingList == null)
        {
            Logger.Debug("currentBuyingList null");
            return;
        }

        foreach (var item in RemoveItemsList)
        {
            // Check if we're currently buying item
            if (currentBuyingList.Any(x => x.ItemId == item.Id))
            {
                Logger.Info("-------------------- ");
                Logger.Info($"Cancelling [{item.Id}] {item.Name}");

                await CancelItem(item);
            }

            Logger.Info("--------------------");
            Logger.Info($"Selling [{item.Id}] {item.Name}");
            Logger.Info($"Buy: {item.BuyPrice} Sell: {item.SellPrice} Profit: {item.Profit} Sold: {item.Sold} Bought: {item.Bought}");

            await SellItem(item);
        }
    }

    private static void CheckNewMap()
    {
        if (DateTime.Now - timeSinceLastNewMapCheck < TimeSpan.FromMinutes(45))
        {
            return;
        }

        timeSinceLastNewMapCheck = DateTime.Now;

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
        if (DateTime.Now - timeSinceLastAfkCheck < TimeSpan.FromMinutes(15))
        {
            return;
        }

        timeSinceLastAfkCheck = DateTime.Now;

        Logger.Info("Moving for anti-afk");
        Input.KeyPress(process!, VirtualKeyCode.RIGHT, 100);
        Thread.Sleep(1000);
        Input.KeyPress(process!, VirtualKeyCode.LEFT, 100);
    }

    private static void PlayCharacterCheck()
    {
        var errorPoint = ImageSearch.FindImageInFullWindow(process!, Resources.ErrorOK, 0.9);
        if (errorPoint != null)
        {
            Logger.Info("Found error message box");
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, errorPoint.Value.X + 10, errorPoint.Value.Y + 7);
            Thread.Sleep(500);
        }

        if (ImageSearch.FindImageInFullWindow(process!, Resources.ReturnToGame, 0.9) != null)
        {
            Logger.Info("Found escape menu");
            Input.KeyPress(process!, VirtualKeyCode.ESCAPE);
            Thread.Sleep(500);
        }

        Input.MouseMove(process!, 540, 0);
        Input.MouseClick(process!, Input.MouseButton.LeftButton);
        Thread.Sleep(1000);

        if (ImageSearch.FindImageInWindow(process!, Resources.Play, 861, 782, Resources.Play.Width, Resources.Play.Height, 0.6) != null)
        {
            Logger.Info("Logging in to character");

            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, 872, 788);
            Thread.Sleep(30000);
        }
    }

    /*private static void CheckPosition()
    {
        using var apiClient = new Gw2Client(ApiConnection);

        apiClient.Mumble.Update();

        var position = apiClient.Mumble.AvatarPosition;
    }*/

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

    /*private static bool IsStackable(string type) => type is "Consumable"
        or "Container"
        or "CraftingMaterial"
        or "Trophy"
        or "UpgradeComponent";*/

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
        var resultBorder = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, tradingPostPoint!.Value.X + 322, tradingPostPoint!.Value.Y + 177, Resources.ResultCorner.Width, Resources.ResultCorner.Height + 72, 0.98);
        return noItems == null && resultBorder == null;
    }

    private static bool FindFilter()
    {
        var find = ImageSearch.FindImageInWindow(process!, Resources.FilterOpen, tradingPostPoint!.Value.X + 50, tradingPostPoint!.Value.Y + 182, Resources.FilterOpen.Width, Resources.FilterOpen.Height, 0.9);
        return find == null;
    }

    private static bool FindItemBorder()
    {
        var find = ImageSearch.FindImageInWindow(process!, Resources.ResultCorner, tradingPostPoint!.Value.X + 271, tradingPostPoint!.Value.Y + 136, Resources.ResultCorner.Width, Resources.ResultCorner.Height, 0.98);
        return find == null;
    }

    private static bool FindVendorE()
    {
        var find = ImageSearch.FindImageInWindow(process!, Resources.VendorE, tradingPostPoint!.Value.X + 345, tradingPostPoint!.Value.Y + 176, Resources.VendorE.Width, Resources.VendorE.Height + 31, 0.8);
        return find == null;
    }

    private static bool FindAvailable()
    {
        var find1 = ImageSearch.FindImageInWindow(process!, Resources.Available, tradingPostPoint!.Value.X + 531, tradingPostPoint!.Value.Y + 469, Resources.Available.Width, Resources.Available.Height, 0.9);
        var find2 = ImageSearch.FindImageInWindow(process!, Resources.Available, tradingPostPoint!.Value.X + 531, tradingPostPoint!.Value.Y + 494, Resources.Available.Width, Resources.Available.Height, 0.9);
        return find1 == null && find2 == null;
    }

    private static bool FindLoadMore()
    {
        var find1 = ImageSearch.FindImageInWindow(process!, Resources.LoadMore, tradingPostPoint!.Value.X + 610, tradingPostPoint!.Value.Y + 679, Resources.LoadMore.Width, Resources.LoadMore.Height, 0.9);
        var find2 = ImageSearch.FindImageInWindow(process!, Resources.LoadMoreLight, tradingPostPoint!.Value.X + 610, tradingPostPoint!.Value.Y + 679, Resources.LoadMoreLight.Width, Resources.LoadMoreLight.Height, 0.9);
        return find1 == null && find2 == null;
    }

    private static void DismissSuccess()
    {
        if (ImageSearch.FindImageInWindow(process!, Resources.SuccessOK, tradingPostPoint!.Value.X + 392, tradingPostPoint!.Value.Y + 365, Resources.SuccessOK.Width, Resources.SuccessOK.Height) != null)
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 450, tradingPostPoint!.Value.Y + 374);
            Thread.Sleep(1000);
        }
        else if (ImageSearch.FindImageInWindow(process!, Resources.SuccessOK, tradingPostPoint!.Value.X + 392, tradingPostPoint!.Value.Y + 365 + 24, Resources.SuccessOK.Width, Resources.SuccessOK.Height) != null)
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 450, tradingPostPoint!.Value.Y + 374 + 24);
            Thread.Sleep(1000);
        }
    }

    private static void DismissSuccessWithOffset(int offset)
    {
        if (ImageSearch.FindImageInWindow(process!, Resources.SuccessOK, tradingPostPoint!.Value.X + 392, tradingPostPoint!.Value.Y + 365 + offset, Resources.SuccessOK.Width, Resources.SuccessOK.Height) != null)
        {
            Input.MouseMoveAndClick(process!, Input.MouseButton.LeftButton, tradingPostPoint!.Value.X + 450, tradingPostPoint!.Value.Y + 374 + offset);
            Thread.Sleep(1000);
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

            var fileName = Path.Combine(directory, $"{DateTime.Now.ToString("s").Replace(":", string.Empty)}.bmp");
            Logger.Error($"File saved as {fileName}");
            image.Save(fileName, ImageFormat.Bmp);
        }
    }

    private static void LogFullImage()
    {
        var image = ImageSearch.CaptureFullWindow(process!);
        if (image != null)
        {
            var directory = Path.Combine("logs", "images");
            if (!Directory.Exists(directory))
            {
                _ = Directory.CreateDirectory(directory);
            }

            var fileName = Path.Combine(directory, $"{DateTime.Now.ToString("s").Replace(":", string.Empty)}.bmp");
            Logger.Error($"File saved as {fileName}");
            image.Save(fileName, ImageFormat.Bmp);
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
            var directory = Path.Combine("logs", "images");
            if (!Directory.Exists(directory))
            {
                _ = Directory.CreateDirectory(directory);
            }

            var fileName = Path.Combine(directory, $"{DateTime.Now.ToString("s").Replace(":", string.Empty)}.bmp");
            Logger.Error($"File saved as {fileName}");
            image.Save(fileName, ImageFormat.Bmp);
        }
    }

    private static void LogSpecificImage(string name, int x, int y, int width, int height)
    {
        var image = ImageSearch.CaptureWindow(process!, x, y, width, height);
        if (image != null)
        {
            var directory = Path.Combine("logs", "images");
            if (!Directory.Exists(directory))
            {
                _ = Directory.CreateDirectory(directory);
            }

            var fileName = Path.Combine(directory, $"{name}.bmp");
            Logger.Error($"File saved as {fileName}");
            image.Save(fileName, ImageFormat.Bmp);
        }
    }
}
