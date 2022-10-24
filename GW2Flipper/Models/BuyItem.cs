namespace GW2Flipper.Models;
internal class BuyItem
{
    public BuyItem(int id, string name, int buyPrice, int sellPrice, int profit, int sold)
    {
        Id = id;
        Name = name;
        BuyPrice = buyPrice;
        SellPrice = sellPrice;
        Profit = profit;
        Sold = sold;
    }

    public int Id { get; set; }

    public string Name { get; set; }

    public int BuyPrice { get; set; }

    public int SellPrice { get; set; }

    public int Profit { get; set; }

    public int Sold { get; set; }

    public override string ToString() => $"[{Id}] {Name} Buy: {BuyPrice} Sell: {SellPrice} Profit: {Profit} Sold: {Sold}";
}
