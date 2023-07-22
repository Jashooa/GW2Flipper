namespace GW2Flipper.Models;

internal class BuyItem
{
    public BuyItem(int id, string name, int buyPrice, int sellPrice, int profit, int sold, int bought)
    {
        Id = id;
        Name = name;
        BuyPrice = buyPrice;
        SellPrice = sellPrice;
        Profit = profit;
        Sold = sold;
        Bought = bought;
    }

    public int Id { get; set; }

    public string Name { get; set; }

    public int BuyPrice { get; set; }

    public int SellPrice { get; set; }

    public int Profit { get; set; }

    public int Sold { get; set; }

    public int Bought { get; set; }

    public override string ToString() => $"[{Id}] {Name} Buy: {BuyPrice} Sell: {SellPrice} Profit: {Profit} Sold: {Sold} Bought: {Bought}";
}
