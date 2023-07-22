namespace GW2Flipper.Extensions;

internal static class ListExtensions
{
    public static void Shuffle<T>(this IList<T> list)
    {
        var random = new Random();

        for (var i = list.Count - 1; i >= 0; i--)
        {
            var swapIndex = random.Next(i + 1);

            (list[i], list[swapIndex]) = (list[swapIndex], list[i]);
        }
    }
}
