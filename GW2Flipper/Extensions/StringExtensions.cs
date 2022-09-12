namespace GW2Flipper.Extensions;

using System.Globalization;
using System.Text;

public static class StringExtensions
{
    public static string MaxSize(this string str, int size) => str.Length < size ? str : str[..size];

    public static string RemoveDiacritics(this string str)
    {
        var normalizedString = str.Normalize(NormalizationForm.FormD);
        var stringBuilder = new StringBuilder();

        foreach (var c in normalizedString.EnumerateRunes())
        {
            var unicodeCategory = Rune.GetUnicodeCategory(c);
            if (unicodeCategory != UnicodeCategory.NonSpacingMark)
            {
                stringBuilder.Append(c);
            }
        }

        return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
    }
}
