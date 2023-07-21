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
                _ = stringBuilder.Append(c);
            }
        }

        return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
    }

    public static string StripPunctuation(this string str)
    {
        var stringBuilder = new StringBuilder();

        foreach (var c in str)
        {
            if (!char.IsPunctuation(c))
            {
                _ = stringBuilder.Append(c);
            }
        }

        return stringBuilder.ToString();
    }
}
