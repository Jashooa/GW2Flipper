namespace GW2Flipper;

using NLog;
using NLog.Config;
using NLog.Targets;

internal static class Program
{
    private static async Task Main()
    {
        SetupLogger();
        await GW2Flipper.Run();
        Console.ReadKey();
    }

    private static void SetupLogger()
    {
        var config = new LoggingConfiguration();

        var consoleTarget = new ColoredConsoleTarget()
        {
            Layout = "[${time}][${level:format=FirstCharacter}] ${message} ${onexception:${exception:format=type}}",
        };

        var fileTarget = new FileTarget()
        {
            FileName = $"${{basedir}}{Path.DirectorySeparatorChar}logs{Path.DirectorySeparatorChar}current.log",
            ArchiveFileName = $"${{basedir}}{Path.DirectorySeparatorChar}logs{Path.DirectorySeparatorChar}{{#}}.log",
            ArchiveNumbering = ArchiveNumberingMode.Date,
            ArchiveEvery = FileArchivePeriod.Day,
            ArchiveDateFormat = "yyyy-MM-dd",
            Layout = "[${longdate}][${level:format=FirstCharacter}](${logger}) ${message} ${onexception:${newline}${exception:format=tostring,data}}",
        };

        var consoleRule = new LoggingRule("*", LogLevel.Info, consoleTarget);
        var fileRule = new LoggingRule("*", LogLevel.Debug, fileTarget);

        config.LoggingRules.Add(consoleRule);
        config.LoggingRules.Add(fileRule);

        LogManager.Configuration = config;
    }
}
