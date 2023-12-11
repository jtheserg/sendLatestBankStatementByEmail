using Microsoft.Extensions.Configuration;
using System.Diagnostics.CodeAnalysis;

namespace sendLatestBankStatementByEmail;

internal class Startup
{
    [RequiresUnreferencedCode("Calls Microsoft.Extensions.Configuration.ConfigurationBinder.Get<T>()")]
    [RequiresDynamicCode("Calls Microsoft.Extensions.Configuration.ConfigurationBinder.Get<T>()")]
    public Startup()
    {
        var builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("settings.json", optional: false);

        IConfiguration config = builder.Build();

        Settings = config.GetSection("Settings").Get<Settings>() ?? throw new InvalidOperationException();
    }

    public Settings Settings { get; private set; }
}