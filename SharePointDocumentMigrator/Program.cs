using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using SharePointDocumentMigrator.Graph;
using SharePointDocumentMigrator.Migration;
using SharePointDocumentMigrator.Options;

using var cancellationSource = new CancellationTokenSource();
Console.CancelKeyPress += (_, args) =>
{
    args.Cancel = true;
    cancellationSource.Cancel();
};

try
{
    var builder = Host.CreateApplicationBuilder(args);

    builder.Services
        .AddOptions<AuthenticationOptions>()
        .Bind(builder.Configuration.GetSection(AuthenticationOptions.SectionName))
        .ValidateOnStart();

    builder.Services
        .AddOptions<MigrationOptions>()
        .Bind(builder.Configuration.GetSection(MigrationOptions.SectionName))
        .ValidateOnStart();

    builder.Services.AddSingleton<IValidateOptions<AuthenticationOptions>, AuthenticationOptionsValidator>();
    builder.Services.AddSingleton<IValidateOptions<MigrationOptions>, MigrationOptionsValidator>();

    builder.Services.AddSingleton<TokenCredential>(serviceProvider =>
    {
        var authOptions = serviceProvider.GetRequiredService<IOptions<AuthenticationOptions>>().Value;
        var logger = serviceProvider.GetRequiredService<ILoggerFactory>().CreateLogger("Authentication");

        if (authOptions.UseDeviceCode)
        {
            return new DeviceCodeCredential(
                new DeviceCodeCredentialOptions
                {
                    TenantId = authOptions.TenantId,
                    ClientId = authOptions.ClientId,
                    DeviceCodeCallback = (code, _) =>
                    {
                        logger.LogInformation("{Message}", code.Message);
                        return Task.CompletedTask;
                    }
                });
        }

        if (string.IsNullOrWhiteSpace(authOptions.ClientSecret))
            throw new InvalidOperationException("Client secret must be provided when device code authentication is disabled.");

        authOptions.ApplyCustomScopes("https://graph.microsoft.com/.default");

        return new ClientSecretCredential(
                authOptions.TenantId,
                authOptions.ClientId,
                authOptions.ClientSecret);
    });

    builder.Services.AddHttpClient<GraphApiClient>(client =>
    {
        client.BaseAddress = new Uri("https://graph.microsoft.com/v1.0/");
        client.Timeout = TimeSpan.FromMinutes(15);
    });

    builder.Services.AddSingleton<MigrationOrchestrator>();

    using var host = builder.Build();

    var orchestrator = host.Services.GetRequiredService<MigrationOrchestrator>();
    var report = await orchestrator.ExecuteAsync(cancellationSource.Token);

    return report.FailedFiles == 0 ? 0 : 1;
}
catch (OptionsValidationException exception)
{
    await Console.Error.WriteLineAsync("Configuration validation failed:");
    foreach (var failure in exception.Failures)
    {
        await Console.Error.WriteLineAsync($" - {failure}");
    }

    return 1;
}
catch (Exception exception)
{
    await Console.Error.WriteLineAsync(exception.Message);
    return 1;
}
