// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Client;
using Microsoft.TeamsFx.Helper;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.TeamsFx.Credential;

namespace Microsoft.TeamsFx.Configuration;

/// <summary>
/// Service Registration
/// </summary>
public static class TeamsFxConfigurationMethods
{
    /// <summary>
    /// Add TeamsFx SDK.
    /// </summary>
    /// <param name="services">service collection for DI</param>
    /// <param name="namedConfigurationSection">configuration instance</param>
    /// <returns></returns>
    public static IServiceCollection AddTeamsFx(
        this IServiceCollection services,
        IConfiguration namedConfigurationSection)
    {
        services.AddHttpClient();
        services.AddOptions();
        services.AddScoped<TeamsFx>();
        services.AddScoped<TeamsUserCredential>();

        services.AddOptions<AuthenticationOptions>().Bind(namedConfigurationSection.GetSection(AuthenticationOptions.Authentication)).ValidateDataAnnotations();
        services.AddOptions<BotAuthenticationOptions>().Configure<IOptions<AuthenticationOptions>>((botAuthOption, authOptions) => {
            AuthenticationOptions authOptionsValue = authOptions.Value;
            botAuthOption.ClientId = authOptionsValue.ClientId;
            botAuthOption.TenantId = authOptionsValue.TenantId;
            botAuthOption.ApplicationIdUri = authOptionsValue.ApplicationIdUri;
        }).ValidateDataAnnotations();
        
        services.AddSingleton<IIdentityClientAdapter>(sp => {
            var authenticationOptions = sp.GetRequiredService<IOptions<AuthenticationOptions>>().Value;
            var builder = ConfidentialClientApplicationBuilder.Create(authenticationOptions.ClientId)
                .WithClientSecret(authenticationOptions.ClientSecret)
                .WithAuthority(authenticationOptions.OAuthAuthority);
            var identityClientAdapter = new IdentityClientAdapter(builder.Build());
            return identityClientAdapter;
        });
        services.AddSingleton<OnBehalfOfUserCredential>();
        
        return services;
    }

    /// <summary>
    /// Add TeamsFx SDK.
    /// </summary>
    /// <param name="services">service collection for DI</param>
    /// <param name="configureOptions">customized action to configure option</param>
    /// <returns></returns>
    public static IServiceCollection AddTeamsFx(
        this IServiceCollection services,
        Action<AuthenticationOptions> configureOptions)
    {
        services.AddHttpClient();
        services.AddOptions();
        services.AddScoped<TeamsFx>();
        services.AddScoped<TeamsUserCredential>();

        services.Configure(configureOptions);
        services.AddOptions<AuthenticationOptions>()
            .Configure(configureOptions).ValidateDataAnnotations();

        services.AddSingleton<IIdentityClientAdapter>(sp => {
            var authenticationOptions = sp.GetRequiredService<IOptions<AuthenticationOptions>>().Value;
            var builder = ConfidentialClientApplicationBuilder.Create(authenticationOptions.ClientId)
                .WithClientSecret(authenticationOptions.ClientSecret)
                .WithAuthority(authenticationOptions.OAuthAuthority);
            var identityClientAdapter = new IdentityClientAdapter(builder.Build());

            return identityClientAdapter;
        });

        return services;
    }

    /// <summary>
    /// Add TeamsFx SDK.
    /// </summary>
    /// <param name="services">service collection for DI</param>
    /// <param name="userOptions">customized option instance</param>
    /// <returns></returns>
    public static IServiceCollection AddTeamsFx(
        this IServiceCollection services,
        AuthenticationOptions userOptions)
    {
        services.AddHttpClient();
        services.AddOptions();
        services.AddScoped<TeamsFx>();
        services.AddScoped<TeamsUserCredential>();

        services.AddOptions<AuthenticationOptions>()
            .Configure(options => {
                options.ClientId = userOptions.ClientId;
                options.ClientSecret = userOptions.ClientSecret;
                options.OAuthAuthority = userOptions.OAuthAuthority;
            }).ValidateDataAnnotations();

        services.AddSingleton<IIdentityClientAdapter>(sp => {
            var authenticationOptions = sp.GetRequiredService<IOptions<AuthenticationOptions>>().Value;
            var builder = ConfidentialClientApplicationBuilder.Create(authenticationOptions.ClientId)
                .WithClientSecret(authenticationOptions.ClientSecret)
                .WithAuthority(authenticationOptions.OAuthAuthority);

            var identityClientAdapter = new IdentityClientAdapter(builder.Build());
            return identityClientAdapter;
        });

        return services;
    }
}
