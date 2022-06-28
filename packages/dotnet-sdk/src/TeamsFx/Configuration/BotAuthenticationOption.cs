// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.ComponentModel.DataAnnotations;

namespace Microsoft.TeamsFx.Configuration;

/// <summary>
/// Bot related authentication configuration.
/// </summary>
public class BotAuthenticationOptions
{
    /// <summary>
    /// The client (application) ID of an App Registration in the tenant.
    /// </summary>
    [Required(ErrorMessage = "Client id is required")]
    [RegularExpression(@"^[0-9A-Fa-f\-]{36}$")]
    public string ClientId { get; set; }

    /// <summary>
    /// AAD tenant id.
    /// </summary>
    [Required(ErrorMessage = "Tenant id is required")]
    [RegularExpression(@"^[0-9A-Fa-f\-]{36}$")]
    public string TenantId { get; set; }

    /// <summary>
    /// Application ID URI.
    /// </summary>
    [Required(ErrorMessage = "Application id uri is required")]
    public string ApplicationIdUri { get; set; }
}
