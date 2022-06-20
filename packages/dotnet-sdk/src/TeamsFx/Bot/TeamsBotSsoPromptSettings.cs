// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx;
using System.ComponentModel.DataAnnotations;

/// <summary>
/// Contains settings for an <see cref="TeamsBotSsoPrompt"/>.
/// </summary>
public class TeamsBotSsoPromptSettings
{

    /// <summary>
    /// Constructor of TeamsBotSsoPromptSettings
    /// </summary>
    public TeamsBotSsoPromptSettings(string[] scopes, string clientId, string tenantId, string applicationIdUri, int timeout = 900000, bool endOnInvalidMessage = true)
    {
        Scopes = scopes;
        Timeout = timeout;
        EndOnInvalidMessage = endOnInvalidMessage;
        ClientId = clientId;
        TenantId = tenantId;
        ApplicationIdUri = applicationIdUri;
    }

    /// <summary>
    /// Gets or sets the array of strings that declare the desired permissions and the resources requested.
    /// </summary>
    /// <value>The array of strings that declare the desired permissions and the resources requested.</value>
    [Required(ErrorMessage = "Scope is required")]
    public string[] Scopes { get; set; }

    /// <summary>
    /// Gets or sets the number of milliseconds the prompt waits for the user to authenticate.
    /// Default is 900,000 (15 minutes).
    /// </summary>
    /// <value>The number of milliseconds the prompt waits for the user to authenticate.</value>
    public int Timeout { get; set; } = (int)TimeSpan.FromMinutes(15).TotalMilliseconds;

    /// <summary>
    /// Gets or sets a value indicating whether the <see cref="TeamsBotSsoPrompt"/> should end upon
    /// receiving an invalid message.  Generally the <see cref="TeamsBotSsoPrompt"/> will end 
    /// the auth flow when receives user message not related to the auth flow.
    /// Setting the flag to false ignores the user's message instead.
    /// Defaults to value `true`
    /// </summary>
    /// <value>True if the <see cref="TeamsBotSsoPrompt"/> should automatically end upon receiving
    /// an invalid message.</value>
    public bool EndOnInvalidMessage { get; set; } = true;

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