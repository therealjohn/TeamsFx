// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace Microsoft.TeamsFx;

/// <summary>
/// Contains settings for an <see cref="TeamsBotSsoPrompt"/>.
/// </summary>
public class TeamsBotSsoPromptSettings
{
    /// <summary>
    /// Gets or sets the array of strings that declare the desired permissions and the resources requested.
    /// </summary>
    /// <value>The array of strings that declare the desired permissions and the resources requested.</value>
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

}