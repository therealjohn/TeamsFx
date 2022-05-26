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
    public int? Timeout { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the <see cref="TeamsBotSsoPrompt"/> should end upon
    /// receiving an invalid message.  Generally the <see cref="TeamsBotSsoPrompt"/> will ignore
    /// incoming messages from the user during the auth flow, if they are not related to the
    /// auth flow.  This flag enables ending the <see cref="TeamsBotSsoPrompt"/> rather than
    /// ignoring the user's message.  Typically, this flag will be set to 'true', but is 'false'
    /// by default for backwards compatibility.
    /// </summary>
    /// <value>True if the <see cref="TeamsBotSsoPrompt"/> should automatically end upon receiving
    /// an invalid message.</value>
    public bool EndOnInvalidMessage { get; set; }

}