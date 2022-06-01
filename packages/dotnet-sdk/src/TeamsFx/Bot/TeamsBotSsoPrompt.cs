// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Extensions.Logging;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Bot.Schema;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Options;
using Microsoft.TeamsFx.Configuration;

namespace Microsoft.TeamsFx;

/// <summary>
/// Creates a new prompt that leverage Teams Single Sign On (SSO) support for bot to automatically sign in user and
/// help receive oauth token, asks the user to consent if needed.
/// </summary>
/// <remarks>
/// The prompt will attempt to retrieve the users current token of the desired scopes and store it in
/// the token store. 
/// User will be automatically signed in leveraging Teams support of Bot Single Sign On(SSO):
/// https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/authentication/auth-aad-sso-bots
/// </remarks>
public class TeamsBotSsoPrompt : Dialog
{
    private TeamsBotSsoPromptSettings _settings;
    private readonly AuthenticationOptions _authenticationOptions;
    private const string PersistedExpires = "expires";
    #region Util
    private readonly ILogger<TeamsBotSsoPrompt> _logger;
    #endregion


    /// <summary>
    /// Initializes a new instance of the <see cref="TeamsBotSsoPrompt"/> class.
    /// </summary>
    /// <param name="authenticationOptions">Authentication options filled by DI.</param>
    /// <param name="dialogId">The ID to assign to this prompt.</param>
    /// <param name="logger">Logger of TeamsBotSsoPrompt Class.</param>
    /// <param name="settings">Additional OAuth settings to use with this instance of the prompt.
    /// custom validation for this prompt.</param>
    /// <remarks>The value of <paramref name="dialogId"/> must be unique within the
    /// <see cref="DialogSet"/> or <see cref="ComponentDialog"/> to which the prompt is added.</remarks>
    public TeamsBotSsoPrompt(string dialogId, TeamsBotSsoPromptSettings settings,
        ILogger<TeamsBotSsoPrompt> logger,
        IOptions<AuthenticationOptions> authenticationOptions) : base(dialogId)
    {
        _logger = logger;
        if (string.IsNullOrWhiteSpace(dialogId))
        {
            throw new ArgumentNullException(nameof(dialogId));
        }

        _settings = settings ?? throw new ArgumentNullException(nameof(settings));

        if (authenticationOptions == null)
        {
            throw new ArgumentNullException(nameof(authenticationOptions));
        }
        try
        {
            _logger.LogTrace("Validate authentication configuration");
            _authenticationOptions = authenticationOptions.Value;
        }
        catch (OptionsValidationException e)
        {
            throw new ExceptionWithCode($"Authentication config is missing or not correct with error: {e.Message}", ExceptionCode.InvalidConfiguration);
        }

     
        _logger.LogInformation("Create a teams bot sso prompt");
    }


    /// <summary>
    /// Called when the dialog is started and pushed onto the dialog stack.
    /// Developer need to configure TeamsFx service before using this class.
    /// </summary>
    /// <param name="dc">The Microsoft.Bot.Builder.Dialogs.DialogContext for the current turn of conversation.</param>
    /// <param name="options">Optional, initial information to pass to the dialog.</param>
    /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
    /// <returns> A System.Threading.Tasks.Task representing the asynchronous operation.</returns>
    /// <exception cref="ArgumentException">if dialog context argument is null</exception>
    /// <remarks>
    /// If the task is successful, the result indicates whether the dialog is still active after the turn has been processed by the dialog.
    /// </remarks>
    public override async Task<DialogTurnResult> BeginDialogAsync(DialogContext dc, object options = null, CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Begin teams bot sso prompt dialog");
        if (dc == null)
        {
            throw new ArgumentNullException(nameof(dc));
        }

        EnsureMsTeamsChannel(dc);

        // Initialize state
        var state = dc.ActiveDialog.State;
        state[PersistedExpires] = DateTime.UtcNow.AddMilliseconds(_settings.Timeout);

        // Send OAuthCard that tells Teams to obtain an authentication token for the bot application.
        await SendOAuthCardToObtainTokenAsync(dc.Context).ConfigureAwait(false);
        return EndOfTurn;
    }

    /// <summary>
    /// Send OAuthCard that tells Teams to obtain an authentication token for the bot application.
    /// For details see https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/authentication/auth-aad-sso-bots.
    /// </summary>
    /// <param name="context">ITurnContext</param>
    /// <returns>The task to await.</returns>
    private async Task SendOAuthCardToObtainTokenAsync(ITurnContext context)
    {
        _logger.LogDebug("Send OAuth card to get SSO token");

        TeamsChannelAccount account = await TeamsInfo.GetMemberAsync(context, context.Activity.From.Id).ConfigureAwait(false);
        _logger.LogDebug(
          "Get Teams member account user principal name: " + account.UserPrincipalName
        );

        string loginHint = account.UserPrincipalName ?? "";
        SignInResource signInResource = GetSignInResource(loginHint);

        // Ensure prompt initialized
        IMessageActivity prompt = Activity.CreateMessageActivity();
        prompt.Attachments = new List<Attachment>();
        prompt.Attachments.Add(new Attachment {
            ContentType = OAuthCard.ContentType,
            Content = new OAuthCard {
                Text = "Sign In",
                Buttons = new[]
                            {
                                new CardAction
                                {
                                    Title = "Teams SSO Sign In",
                                    Value = signInResource.SignInLink,
                                    Type = ActionTypes.Signin,
                                },
                            },
                TokenExchangeResource = signInResource.TokenExchangeResource
            },
        });

        // Send prompt
        await context.SendActivityAsync(prompt).ConfigureAwait(false);
    }


    /// <summary>
    /// Get sign in authentication configuration
    /// </summary>
    /// <param name="loginHint"></param>
    /// <returns>sign in resource</returns>
    private SignInResource GetSignInResource(string loginHint)
    {
        _logger.LogDebug("Get sign in authentication configuration");
        string signInLink = $"{_authenticationOptions.InitiateLoginEndpoint}?scope={Uri.EscapeDataString(string.Join(" ", _settings.Scopes))}&clientId={_authenticationOptions.ClientId}&tenantId={_authenticationOptions.TenantId}&loginHint={loginHint}";
        _logger.LogDebug("Sign in link: " + signInLink);

        SignInResource signInResource = new SignInResource {
            SignInLink = signInLink,
            TokenExchangeResource = new TokenExchangeResource {
                Id = Guid.NewGuid().ToString(),
                Uri = Regex.Replace(_authenticationOptions.ApplicationIdUri, @"/\/$/", "/access_as_user")
            }
        };
        _logger.LogDebug("Token exchange resource uri: " + signInResource.TokenExchangeResource.Uri);

        return signInResource;
    }

    /// <summary>
    /// Ensure bot is running in MS Teams since TeamsBotSsoPrompt is only supported in MS Teams channel.
    /// </summary>
    /// <param name="dc">dialog context</param>
    /// <exception cref="ExceptionCode.ChannelNotSupported"> if bot channel is not MS Teams </exception>
    private void EnsureMsTeamsChannel(DialogContext dc)
    {
        if (dc.Context.Activity.ChannelId != Bot.Connector.Channels.Msteams)
        {
            var errorMessage = "Teams Bot SSO Prompt is only supported in MS Teams Channel";
            _logger.LogError(errorMessage);
            throw new ExceptionWithCode(errorMessage, ExceptionCode.ChannelNotSupported);
        }
    }
}
