// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Extensions.Logging;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Bot.Schema;
using System.Text.RegularExpressions;
using System.ComponentModel.DataAnnotations;
using System.Net;
using Azure.Core;
using System.IdentityModel.Tokens.Jwt;
using Newtonsoft.Json.Linq;

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
    private const string PersistedExpires = "expires";
    #region Util
    private readonly ILogger<TeamsBotSsoPrompt> _logger;
    #endregion


    /// <summary>
    /// Initializes a new instance of the <see cref="TeamsBotSsoPrompt"/> class.
    /// </summary>
    /// <param name="dialogId">The ID to assign to this prompt.</param>
    /// <param name="logger">Logger of TeamsBotSsoPrompt Class.</param>
    /// <param name="settings">Additional OAuth settings to use with this instance of the prompt.
    /// custom validation for this prompt.</param>
    /// <remarks>The value of <paramref name="dialogId"/> must be unique within the
    /// <see cref="DialogSet"/> or <see cref="ComponentDialog"/> to which the prompt is added.</remarks>
    public TeamsBotSsoPrompt(string dialogId, TeamsBotSsoPromptSettings settings, ILogger<TeamsBotSsoPrompt> logger) : base(dialogId)
    {
        _logger = logger;
        if (string.IsNullOrWhiteSpace(dialogId))
        {
            throw new ArgumentNullException(nameof(dialogId));
        }
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        ValidateTeamsBotSsoPromptSettings(_settings);

        _logger.LogInformation("Create a teams bot sso prompt");
    }

    private void ValidateTeamsBotSsoPromptSettings(TeamsBotSsoPromptSettings settings)
    {
        _logger.LogTrace("Validate teams bot sso prompt settings.");
        var results = new List<ValidationResult>();
        var isValid = Validator.TryValidateObject(settings, new ValidationContext(settings), results, true);
        if (isValid)
        {
            _logger.LogInformation("Teams bot sso prompt settings are valid");
        } else
        {
            string errorMessage = "Teams bot sso prompt settings are missing or not correct with error: ";
            foreach (var validationResult in results)
            {
                errorMessage += validationResult.ErrorMessage + ". ";
            }
            throw new ExceptionWithCode(errorMessage, ExceptionCode.InvalidConfiguration);
        }
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

        var state = dc.ActiveDialog?.State;
        state[PersistedExpires] = DateTime.UtcNow.AddMilliseconds(_settings.Timeout);

        // Send OAuthCard that tells Teams to obtain an authentication token for the bot application.
        await SendOAuthCardToObtainTokenAsync(dc.Context, cancellationToken).ConfigureAwait(false);
        return EndOfTurn;
    }

    /// <summary>
    /// Called when the dialog is _continued_, where it is the active dialog and the
    /// user replies with a new activity.
    /// </summary>
    /// <param name="dc">The <see cref="DialogContext"/> for the current turn of conversation.</param>
    /// <param name="cancellationToken">A cancellation token that can be used by other objects
    /// or threads to receive notice of cancellation.</param>
    /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
    /// <remarks>If the task is successful, the result indicates whether the dialog is still
    /// active after the turn has been processed by the dialog. The result may also contain a
    /// return value.
    ///
    /// If this method is *not* overridden, the dialog automatically ends when the user replies.
    /// </remarks>
    /// <seealso cref="DialogContext.ContinueDialogAsync(CancellationToken)"/>
    public override async Task<DialogTurnResult> ContinueDialogAsync(DialogContext dc, CancellationToken cancellationToken = default(CancellationToken))
    {
        _logger.LogInformation("Teams bot sso prompt continue dialog");
        EnsureMsTeamsChannel(dc);

        // Check for timeout
        var state = dc.ActiveDialog?.State;
        bool isMessage = (dc.Context.Activity.Type == ActivityTypes.Message);
        bool isTimeoutActivityType =
          isMessage ||
          IsTeamsVerificationInvoke(dc.Context) ||
          IsTokenExchangeRequestInvoke(dc.Context);

        // If the incoming Activity is a message, or an Activity Type normally handled by TeamsBotSsoPrompt,
        // check to see if this TeamsBotSsoPrompt Expiration has elapsed, and end the dialog if so.
        bool hasTimedOut = isTimeoutActivityType && DateTime.Compare(DateTime.UtcNow, (DateTime)state[PersistedExpires]) > 0;
        if (hasTimedOut)
        {
            _logger.LogWarning("End Teams Bot SSO Prompt due to timeout");
            return await dc.EndDialogAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
        }
        else
        {
            if (IsTeamsVerificationInvoke(dc.Context) || IsTokenExchangeRequestInvoke(dc.Context))
            {
                // Recognize token
                PromptRecognizerResult<TeamsBotSsoPromptTokenResponse> recognized = await RecognizeTokenAsync(dc, cancellationToken).ConfigureAwait(false);

                if (recognized.Succeeded)
                {
                    return await dc.EndDialogAsync(recognized.Value, cancellationToken).ConfigureAwait(false);
                }
            }
            else if (isMessage && _settings.EndOnInvalidMessage)
            {
                _logger.LogWarning("End Teams Bot SSO Prompt due to invalid message");
                return await dc.EndDialogAsync(cancellationToken: cancellationToken).ConfigureAwait(false);
            }

            return Dialog.EndOfTurn;
        }
    }

    /// <summary>
    /// Shared implementation of the RecognizeTokenAsync function. This is intended for internal use, to
    /// consolidate the implementation of the OAuthPrompt and OAuthInput. Application logic should use
    /// those dialog classes.
    /// </summary>
    /// <param name="dc">DialogContext.</param>
    /// <param name="cancellationToken">CancellationToken.</param>
    /// <returns>PromptRecognizerResult.</returns>
    private async Task<PromptRecognizerResult<TeamsBotSsoPromptTokenResponse>> RecognizeTokenAsync(DialogContext dc, CancellationToken cancellationToken)
    {

        ITurnContext context = dc.Context;
        var result = new PromptRecognizerResult<TeamsBotSsoPromptTokenResponse>();
        TeamsBotSsoPromptTokenResponse tokenResponse = null;

        if (IsTokenExchangeRequestInvoke(context))
        {
            _logger.LogDebug("Receive token exchange request");

            var tokenResponseObject = context.Activity.Value as JObject;
            string ssoToken = tokenResponseObject?.ToObject<TokenExchangeInvokeRequest>().Token;
            // Received activity is not a token exchange request
            if (String.IsNullOrEmpty(ssoToken))
            {
                var warningMsg =
                  "The bot received an InvokeActivity that is missing a TokenExchangeInvokeRequest value. This is required to be sent with the InvokeActivity.";
                _logger.LogWarning(warningMsg);
                await SendInvokeResponseAsync(context, HttpStatusCode.BadRequest, warningMsg, cancellationToken).ConfigureAwait(false);
            }
            else
            {
                // TODO: Use ssoToken to construct obo credential and exchange access token for the given scope
                // 1. create oboCredential instance using ssoToken
                // 2. try use the oboCredential to get exchanged token for the given scopes
                AccessToken exchangedToken;
                try
                {
                    exchangedToken = new AccessToken("accessToken", new DateTimeOffset());
                    await SendInvokeResponseAsync(context, HttpStatusCode.OK, null, cancellationToken).ConfigureAwait(false);

                    var ssoTokenJwt = new JwtSecurityToken(ssoToken);
                    tokenResponse = new TeamsBotSsoPromptTokenResponse {
                        SsoToken = ssoToken,
                        SsoTokenExpiration = ssoTokenJwt.ValidTo.ToString(),
                        Token = exchangedToken.Token,
                        Expiration = exchangedToken.ExpiresOn.ToString()
                    };
                }
                catch (Exception e)
                {
                    var warningMsg = "The bot is unable to exchange token. Ask for user consent.";
                    _logger.LogInformation(warningMsg);
                    await SendInvokeResponseAsync(context, HttpStatusCode.PreconditionFailed, warningMsg, cancellationToken).ConfigureAwait(false);
                }

            }
        }
        else if (IsTeamsVerificationInvoke(context))
        {
            _logger.LogCritical("Receive Teams state verification request");
            await SendOAuthCardToObtainTokenAsync(context, cancellationToken).ConfigureAwait(false);
            await SendInvokeResponseAsync(context, HttpStatusCode.OK, null, cancellationToken).ConfigureAwait(false);
        }

        if (tokenResponse != null)
        {
            result.Succeeded = true;
            result.Value = tokenResponse;
        } else
        {
            result.Succeeded = false;
        }
        return result;
    }

    private static async Task SendInvokeResponseAsync(ITurnContext turnContext, HttpStatusCode statusCode, object body, CancellationToken cancellationToken)
    {
        await turnContext.SendActivityAsync(
            new Activity {
                Type = ActivityTypesEx.InvokeResponse,
                Value = new InvokeResponse {
                    Status = (int)statusCode,
                    Body = body,
                },
            }, cancellationToken).ConfigureAwait(false);
    }

    private bool IsTeamsVerificationInvoke(ITurnContext context) {
       return (context.Activity.Type == ActivityTypes.Message) && (context.Activity.Name == SignInConstants.VerifyStateOperationName);
    }
    private bool IsTokenExchangeRequestInvoke(ITurnContext context) {
        return (context.Activity.Type == ActivityTypes.Invoke) && (context.Activity.Name == SignInConstants.TokenExchangeOperationName);
    }

    /// <summary>
    /// Send OAuthCard that tells Teams to obtain an authentication token for the bot application.
    /// For details see https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/authentication/auth-aad-sso-bots.
    /// </summary>
    /// <param name="context">ITurnContext</param>
    /// <param name="cancellationToken">CancellationToken.</param>
    /// <returns>The task to await.</returns>
    private async Task SendOAuthCardToObtainTokenAsync(ITurnContext context, CancellationToken cancellationToken)
    {
        _logger.LogDebug("Send OAuth card to get SSO token");

        TeamsChannelAccount account = await TeamsInfo.GetMemberAsync(context, context.Activity.From.Id, cancellationToken).ConfigureAwait(false);
        _logger.LogDebug(
          "Get Teams member account user principal name: " + account.UserPrincipalName
        );

        string loginHint = account.UserPrincipalName ?? "";
        SignInResource signInResource = GetSignInResource(loginHint);

        // Ensure prompt initialized
        IMessageActivity prompt = Activity.CreateMessageActivity();
        //prompt.Id = context.Activity.ReplyToId;
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
                TokenExchangeResource = signInResource.TokenExchangeResource,
            },
        });
        // Send prompt
        await context.SendActivityAsync(prompt, cancellationToken).ConfigureAwait(false);
    }


    /// <summary>
    /// Get sign in authentication configuration
    /// </summary>
    /// <param name="loginHint"></param>
    /// <returns>sign in resource</returns>
    private SignInResource GetSignInResource(string loginHint)
    {
        _logger.LogDebug("Get sign in authentication configuration");
        string signInLink = $"bot-auth-start?scope={Uri.EscapeDataString(string.Join(" ", _settings.Scopes))}&clientId={_settings.ClientId}&tenantId={_settings.TenantId}&loginHint={loginHint}";
        _logger.LogDebug("Sign in link: " + signInLink);

        SignInResource signInResource = new SignInResource {
            SignInLink = signInLink,
            TokenExchangeResource = new TokenExchangeResource {
                Id = Guid.NewGuid().ToString(),
                Uri = Regex.Replace(_settings.ApplicationIdUri, @"/\/$/", "/access_as_user")
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
