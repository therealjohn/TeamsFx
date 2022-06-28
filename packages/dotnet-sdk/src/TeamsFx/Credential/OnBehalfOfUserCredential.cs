// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Azure.Core;
using Microsoft.Extensions.Logging;
using Microsoft.TeamsFx.Helper;
using System.IdentityModel.Tokens.Jwt;
using AccessToken = Microsoft.TeamsFx.Model.AccessToken;

namespace Microsoft.TeamsFx.Credential;

/// <summary>
/// Represent on-behalf-of flow to get user identity, and it is designed to be used in server side.
/// </summary>
/// <remarks>
/// Can only be used in server side.
/// </remarks>
public class OnBehalfOfUserCredential : TokenCredential
{
    private IIdentityClientAdapter _identityClientAdapter;
    private AccessToken _ssoToken;

    #region Util
    private readonly ILogger<OnBehalfOfUserCredential> _logger;
    #endregion

    /// <summary>
    /// Constructor of OnBehalfOfUserCredential
    /// </summary>
    /// <param name="logger">Logger of OnBehalfOfUserCredential Class.</param>
    /// <param name="identityClientAdapter">Global instance of adaptor to call MSAL.NET library</param>
    public OnBehalfOfUserCredential(IIdentityClientAdapter identityClientAdapter, ILogger<OnBehalfOfUserCredential> logger)
    {
        _identityClientAdapter = identityClientAdapter;
        _logger = logger;
    }

    /// <summary>
    /// Set SSO token of the on-behalf-of credential.
    /// </summary>
    /// <param name="ssoToken">Single Sign On(SSO) token</param>
    public void SetSsoToken(string ssoToken)
    {
        var ssoTokenJwt = new JwtSecurityToken(ssoToken);
        _ssoToken = new AccessToken(ssoToken, ssoTokenJwt.ValidTo);
    }

    /// <summary>
    /// Sync version is not supported now. Please use GetTokenAsync instead.
    /// </summary>
    public override global::Azure.Core.AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
    {
        throw new NotImplementedException();
    }



    /// <summary>
    /// Gets an Azure.Core.AccessToken for the specified set of scopes.
    /// </summary>
    /// <param name="requestContext">The Azure.Core.TokenRequestContext with authentication information.</param>
    /// <param name="cancellationToken">The System.Threading.CancellationToken to use.</param>
    /// <returns>A valid Azure.Core.AccessToken with expected scopes..</returns>
    ///
    /// <exception cref="ExceptionCode.InternalError">When failed to login with unknown error.</exception>
    /// <exception cref="ExceptionCode.UiRequiredError">When need user consent to get access token.</exception>
    /// <exception cref="ExceptionCode.ServiceError">When failed to get access token from identity server(AAD).</exception>
    ///
    /// <example>
    /// For example:
    /// <code>
    /// // Get Graph access token for single scope
    /// await OnBehalfOfUserCredential.GetTokenAsync(new TokenRequestContext(new string[] { "User.Read" }), new CancellationToken());
    /// // Get Graph access token for multiple scopes
    /// await OnBehalfOfUserCredential.GetTokenAsync(new TokenRequestContext(new string[] { "User.Read", "Application.Read.All" }), new CancellationToken());
    /// </code>
    /// </example>
    /// <remarks>
    /// Can only be used within Teams.
    /// If scopes is empty string or array, it returns SSO token.
    /// If scopes is non-empty, it returns access token for target scope.
    /// </remarks>
    public override async ValueTask<global::Azure.Core.AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
    {
        _logger.LogInformation("OnBehalfOfUserCredential get access token async.");
        
        var scopes = requestContext.Scopes;
        if (scopes == null || scopes.Length == 0)
        {
            _logger.LogInformation("Scopes is null or empty. Return sso token directly.");
            return _ssoToken.ToAzureAccessToken();
        }
        else
        {
            if (_ssoToken == null)
            {
                var errorMsg = "SSO token of the on-behalf-of credential has not yet been set. Please call SetSsoToken() first";
                _logger.LogError(errorMsg);
                throw new ExceptionWithCode(errorMsg, ExceptionCode.SsoTokenNotSet);
            }
            else
            {
                _logger.LogInformation($"Get access token with scopes: {string.Join(' ', scopes)}");
                var accessToken = await Utils.GetAccessTokenByOnBehalfOfFlow(_ssoToken.Token, scopes, _identityClientAdapter, _logger).ConfigureAwait(false);
                return accessToken;
            }
        }
    }
}
