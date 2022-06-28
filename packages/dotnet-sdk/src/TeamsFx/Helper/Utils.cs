// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
namespace Microsoft.TeamsFx.Helper;

using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using System.IdentityModel.Tokens.Jwt;

static internal class Utils
{
    static internal JwtSecurityToken ParseJwt(string token)
    {
        if (string.IsNullOrEmpty(token))
        {
            throw new ExceptionWithCode("SSO token is null or empty.", ExceptionCode.InvalidParameter);
        }
        var handler = new JwtSecurityTokenHandler();
        try
        {
            var jsonToken = handler.ReadToken(token);
            if (jsonToken is not JwtSecurityToken tokenS || string.IsNullOrEmpty(tokenS.Payload["exp"].ToString()))
            {
                throw new ExceptionWithCode("Decoded token is null or exp claim does not exists.", ExceptionCode.InternalError);
            }
            return tokenS;
        }
        catch (ArgumentException e)
        {
            var errorMessage = $"Parse jwt token failed with error: {e.Message}";
            throw new ExceptionWithCode(errorMessage, ExceptionCode.InternalError);
        }
    }

    static internal string GetCacheKey(string token, string scopes, string clientId)
    {
        var parsedJwt = ParseJwt(token);
        var userObjectId = parsedJwt.Payload["oid"].ToString();
        var tenantId = parsedJwt.Payload["tid"].ToString();

        var key = string.Join("-",
            new string[] { "accessToken", userObjectId, clientId, tenantId, scopes }).Replace(' ', '_');
        return key;
    }

    static internal UserInfo GetUserInfoFromSsoToken(string ssoToken)
    {
        var tokenObject = ParseJwt(ssoToken);

        var userInfo = new UserInfo() {
            DisplayName = tokenObject.Payload["name"].ToString(),
            ObjectId = tokenObject.Payload["oid"].ToString(),
            PreferredUserName = "",
        };

        var version = tokenObject.Payload["ver"].ToString();

        if (version == "2.0")
        {
            userInfo.PreferredUserName = tokenObject.Payload["preferred_username"].ToString();
        }
        else if (version == "1.0")
        {
            userInfo.PreferredUserName = tokenObject.Payload["upn"].ToString();
        }
        return userInfo;
    }

    static internal async ValueTask<global::Azure.Core.AccessToken> GetAccessTokenByOnBehalfOfFlow(string ssoToken, IEnumerable<string> scopes, IIdentityClientAdapter identityClientAdapter, ILogger logger)
    {
        logger.LogTrace($"Get access token from authentication server with scopes: {string.Join(' ', scopes)}");

        try
        {
            logger.LogDebug("Acquiring token via on-behalf-of flow.");
            var result = await identityClientAdapter
                .GetAccessToken(ssoToken, scopes)
                .ConfigureAwait(false);

            var accessToken = new global::Azure.Core.AccessToken(result.AccessToken, result.ExpiresOn);
            return accessToken;
        }
        catch (MsalUiRequiredException) // Need user interaction
        {
            var fullErrorMsg = $"Failed to get access token from OAuth identity server, please login(consent) first";
            logger.LogWarning(fullErrorMsg);
            throw new ExceptionWithCode(fullErrorMsg, ExceptionCode.UiRequiredError);
        }
        catch (MsalServiceException ex) // Errors that returned from AAD service
        {
            var fullErrorMsg = $"Failed to get access token from OAuth identity server with error: {ex.ResponseBody}";
            logger.LogWarning(fullErrorMsg);
            throw new ExceptionWithCode(fullErrorMsg, ExceptionCode.ServiceError);
        }
        catch (MsalClientException ex) // Exceptions that are local to the MSAL library
        {
            var fullErrorMsg = $"Failed to get access token with error: {ex.Message}";
            logger.LogWarning(fullErrorMsg);
            throw new ExceptionWithCode(fullErrorMsg, ExceptionCode.InternalError);
        }
    }
}
