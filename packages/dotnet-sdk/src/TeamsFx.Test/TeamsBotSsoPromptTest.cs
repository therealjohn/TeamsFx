// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.TeamsFx.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Microsoft.Bot.Builder.Adapters;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;

namespace Microsoft.TeamsFx.Test;

[TestClass]
public class TeamsBotSsoPromptTest
{
    private static Mock<IOptions<AuthenticationOptions>> authOptionsMock;

    private static readonly string fakeClientId = "fake_client_id";
    private static readonly string fakeClientSecret = "fake_client_secret";
    private static readonly string fakeTenantId = "fake_tenant_id";
    private static readonly string fakeApplicationIdUri = "fake_application_id_url";
    private static readonly string fakeInitiateLoginEndpoint = "fake_initiate_login_endpoint";
    private static readonly string fakeOauthAuthority = "https://localhost";
    private static readonly string fakeDialogId = "MOCK_TEAMS_BOT_SSO_PROMPT";
    
    [ClassInitialize]
    public static void TestFixtureSetup(TestContext _)
    {
        // Executes once for the test class. (Optional)
        authOptionsMock = new Mock<IOptions<AuthenticationOptions>>();
        authOptionsMock.SetupGet(option => option.Value).Returns(
            new AuthenticationOptions()
            {
                ClientId = fakeClientId,
                ClientSecret = fakeClientSecret,
                OAuthAuthority = fakeOauthAuthority,
                TenantId = fakeTenantId,
                ApplicationIdUri = fakeApplicationIdUri,
                InitiateLoginEndpoint = fakeInitiateLoginEndpoint
            });
    }

    #region ConstructTeamsBotSsoPrompt
    [TestMethod]
    public void TeamsBotSsoPromptWithEmptyDialogIdShouldFail()
    {
        var loggerMock = new Mock<ILogger<TeamsBotSsoPrompt>>();
        Assert.ThrowsException<ArgumentNullException>(() => new TeamsBotSsoPrompt(string.Empty, new TeamsBotSsoPromptSettings(), loggerMock.Object, authOptionsMock.Object));
    }
    
    [TestMethod]
    public void TeamsBotSsoPromptWithEmptySettingShouldFail()
    {
        var loggerMock = new Mock<ILogger<TeamsBotSsoPrompt>>();
        Assert.ThrowsException<ArgumentNullException>(() => new TeamsBotSsoPrompt(fakeDialogId, null, loggerMock.Object,authOptionsMock.Object));
    }

    [TestMethod]
    public void TeamsBotSsoPromptWithEmptyAuthOptionShouldFail()
    {
        var loggerMock = new Mock<ILogger<TeamsBotSsoPrompt>>();
        Assert.ThrowsException<ArgumentNullException>(() => new TeamsBotSsoPrompt(fakeDialogId, new TeamsBotSsoPromptSettings(), loggerMock.Object, null));
    }
    #endregion

    #region BeginDialog
    [TestMethod]
    public async Task TeamsBotSsoPromptBeginDialogWithNoDialogContextShouldFail()
    {
        var loggerMock = new Mock<ILogger<TeamsBotSsoPrompt>>();
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () =>
        {
            var prompt = new TeamsBotSsoPrompt(fakeDialogId, new TeamsBotSsoPromptSettings(), loggerMock.Object, authOptionsMock.Object);
            await prompt.BeginDialogAsync(null);
        });
    }
    
    [TestMethod]
    public async Task TeamsBotSsoPromptBeginDialogNotInTeamsShouldFail()
    {
        var loggerMock = new Mock<ILogger<TeamsBotSsoPrompt>>();
        var ex = await Assert.ThrowsExceptionAsync<ExceptionWithCode>(async () =>
        {
            var prompt = new TeamsBotSsoPrompt(fakeDialogId, new TeamsBotSsoPromptSettings(), loggerMock.Object, authOptionsMock.Object);
            var convoState = new ConversationState(new MemoryStorage());
            var dialogState = convoState.CreateProperty<DialogState>("dialogState");

            var adapter = new TestAdapter()
                .Use(new AutoSaveStateMiddleware(convoState));

            // Create new DialogSet.
            var dialogs = new DialogSet(dialogState);
            dialogs.Add(prompt);

            var tc = new TurnContext(adapter, new Activity() { Type = ActivityTypes.Message, Conversation = new ConversationAccount() { Id = "123" }, ChannelId = "not-teams" });

            var dc = await dialogs.CreateContextAsync(tc);

            await prompt.BeginDialogAsync(dc);
        });
        Assert.AreEqual(ExceptionCode.ChannelNotSupported, ex.Code);
    }
    
    #endregion
}
