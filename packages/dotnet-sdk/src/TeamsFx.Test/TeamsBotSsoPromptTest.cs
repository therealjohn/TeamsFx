// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Extensions.Logging;
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
    private static TeamsBotSsoPromptSettings teamsBotSsoPromptSettingsMock;
    private static readonly string fakeClientId = Guid.NewGuid().ToString();
    private static readonly string fakeTenantId = Guid.NewGuid().ToString();
    private static readonly string fakeApplicationIdUri = "fake_application_id_url";
    private static readonly string fakeInitiateLoginEndpoint = "fake_initiate_login_endpoint";
    private static readonly string fakeDialogId = "MOCK_TEAMS_BOT_SSO_PROMPT";
    
    [ClassInitialize]
    public static void TestFixtureSetup(TestContext _)
    {
        // Executes once for the test class. (Optional)
        teamsBotSsoPromptSettingsMock = new TeamsBotSsoPromptSettings
        {
            ClientId = fakeClientId,
            TenantId = fakeTenantId,
            ApplicationIdUri = fakeApplicationIdUri,
            //InitiateLoginEndpoint = fakeInitiateLoginEndpoint,
            Scopes = new string[] { "User.Read" }
        };
    }

    #region ConstructTeamsBotSsoPrompt
    [TestMethod]
    public void TeamsBotSsoPromptWithEmptyDialogIdShouldFail()
    {
        var loggerMock = new Mock<ILogger<TeamsBotSsoPrompt>>();
        Assert.ThrowsException<ArgumentNullException>(() => new TeamsBotSsoPrompt(string.Empty, new TeamsBotSsoPromptSettings(), loggerMock.Object));
    }
    
    [TestMethod]
    public void TeamsBotSsoPromptWithEmptySettingShouldFail()
    {
        var loggerMock = new Mock<ILogger<TeamsBotSsoPrompt>>();
        Assert.ThrowsException<ArgumentNullException>(() => new TeamsBotSsoPrompt(fakeDialogId, null, loggerMock.Object));
    }

    [TestMethod]
    public void TeamsBotSsoPromptWithInvalidSettingShouldFail()
    {
        var loggerMock = new Mock<ILogger<TeamsBotSsoPrompt>>();
        var invalidTeamsBotSsoPromptSettings = new TeamsBotSsoPromptSettings
        {
            ClientId = String.Empty,
            TenantId = fakeTenantId,
            ApplicationIdUri = fakeApplicationIdUri,
            InitiateLoginEndpoint = fakeInitiateLoginEndpoint,
            Scopes = new string[] { "User.Read" }
        };
        var ex = Assert.ThrowsException<ExceptionWithCode>(() => new TeamsBotSsoPrompt(fakeDialogId, invalidTeamsBotSsoPromptSettings, loggerMock.Object));
        Assert.AreEqual(ExceptionCode.InvalidConfiguration, ex.Code);
        Assert.AreEqual("Teams bot sso prompt settings are missing or not correct with error: Client id is required. ", ex.Message);

        invalidTeamsBotSsoPromptSettings = new TeamsBotSsoPromptSettings
        {
            ClientId = fakeClientId,
            TenantId = "invalid-tenant-id-string",
            ApplicationIdUri = fakeApplicationIdUri,
            InitiateLoginEndpoint = fakeInitiateLoginEndpoint,
            Scopes = new string[] { "User.Read" }
        };
        ex = Assert.ThrowsException<ExceptionWithCode>(() => new TeamsBotSsoPrompt(fakeDialogId, invalidTeamsBotSsoPromptSettings, loggerMock.Object));
        Assert.AreEqual(ExceptionCode.InvalidConfiguration, ex.Code);
        StringAssert.StartsWith(ex.Message, "Teams bot sso prompt settings are missing or not correct with error: The field TenantId must match the regular expression");

        invalidTeamsBotSsoPromptSettings = new TeamsBotSsoPromptSettings
        {
            ClientId = fakeClientId,
            TenantId = fakeTenantId,
            InitiateLoginEndpoint = fakeInitiateLoginEndpoint,
            Scopes = new string[] { "User.Read" }
        };
        ex = Assert.ThrowsException<ExceptionWithCode>(() => new TeamsBotSsoPrompt(fakeDialogId, invalidTeamsBotSsoPromptSettings, loggerMock.Object));
        Assert.AreEqual(ExceptionCode.InvalidConfiguration, ex.Code);
        Assert.AreEqual("Teams bot sso prompt settings are missing or not correct with error: Application id uri is required. ", ex.Message);

        invalidTeamsBotSsoPromptSettings = new TeamsBotSsoPromptSettings
        {
            ClientId = fakeClientId,
            TenantId = fakeTenantId,
            ApplicationIdUri = fakeApplicationIdUri,
            Scopes = new string[] { "User.Read" }
        };
        ex = Assert.ThrowsException<ExceptionWithCode>(() => new TeamsBotSsoPrompt(fakeDialogId, invalidTeamsBotSsoPromptSettings, loggerMock.Object));
        Assert.AreEqual(ExceptionCode.InvalidConfiguration, ex.Code);
        Assert.AreEqual("Teams bot sso prompt settings are missing or not correct with error: Initiate login endpoint is required. ", ex.Message);

        invalidTeamsBotSsoPromptSettings = new TeamsBotSsoPromptSettings
        {
            ClientId = fakeClientId,
            TenantId = fakeTenantId,
            ApplicationIdUri = fakeApplicationIdUri,
            InitiateLoginEndpoint = fakeInitiateLoginEndpoint,
        };
        ex = Assert.ThrowsException<ExceptionWithCode>(() => new TeamsBotSsoPrompt(fakeDialogId, invalidTeamsBotSsoPromptSettings, loggerMock.Object));
        Assert.AreEqual(ExceptionCode.InvalidConfiguration, ex.Code);
        Assert.AreEqual("Teams bot sso prompt settings are missing or not correct with error: Scope is required. ", ex.Message);
    }

    #endregion

    #region BeginDialog
    [TestMethod]
    public async Task TeamsBotSsoPromptBeginDialogWithNoDialogContextShouldFail()
    {
        var loggerMock = new Mock<ILogger<TeamsBotSsoPrompt>>();
        await Assert.ThrowsExceptionAsync<ArgumentNullException>(async () =>
        {
            var prompt = new TeamsBotSsoPrompt(fakeDialogId, teamsBotSsoPromptSettingsMock, loggerMock.Object);
            await prompt.BeginDialogAsync(null);
        });
    }
    
    [TestMethod]
    public async Task TeamsBotSsoPromptBeginDialogNotInTeamsShouldFail()
    {
        var loggerMock = new Mock<ILogger<TeamsBotSsoPrompt>>();
        var ex = await Assert.ThrowsExceptionAsync<ExceptionWithCode>(async () =>
        {
            var prompt = new TeamsBotSsoPrompt(fakeDialogId, teamsBotSsoPromptSettingsMock, loggerMock.Object);
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


    //[TestMethod]
    //public async Task TeamsBotSsoPromptBeginDialogShouldSuccess()
    //{
    //    var loggerMock = new Mock<ILogger<TeamsBotSsoPrompt>>();
    //    var ex = await Assert.ThrowsExceptionAsync<ExceptionWithCode>(async () =>
    //    {
    //        var prompt = new TeamsBotSsoPrompt(fakeDialogId, teamsBotSsoPromptSettingsMock, loggerMock.Object);
    //        var convoState = new ConversationState(new MemoryStorage());
    //        var dialogState = convoState.CreateProperty<DialogState>("dialogState");

    //        var adapter = new TestAdapter()
    //            .Use(new AutoSaveStateMiddleware(convoState));

    //        // Create new DialogSet.
    //        var dialogs = new DialogSet(dialogState);
    //        dialogs.Add(prompt);

    //        var tc = new TurnContext(adapter, new Activity() { Type = ActivityTypes.Message, Conversation = new ConversationAccount() { Id = "123" }, ChannelId = Bot.Connector.Channels.Msteams });

    //        var dc = await dialogs.CreateContextAsync(tc);

    //        await prompt.BeginDialogAsync(dc);
    //    });
    //    Assert.AreEqual(ExceptionCode.ChannelNotSupported, ex.Code);
    //}

    #endregion
}
