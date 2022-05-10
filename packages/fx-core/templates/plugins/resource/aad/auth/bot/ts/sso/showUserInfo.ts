// This file implements a function to call Graph API with TeamsFx SDK to get user profile with SSO token.

import { createMicrosoftGraphClient, TeamsFx } from "@microsoft/teamsfx";
import { TurnContext } from "botbuilder";
import { DialogTurnResult } from "botbuilder-dialogs";

// If you need extra parameters, you can include the parameters in `addCommand`as parameter
export async function showUserInfo(
  context: TurnContext,
  ssoToken: string
): Promise<DialogTurnResult> {
  await context.sendActivity("Retrieving user information from Microsoft Graph ...");

  // Init TeamsFx instance with SSO token
  const teamsfx = new TeamsFx().setSsoToken(ssoToken);
  const graphClient = createMicrosoftGraphClient(teamsfx, ["User.Read"]);
  const me = await graphClient.api("/me").get();
  if (me) {
    await context.sendActivity(
      `You're logged in as ${me.displayName} (${me.userPrincipalName})${
        me.jobTitle ? `; your job title is: ${me.jobTitle}` : ""
      }.`
    );
  } else {
    await context.sendActivity("Could not retrieve profile information from Microsoft Graph.");
  }

  return;
}
