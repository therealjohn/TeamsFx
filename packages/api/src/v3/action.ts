// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConsoleHttpPipelineLogger } from "@azure/core-http/types/latest/src/httpPipelineLogger";
import { pbkdf2 } from "crypto";
import { ok, Result } from "neverthrow";
import { v2, v3 } from "..";
import { Platform } from "../constants";
import { FxError } from "../error";
import { QTreeNode } from "../qm";
import { AzureSolutionSettings, Inputs } from "../types";
import { TokenProvider } from "../utils";

export type MaybePromise<T> = T | Promise<T>;

/**
 * Action is the basic concept to finish some lifecycle operation (create, provision, deploy, ...)
 * Action can be named action or anonymous action: named actions can be called by other actions, anonymous actions can not be called by other actions
 * An action can have the following types:
 * 1. shell - execute a shell script
 * 2. call - call an existing action
 * 3. function - run a javascript function
 * 4. group - a group of actions that can be executed in parallel or in sequence
 */
export type Action = GroupAction | CallAction | FunctionAction | ShellAction;
export enum ActionPriority {
  P0 = 0,
  P1 = 1,
  P2 = 2,
  P3 = 3,
  P4 = 4,
  P5 = 5,
  P6 = 6,
}
/**
 * group action: group action make it possible to leverage multiple sub-actions to accomplishment more complex task
 */
export interface GroupAction {
  name?: string;
  type: "group";
  /**
   * execution mode, in sequence or in parallel, if undefined, default is sequential
   */
  mode?: "sequential" | "parallel";
  actions: Action[];
  /**
   * execution priority in a sequential group, default is 3
   */
  priority?: ActionPriority;
}

/**
 * shell action: execute a shell script
 */
export interface ShellAction {
  name?: string;
  type: "shell";
  description: string;
  command: string;
  cwd?: string;
  async?: boolean;
  captureStdout?: boolean;
  captureStderr?: boolean;
  /**
   * execution priority in a sequential group, default is 3
   */
  priority?: ActionPriority;
}

/**
 * call action: call an existing action (defined locally or in other package)
 */
export interface CallAction {
  name?: string;
  type: "call";
  required: boolean; // required=true, throw error of target action does not exist; required=false, ignore this step if target action does not exist.
  targetAction: string;
  inputs?: {
    [k: string]: string;
  };
  /**
   * execution priority in a sequential group, default is 3
   */
  priority?: ActionPriority;
}

/**
 * function action: run a javascript function call that can do any kinds of work
 */
export interface FunctionAction {
  name?: string;
  type: "function";
  plan(context: any, inputs: Inputs): MaybePromise<string>;
  /**
   * question is to define inputs of the task
   */
  question?: (context: any, inputs: Inputs) => MaybePromise<Result<QTreeNode | undefined, FxError>>;
  /**
   * function body is a function that takes some context and inputs as parameter
   */
  execute: (context: any, inputs: Inputs) => MaybePromise<Result<any, FxError>>;
  /**
   * execution priority in a sequential group, default is 3
   */
  priority?: ActionPriority;
}

/**
 * a resource defines a collection of actions
 */
export interface Resource {
  readonly name: string;
  readonly description?: string;
  actions: (context: any) => MaybePromise<Action[]>;
}

/**
 * common function actions used in the built-in plugins
 */
export interface GenerateCodeAction extends FunctionAction {
  plan(context: v2.Context, inputs: Inputs): MaybePromise<string>;
  question?: (
    context: v2.Context,
    inputs: Inputs
  ) => MaybePromise<Result<QTreeNode | undefined, FxError>>;
  execute: (context: v2.Context, inputs: Inputs) => MaybePromise<Result<undefined, FxError>>;
}

export interface GenerateBicepAction extends FunctionAction {
  plan(context: v2.Context, inputs: Inputs): MaybePromise<string>;
  question?: (
    context: v2.Context,
    inputs: Inputs
  ) => MaybePromise<Result<QTreeNode | undefined, FxError>>;
  execute: (
    context: v2.Context,
    inputs: Inputs
  ) => MaybePromise<Result<v3.BicepTemplate[], FxError>>;
}

export interface ProvisionAction extends FunctionAction {
  plan(context: v2.Context, inputs: Inputs): MaybePromise<string>;
  question?: (
    context: v2.Context,
    inputs: Inputs
  ) => MaybePromise<Result<QTreeNode | undefined, FxError>>;
  execute: (
    context: { ctx: v2.Context; envInfo: v3.EnvInfoV3; tokenProvider: TokenProvider },
    inputs: Inputs
  ) => MaybePromise<Result<undefined, FxError>>;
}

export interface ConfigureAction extends FunctionAction {
  plan(context: v2.Context, inputs: Inputs): MaybePromise<string>;
  question?: (
    context: v2.Context,
    inputs: Inputs
  ) => MaybePromise<Result<QTreeNode | undefined, FxError>>;
  execute: (
    context: { ctx: v2.Context; envInfo: v3.EnvInfoV3; tokenProvider: TokenProvider },
    inputs: Inputs
  ) => MaybePromise<Result<undefined, FxError>>;
}

export interface BuildAction extends FunctionAction {
  plan(context: v2.Context, inputs: Inputs): MaybePromise<string>;
  question?: (
    context: v2.Context,
    inputs: Inputs
  ) => MaybePromise<Result<QTreeNode | undefined, FxError>>;
  execute: (context: v2.Context, inputs: Inputs) => MaybePromise<Result<undefined, FxError>>;
}

export interface DeployAction extends FunctionAction {
  plan(context: v2.Context, inputs: Inputs): MaybePromise<string>;
  question?: (
    context: v2.Context,
    inputs: Inputs
  ) => MaybePromise<Result<QTreeNode | undefined, FxError>>;
  execute: (
    context: { ctx: v2.Context; envInfo: v3.EnvInfoV3; tokenProvider: TokenProvider },
    inputs: Inputs
  ) => MaybePromise<Result<undefined, FxError>>;
}

export class AADResource implements Resource {
  name = "aad";
  actions(context: any): Action[] {
    const provision: ProvisionAction = {
      name: "aad.provision",
      type: "function",
      plan: (context: v2.Context, inputs: Inputs) => {
        inputs["aad.clientId"] = "mockClientId";
        inputs["aad.clientSecret"] = "mockSecret";
        return "provision aad app registration";
      },
      execute: async (
        context: { ctx: v2.Context; envInfo: v3.EnvInfoV3; tokenProvider: TokenProvider },
        inputs: Inputs
      ): Promise<Result<undefined, FxError>> => {
        inputs["aad.clientId"] = "mockClientId";
        inputs["aad.clientSecret"] = "mockSecret";
        return ok(undefined);
      },
    };
    const configure: ProvisionAction = {
      name: "aad.configure",
      type: "function",
      plan: (context: v2.Context, inputs: Inputs) => {
        return "configure aad app registration";
      },
      execute: async (
        context: { ctx: v2.Context; envInfo: v3.EnvInfoV3; tokenProvider: TokenProvider },
        inputs: Inputs
      ): Promise<Result<undefined, FxError>> => {
        return ok(undefined);
      },
    };
    return [provision, configure];
  }
}

export class AzureStorageResource implements Resource {
  name = "azure-storage";
  actions(context: any): Action[] {
    const generateBicep: GenerateBicepAction = {
      name: "azure-storage.configure",
      type: "function",
      plan: (context: v2.Context, inputs: Inputs) => {
        return "create azure storage bicep";
      },
      execute: async (
        context: v2.Context,
        inputs: Inputs
      ): Promise<Result<v3.BicepTemplate[], FxError>> => {
        return ok([]);
      },
    };
    const configure: ProvisionAction = {
      name: "azure-storage.configure",
      type: "function",
      plan: (context: v2.Context, inputs: Inputs) => {
        return "configure azure storage";
      },
      execute: async (
        context: { ctx: v2.Context; envInfo: v3.EnvInfoV3; tokenProvider: TokenProvider },
        inputs: Inputs
      ): Promise<Result<undefined, FxError>> => {
        return ok(undefined);
      },
    };
    return [generateBicep, configure];
  }
}

export class AzureWebAppResource implements Resource {
  name = "azure-web-app";
  actions(context: any): Action[] {
    const configure: ConfigureAction = {
      name: "azure-web-app.configure",
      type: "function",
      plan: (context: v2.Context, inputs: Inputs) => {
        return "configure azure web app";
      },
      execute: async (
        context: { ctx: v2.Context; envInfo: v3.EnvInfoV3; tokenProvider: TokenProvider },
        inputs: Inputs
      ): Promise<Result<undefined, FxError>> => {
        return ok(undefined);
      },
    };
    return [configure];
  }
}
export class AzureBotResource implements Resource {
  name = "azure-bot";
  actions(context: any): Action[] {
    const provision: ProvisionAction = {
      name: "azure-bot.provision",
      type: "function",
      plan: (context: v2.Context, inputs: Inputs) => {
        return "provision azure-bot (1.create AAD app for bot service; 2. create azure bot service)";
      },
      execute: async (
        context: { ctx: v2.Context; envInfo: v3.EnvInfoV3; tokenProvider: TokenProvider },
        inputs: Inputs
      ): Promise<Result<undefined, FxError>> => {
        inputs["azure-bot.botAadAppClientId"] = "MockBotAadAppClientId";
        inputs["azure-bot.botId"] = "MockBotId";
        return ok(undefined);
      },
    };
    return [provision];
  }
}

export class TeamsManifestResource implements Resource {
  name = "teams-manifest";
  async actions(context: any): Promise<Action[]> {
    const init: GenerateCodeAction = {
      name: "teams-manifest.init",
      type: "function",
      plan: (context: v2.Context, inputs: Inputs) => {
        return "create a new manifest";
      },
      execute: async (context: v2.Context, inputs: Inputs) => {
        return ok(undefined);
      },
    };
    const addCapability: GenerateCodeAction = {
      name: "teams-manifest.addCapability",
      type: "function",
      plan: (context: v2.Context, inputs: Inputs) => {
        return `add capability in teams manifest: ${inputs["add-capability"]}`;
      },
      execute: async (context: v2.Context, inputs: Inputs) => {
        return ok(undefined);
      },
    };
    const provision: ProvisionAction = {
      name: "teams-manifest.provision",
      type: "function",
      plan: (context: v2.Context, inputs: Inputs) => {
        return "provision teams manifest";
      },
      execute: async (
        context: { ctx: v2.Context; envInfo: v3.EnvInfoV3; tokenProvider: TokenProvider },
        inputs: Inputs
      ) => {
        console.log(
          `provision teams manifest with tab:${inputs["tab.endpoint"]} and bot:${inputs["azure-bot.botId"]}`
        );
        return ok(undefined);
      },
    };
    return [init, addCapability, provision];
  }
}

export class TeamsfxSolutionResource implements Resource {
  name = "teamsfx-solution";
  async deployBicep(context: v2.Context): Promise<Action> {
    return {
      type: "function",
      name: "teamsfx-solution.deployBicep",
      plan: (context: v2.Context, inputs: Inputs) => {
        return "deploy bicep";
      },
      execute: async (context: v2.Context, inputs: Inputs) => {
        console.log("deploy bicep");
        inputs["tab.endpoint"] = "MockTabEndpoint";
        return ok(undefined);
      },
    };
  }
  async preProvision(context: v2.Context): Promise<Action> {
    return {
      type: "function",
      name: "teamsfx-solution.preProvision",
      plan: (context: v2.Context, inputs: Inputs) => {
        return "preProvision: check common configs (account, resource group)";
      },
      execute: async (context: v2.Context, inputs: Inputs) => {
        console.log("preProvision: check common configs (account, resource group)");
        return ok(undefined);
      },
    };
  }
  async provision(context: v2.Context): Promise<Action> {
    const solutionSetting = context.projectSetting.solutionSettings as AzureSolutionSettings;
    const provisionActions: Action[] = solutionSetting.activeResourcePlugins
      .filter((p) => p !== "azure-bot")
      .map((p) => {
        return {
          type: "call",
          required: false,
          targetAction: `${p}.provision`,
        };
      });
    const configureActions: Action[] = solutionSetting.activeResourcePlugins.map((p) => {
      return {
        type: "call",
        required: false,
        targetAction: `${p}.configure`,
      };
    });
    const provisionSequences: Action[] = [
      {
        type: "call",
        required: false,
        targetAction: "teamsfx-solution.preProvision",
      },
      {
        type: "group",
        mode: "parallel",
        actions: provisionActions,
      },
      {
        type: "call",
        required: true,
        targetAction: "teamsfx-solution.deployBicep",
      },
    ];
    if (solutionSetting.activeResourcePlugins.includes("azure-bot")) {
      provisionSequences.push({
        type: "call",
        required: false,
        targetAction: "azure-bot.provision",
      });
    }
    provisionSequences.push({
      type: "function",
      plan: (context: v2.Context, inputs: Inputs) => {
        return "set configuration after bicep deployment";
      },
      execute: async (context: any, inputs: Inputs) => {
        return ok(undefined);
      },
    });
    provisionSequences.push({
      type: "group",
      mode: "parallel",
      actions: configureActions,
    });
    provisionSequences.push({
      type: "call",
      required: true,
      targetAction: "teams-manifest.provision",
    });
    return {
      name: "teamsfx-solution.provision",
      type: "group",
      actions: provisionSequences,
    };
  }
  async actions(context: any): Promise<Action[]> {
    return [
      await this.deployBicep(context),
      await this.provision(context),
      await this.preProvision(context),
    ];
  }
}

function getActionPriority(action: Action, actions: Map<string, Action>): ActionPriority {
  if (action.priority) return action.priority;
  if (action.type === "call") {
    const targetAction = actions.get(action.targetAction);
    if (targetAction && targetAction.priority) return targetAction.priority;
  }
  return ActionPriority.P3;
}

async function planAction(
  context: any,
  inputs: Inputs,
  action: Action,
  actions: Map<string, Action>
) {
  if (action.type === "function") {
    console.log("plan:" + (await action.plan(context, inputs)));
  } else if (action.type === "shell") {
    console.log("plan: shell " + action.command);
  } else if (action.type === "call") {
    const targetAction = actions.get(action.targetAction);
    if (action.required && !targetAction) {
      throw new Error("targetAction not exist: " + action.targetAction);
    }
    if (targetAction) {
      if (action.inputs) {
        for (const target of Object.keys(action.inputs)) {
          const source = action.inputs[target];
          inputs[target] = inputs[source];
        }
      }
      planAction(context, inputs, targetAction, actions);
    }
  } else {
    if (!action.mode || action.mode === "sequential") {
      action.actions = action.actions.sort((a1, a2) => {
        const p1 = getActionPriority(a1, actions);
        const p2 = getActionPriority(a2, actions);
        return p1 - p2;
      });
    }
    for (const act of action.actions) {
      await planAction(context, inputs, act, actions);
    }
  }
}

async function executeAction(
  context: any,
  inputs: Inputs,
  action: Action,
  actions: Map<string, Action>
) {
  if (action.type === "function") {
    console.log("execute:" + action.name);
    await action.execute(context, inputs);
  } else if (action.type === "shell") {
    console.log("shell:" + action.command);
  } else if (action.type === "call") {
    const targetAction = actions.get(action.targetAction);
    if (action.required && !targetAction) {
      throw new Error("action not exist: " + action.targetAction);
    }
    if (targetAction) {
      if (action.inputs) {
        for (const target of Object.keys(action.inputs)) {
          const source = action.inputs[target];
          inputs[target] = inputs[source];
        }
      }
      await executeAction(context, inputs, targetAction, actions);
    }
  } else {
    for (const act of action.actions) {
      await executeAction(context, inputs, act, actions);
    }
  }
}

async function test() {
  const actionMap = new Map<string, Action>();
  const context = {
    projectSetting: {
      appName: "huajie0316",
      solutionSettings: {
        activeResourcePlugins: ["aad", "azure-storage", "azure-web-app", "azure-bot"],
      },
    },
  };
  const resources = [
    new TeamsfxSolutionResource(),
    new AADResource(),
    new TeamsManifestResource(),
    new AzureBotResource(),
    new AzureWebAppResource(),
    new AzureStorageResource(),
  ];
  for (const resource of resources) {
    const actions = await resource.actions(context);
    actions.forEach((action) => {
      if (action.name) {
        actionMap.set(action.name, action);
      }
    });
  }
  const rootAction = actionMap.get("teamsfx-solution.provision") as Action;
  console.log(JSON.stringify(rootAction));
  const inputs: Inputs = { platform: Platform.VSCode };
  await planAction(context, inputs, rootAction, actionMap);
  // await executeAction(context, inputs, rootAction, actionMap);
}

test();
