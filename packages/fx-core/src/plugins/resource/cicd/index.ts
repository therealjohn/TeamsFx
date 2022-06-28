// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  err,
  UserError,
  SystemError,
  v2,
  TokenProvider,
  Inputs,
  Json,
  Func,
  ok,
  QTreeNode,
  ProjectSettings,
  SingleSelectQuestion,
  Platform,
  Stage,
  StaticOptions,
  MultiSelectQuestion,
  OptionItem,
  ContextV3,
} from "@microsoft/teamsfx-api";

import { FxResult, FxCICDPluginResultFactory as ResultFactory } from "./result";
import { CICDImpl } from "./plugin";
import { ErrorType, InternalError, NoProjectOpenedError, PluginError } from "./errors";
import { Alias, LifecycleFuncNames, PluginCICD } from "./constants";
import { Service } from "typedi";
import { ResourcePluginsV2 } from "../../solution/fx-solution/ResourcePluginContainer";
import {
  ResourcePlugin,
  Context,
  DeepReadonly,
  InputsWithProjectPath,
} from "@microsoft/teamsfx-api/build/v2";
import {
  githubOption,
  azdoOption,
  jenkinsOption,
  ciOption,
  cdOption,
  provisionOption,
  publishOption,
  questionNames,
} from "./questions";
import { Logger } from "./logger";
import { environmentManager } from "../../../core/environment";
import { telemetryHelper } from "./utils/telemetry-helper";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";
import { isExistingTabApp } from "../../../common";
import { NoCapabilityFoundError } from "../../../core/error";
import { ExistingTemplatesStat } from "./utils/existingTemplatesStat";
import { addCicdQuestion } from "../../../component/feature/cicd";

@Service(ResourcePluginsV2.CICDPlugin)
export class CICDPluginV2 implements ResourcePlugin {
  name = PluginCICD.PLUGIN_NAME;
  displayName = Alias.TEAMS_CICD_PLUGIN;
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  activate(projectSettings: ProjectSettings): boolean {
    return true;
  }

  public cicdImpl: CICDImpl = new CICDImpl();

  public async addCICDWorkflows(
    context: Context,
    inputs: Inputs,
    envInfo: v2.EnvInfoV2
  ): Promise<FxResult> {
    Logger.setLogger(context.logProvider);
    let envName = inputs[questionNames.Environment];
    // TODO: add support for VS/.Net Projects.
    if (inputs.platform === Platform.CLI) {
      // In CLI, get env name from the default `env` question.
      envName = envInfo.envName;
    }
    return await this.cicdImpl.addCICDWorkflows(context, inputs, envName);
  }

  public async getQuestionsForUserTask(
    ctx: Context,
    inputs: Inputs,
    func: Func,
    envInfo: DeepReadonly<v2.EnvInfoV2>,
    tokenProvider: TokenProvider
  ): Promise<FxResult> {
    return await addCicdQuestion(ctx as ContextV3, inputs as InputsWithProjectPath);
  }

  public async executeUserTask(
    ctx: Context,
    inputs: Inputs,
    func: Func,
    localSettings: Json,
    envInfo: v2.EnvInfoV2,
    tokenProvider: TokenProvider
  ): Promise<FxResult> {
    if (func.method === "addCICDWorkflows") {
      return await this.runWithExceptionCatching(
        ctx,
        envInfo,
        () => this.addCICDWorkflows(ctx, inputs, envInfo),
        true,
        LifecycleFuncNames.ADD_CICD_WORKFLOWS
      );
    }
    return ok(undefined);
  }

  private async runWithExceptionCatching(
    context: Context,
    envInfo: v2.EnvInfoV2,
    fn: () => Promise<FxResult>,
    sendTelemetry: boolean,
    name: string
  ): Promise<FxResult> {
    try {
      sendTelemetry &&
        telemetryHelper.sendStartEvent(context, envInfo, name, this.cicdImpl.commonProperties);
      const res: FxResult = await fn();
      sendTelemetry &&
        telemetryHelper.sendResultEvent(
          context,
          envInfo,
          name,
          res,
          this.cicdImpl.commonProperties
        );
      return res;
    } catch (e) {
      if (e instanceof UserError || e instanceof SystemError) {
        const res = err(e);
        sendTelemetry &&
          telemetryHelper.sendResultEvent(
            context,
            envInfo,
            name,
            res,
            this.cicdImpl.commonProperties
          );
        return res;
      }

      if (e instanceof PluginError) {
        const result =
          e.errorType === ErrorType.System
            ? ResultFactory.SystemError(
                e.name,
                [e.genDefaultMessage(), e.genMessage()],
                e.innerError
              )
            : ResultFactory.UserError(
                e.name,
                [e.genDefaultMessage(), e.genMessage()],
                e.showHelpLink,
                e.innerError
              );
        sendTelemetry &&
          telemetryHelper.sendResultEvent(
            context,
            envInfo,
            name,
            result,
            this.cicdImpl.commonProperties
          );
        return result;
      } else {
        // Unrecognized Exception.
        const UnhandledErrorCode = "UnhandledError";
        sendTelemetry &&
          telemetryHelper.sendResultEvent(
            context,
            envInfo,
            name,
            ResultFactory.SystemError(
              UnhandledErrorCode,
              [`Got an unhandled error: ${e.message}`, `Got an unhandled error: ${e.message}`],
              e.innerError
            ),
            this.cicdImpl.commonProperties
          );
        return ResultFactory.SystemError(UnhandledErrorCode, e.message, e);
      }
    }
  }
}

export default new CICDPluginV2();
