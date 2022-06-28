// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";
import * as path from "path";
import * as fs from "fs-extra";
import {
  AzureSolutionSettings,
  Inputs,
  QTreeNode,
  SystemError,
  UserError,
  ok,
  Platform,
  FxError,
  ProjectSettingsV3,
} from "@microsoft/teamsfx-api";
import { Context } from "@microsoft/teamsfx-api/build/v2";
import {
  generateTempFolder,
  copyFileIfExist,
  removeFileIfExist,
  getSampleFileName,
  checkInputEmpty,
  Notification,
} from "./utils";
import {
  ApiConnectorConfiguration,
  AuthConfig,
  BasicAuthConfig,
  AADAuthConfig,
  APIKeyAuthConfig,
} from "./config";
import { ApiConnectorResult, ResultFactory, QuestionResult, FileChange } from "./result";
import { AuthType, Constants, KeyLocation, ComponentType } from "./constants";
import { EnvHandler } from "./envHandler";
import { ErrorMessage } from "./errors";
import { ResourcePlugins } from "../../../common/constants";
import {
  ApiNameQuestion,
  basicAuthUsernameQuestion,
  botOption,
  functionOption,
  apiEndpointQuestion,
  BasicAuthOption,
  CertAuthOption,
  AADAuthOption,
  APIKeyAuthOption,
  ImplementMyselfOption,
  reuseAppOption,
  anotherAppOption,
  appTenantIdQuestion,
  appIdQuestion,
  requestHeaderOption,
  queryParamsOption,
  buildAPIKeyNameQuestion,
} from "./questions";
import { getLocalizedString } from "../../../common/localizeUtils";
import { SampleHandler } from "./sampleHandler";
import { isAADEnabled } from "../../../common";
import { getAzureSolutionSettings } from "../../solution/fx-solution/v2/utils";
import { DepsHandler } from "./depsHandler";
import { checkEmptySelect } from "./checker";
import { Telemetry, TelemetryUtils } from "./telemetry";
import { isV3 } from "../../../core";
import { hasAAD, hasBot, hasFunction } from "../../../common/projectSettingsHelperV3";
export class ApiConnectorImpl {
  public async scaffold(ctx: Context, inputs: Inputs): Promise<ApiConnectorResult> {
    if (!inputs.projectPath) {
      throw ResultFactory.UserError(
        ErrorMessage.InvalidProjectError.name,
        ErrorMessage.InvalidProjectError.message()
      );
    }
    const projectPath = inputs.projectPath;
    const languageType: string = ctx.projectSetting!.programmingLanguage!;
    const config: ApiConnectorConfiguration = this.getUserDataFromInputs(inputs);

    const telemetryProperties = this.getTelemetryProperties(config);

    TelemetryUtils.init(ctx.telemetryReporter);
    TelemetryUtils.sendEvent(
      Telemetry.stage.scaffold + Telemetry.startSuffix,
      undefined,
      telemetryProperties
    );
    // CLI checker
    const bot = isV3()
      ? hasBot(ctx.projectSetting as ProjectSettingsV3)
      : (
          ctx.projectSetting.solutionSettings as AzureSolutionSettings
        )?.activeResourcePlugins?.includes(ResourcePlugins.Bot);
    const hasFunc = isV3()
      ? hasFunction(ctx.projectSetting as ProjectSettingsV3)
      : (
          ctx.projectSetting.solutionSettings as AzureSolutionSettings
        )?.activeResourcePlugins?.includes(ResourcePlugins.Function);
    if (!bot && config.ComponentType.includes(ComponentType.BOT)) {
      throw ResultFactory.UserError(
        ErrorMessage.componentNotExistError.name,
        ErrorMessage.componentNotExistError.message(ComponentType.BOT)
      );
    }
    if (!hasFunc && config.ComponentType.includes(ComponentType.API)) {
      throw ResultFactory.UserError(
        ErrorMessage.componentNotExistError.name,
        ErrorMessage.componentNotExistError.message(ComponentType.API)
      );
    }

    // backup relative files.
    const backupFolderName = generateTempFolder();
    await Promise.all(
      config.ComponentType.map(async (component) => {
        await this.backupExistingFiles(path.join(projectPath, component), backupFolderName);
      })
    );

    try {
      let filesChanged: FileChange[] = [];
      await Promise.all(
        config.ComponentType.map(async (component) => {
          const changes = await this.scaffoldInComponent(
            projectPath,
            component,
            config,
            languageType
          );
          filesChanged = filesChanged.concat(changes);
        })
      );
      const msg: string = Notification.getNotificationMsg(config, languageType);
      const logMessage = getLocalizedString(
        "plugins.apiConnector.Log.CommandSuccess",
        filesChanged.reduce(
          (previousValue, currentValue) =>
            previousValue +
            `[${currentValue.changeType}] ${path.relative(
              inputs.projectPath!,
              currentValue.filePath
            )}` +
            "\n",
          ""
        )
      );
      ctx.logProvider?.info(logMessage); // Print generated/updated files for users

      if (inputs.platform != Platform.CLI) {
        ctx.userInteraction
          ?.showMessage("info", msg, false, "OK", Notification.READ_MORE)
          .then((result) => {
            const userSelected = result.isOk() ? result.value : undefined;
            if (userSelected === Notification.READ_MORE) {
              ctx.userInteraction?.openUrl(Notification.READ_MORE_URL);
            }
          });
      } else {
        ctx.userInteraction.showMessage(
          "info",
          msg + ` ${Notification.GetLinkNotification()}`,
          false
        );
      }
    } catch (err) {
      await Promise.all(
        config.ComponentType.map(async (component) => {
          await fs.copy(
            path.join(projectPath, component, backupFolderName),
            path.join(projectPath, component),
            { overwrite: true }
          );
          await this.removeSampleFilesWhenRestore(
            projectPath,
            component,
            config.APIName,
            languageType
          );
        })
      );
      if (!(err instanceof SystemError) && !(err instanceof UserError)) {
        err = ResultFactory.SystemError(
          ErrorMessage.generateApiConFilesError.name,
          ErrorMessage.generateApiConFilesError.message(err.message)
        );
      }
      this.sendErrorTelemetry(err as FxError);
      throw err;
    } finally {
      await Promise.all(
        config.ComponentType.map(async (component) => {
          await removeFileIfExist(path.join(projectPath, component, backupFolderName));
        })
      );
    }
    TelemetryUtils.sendEvent(Telemetry.stage.scaffold, true, telemetryProperties);
    const result = config.ComponentType.map((item) => {
      return path.join(projectPath, item, getSampleFileName(config.APIName, languageType));
    });
    return { generatedFiles: result };
  }

  private sendErrorTelemetry(thrownErr: FxError) {
    const errorCode = thrownErr.source + "." + thrownErr.name;
    const errorType =
      thrownErr instanceof SystemError ? Telemetry.systemError : Telemetry.userError;
    const errorMessage = thrownErr.message;
    TelemetryUtils.sendErrorEvent(Telemetry.stage.scaffold, errorCode, errorType, errorMessage);
    return thrownErr;
  }

  private async scaffoldInComponent(
    projectPath: string,
    componentItem: string,
    config: ApiConnectorConfiguration,
    languageType: string
  ): Promise<FileChange[]> {
    const updatedPackageFile = await this.addSDKDependency(projectPath, componentItem);
    const updatedEnvFile = await this.scaffoldEnvFileToComponent(
      projectPath,
      config,
      componentItem
    );
    const generatedSampleFile = await this.scaffoldSampleCodeToComponent(
      projectPath,
      config,
      componentItem,
      languageType
    );

    const fileChanges: FileChange[] = [updatedEnvFile, generatedSampleFile];
    if (updatedPackageFile) {
      // if we didn't update package.json, the result will be undefined
      fileChanges.push(updatedPackageFile);
    }
    return fileChanges;
  }

  private async backupExistingFiles(folderPath: string, backupFolder: string) {
    await fs.ensureDir(path.join(folderPath, backupFolder));
    await copyFileIfExist(
      path.join(folderPath, Constants.envFileName),
      path.join(folderPath, backupFolder, Constants.envFileName)
    );
    await copyFileIfExist(
      path.join(folderPath, Constants.pkgJsonFile),
      path.join(folderPath, backupFolder, Constants.pkgJsonFile)
    );
    await copyFileIfExist(
      path.join(folderPath, Constants.pkgLockFile),
      path.join(folderPath, backupFolder, Constants.pkgLockFile)
    );
  }

  private async removeSampleFilesWhenRestore(
    projectPath: string,
    component: string,
    apiName: string,
    languageType: string
  ) {
    const apiFileName = getSampleFileName(apiName, languageType);
    const sampleFilePath = path.join(projectPath, component, apiFileName);
    await removeFileIfExist(sampleFilePath);
  }

  private getAuthConfigFromInputs(inputs: Inputs): AuthConfig {
    let config: AuthConfig;
    const apiType = inputs[Constants.questionKey.apiType];
    switch (apiType) {
      case AuthType.BASIC:
        checkInputEmpty(inputs, Constants.questionKey.apiUserName);
        config = {
          AuthType: AuthType.BASIC,
          UserName: inputs[Constants.questionKey.apiUserName],
        } as BasicAuthConfig;
        break;
      case AuthType.AAD:
        const AADConfig = {
          AuthType: AuthType.AAD,
        } as AADAuthConfig;
        if (inputs[Constants.questionKey.apiAppType] === reuseAppOption.id) {
          AADConfig.ReuseTeamsApp = true;
        } else {
          AADConfig.ReuseTeamsApp = false;
          checkInputEmpty(
            inputs,
            Constants.questionKey.apiAppTenentId,
            Constants.questionKey.apiAppId
          );
          AADConfig.TenantId = inputs[Constants.questionKey.apiAppTenentId];
          AADConfig.ClientId = inputs[Constants.questionKey.apiAppId];
        }
        config = AADConfig;
        break;
      case AuthType.APIKEY:
        const APIKeyConfig = {
          AuthType: AuthType.APIKEY,
        } as APIKeyAuthConfig;
        if (inputs[Constants.questionKey.apiAPIKeyLocation] === requestHeaderOption.id) {
          APIKeyConfig.Location = KeyLocation.Header;
        } else {
          APIKeyConfig.Location = KeyLocation.QueryParams;
        }
        checkInputEmpty(inputs, Constants.questionKey.apiAPIKeyName);
        APIKeyConfig.Name = inputs[Constants.questionKey.apiAPIKeyName];
        config = APIKeyConfig;
        break;
      case AuthType.CUSTOM:
      case AuthType.CERT:
        config = {
          AuthType: apiType,
        };
        break;
      default:
        throw ResultFactory.SystemError(
          ErrorMessage.ApiConnectorInputError.name,
          ErrorMessage.ApiConnectorInputError.message(inputs[Constants.questionKey.apiAppType])
        );
    }
    return config;
  }

  private getUserDataFromInputs(inputs: Inputs): ApiConnectorConfiguration {
    checkInputEmpty(
      inputs,
      Constants.questionKey.componentsSelect,
      Constants.questionKey.apiName,
      Constants.questionKey.endpoint
    );
    const authConfig = this.getAuthConfigFromInputs(inputs);
    const config: ApiConnectorConfiguration = {
      ComponentType: inputs[Constants.questionKey.componentsSelect],
      APIName: inputs[Constants.questionKey.apiName],
      AuthConfig: authConfig,
      EndPoint: inputs[Constants.questionKey.endpoint],
    };
    return config;
  }

  private async scaffoldEnvFileToComponent(
    projectPath: string,
    config: ApiConnectorConfiguration,
    component: string
  ): Promise<FileChange> {
    const envHander = new EnvHandler(projectPath, component);
    envHander.updateEnvs(config);
    return await envHander.saveLocalEnvFile();
  }

  private async scaffoldSampleCodeToComponent(
    projectPath: string,
    config: ApiConnectorConfiguration,
    component: string,
    languageType: string
  ): Promise<FileChange> {
    const sampleHandler = new SampleHandler(projectPath, languageType, component);
    return await sampleHandler.generateSampleCode(config);
  }

  private async addSDKDependency(
    projectPath: string,
    component: string
  ): Promise<FileChange | undefined> {
    const depsHandler: DepsHandler = new DepsHandler(projectPath, component);
    return await depsHandler.addPkgDeps();
  }

  public async generateQuestion(ctx: Context, inputs: Inputs): Promise<QuestionResult> {
    const componentOptions = [];
    if (inputs.platform === Platform.CLI_HELP) {
      componentOptions.push(botOption);
      componentOptions.push(functionOption);
    } else {
      if (!isV3()) {
        const activePlugins = (ctx.projectSetting.solutionSettings as AzureSolutionSettings)
          ?.activeResourcePlugins;
        if (!activePlugins) {
          throw ResultFactory.UserError(
            ErrorMessage.NoActivePluginsExistError.name,
            ErrorMessage.NoActivePluginsExistError.message()
          );
        }
        if (activePlugins.includes(ResourcePlugins.Bot)) {
          componentOptions.push(botOption);
        }
        if (activePlugins.includes(ResourcePlugins.Function)) {
          componentOptions.push(functionOption);
        }
      } else {
        if (hasBot(ctx.projectSetting as ProjectSettingsV3)) {
          componentOptions.push(botOption);
        }
        if (hasFunction(ctx.projectSetting as ProjectSettingsV3)) {
          componentOptions.push(functionOption);
        }
      }
      if (componentOptions.length === 0) {
        throw ResultFactory.UserError(
          ErrorMessage.NoValidCompoentExistError.name,
          ErrorMessage.NoValidCompoentExistError.message()
        );
      }
    }
    const whichComponent = new QTreeNode({
      name: Constants.questionKey.componentsSelect,
      type: "multiSelect",
      staticOptions: componentOptions,
      title: getLocalizedString("plugins.apiConnector.whichService.title"),
      validation: {
        validFunc: checkEmptySelect,
      },
      placeholder: getLocalizedString("plugins.apiConnector.whichService.placeholder"), // Use the placeholder to display some description
    });
    const apiNameQuestion = new ApiNameQuestion(ctx);
    const whichAuthType = this.buildAuthTypeQuestion(ctx, inputs);
    const question = new QTreeNode({
      type: "group",
    });
    question.addChild(new QTreeNode(apiEndpointQuestion));
    question.addChild(whichComponent);
    question.addChild(new QTreeNode(apiNameQuestion.getQuestion()));
    question.addChild(whichAuthType);

    return ok(question);
  }

  public buildAuthTypeQuestion(ctx: Context, inputs: Inputs): QTreeNode {
    const whichAuthType = new QTreeNode({
      name: Constants.questionKey.apiType,
      type: "singleSelect",
      staticOptions: [
        BasicAuthOption,
        CertAuthOption,
        AADAuthOption,
        APIKeyAuthOption,
        ImplementMyselfOption,
      ],
      title: getLocalizedString("plugins.apiConnector.whichAuthType.title"),
      placeholder: getLocalizedString("plugins.apiConnector.whichAuthType.placeholder"), // Use the placeholder to display some description
    });
    whichAuthType.addChild(this.buildAADAuthQuestion(ctx, inputs));
    whichAuthType.addChild(this.buildBasicAuthQuestion());
    whichAuthType.addChild(this.buildAPIKeyAuthQuestion());
    return whichAuthType;
  }

  public buildBasicAuthQuestion(): QTreeNode {
    const node = new QTreeNode(basicAuthUsernameQuestion);
    node.condition = { equals: BasicAuthOption.id };
    return node;
  }

  public buildAADAuthQuestion(ctx: Context, inputs: Inputs): QTreeNode {
    let aad;
    if (isV3()) {
      aad = hasAAD(ctx.projectSetting as ProjectSettingsV3);
    } else {
      const solutionSettings = getAzureSolutionSettings(ctx)!;
      aad = isAADEnabled(solutionSettings);
    }
    let node: QTreeNode;
    if (aad || inputs.platform === Platform.CLI_HELP) {
      node = new QTreeNode({
        name: Constants.questionKey.apiAppType,
        type: "singleSelect",
        staticOptions: [reuseAppOption, anotherAppOption],
        title: getLocalizedString("plugins.apiConnector.getQuestion.appType.title"),
      });
      node.condition = { equals: AADAuthOption.id };
      const tenentQuestionNode = new QTreeNode(appTenantIdQuestion);
      tenentQuestionNode.condition = { equals: anotherAppOption.id };
      tenentQuestionNode.addChild(new QTreeNode(appIdQuestion));
      node.addChild(tenentQuestionNode);
    } else {
      node = new QTreeNode(appTenantIdQuestion);
      node.condition = { equals: AADAuthOption.id };
      node.addChild(new QTreeNode(appIdQuestion));
    }
    return node;
  }

  public buildAPIKeyAuthQuestion(): QTreeNode {
    const node = new QTreeNode({
      name: Constants.questionKey.apiAPIKeyLocation,
      type: "singleSelect",
      staticOptions: [requestHeaderOption, queryParamsOption],
      title: getLocalizedString("plugins.apiConnector.getQuestion.apiKeyLocation.title"),
    });
    node.condition = { equals: APIKeyAuthOption.id };

    const keyNameQuestionNode = new QTreeNode(buildAPIKeyNameQuestion());

    node.addChild(keyNameQuestionNode);
    return node;
  }

  public getTelemetryProperties(config: ApiConnectorConfiguration): { [key: string]: string } {
    const properties = {
      [Telemetry.properties.authType]: config.AuthConfig.AuthType.toString(),
      [Telemetry.properties.componentType]: config.ComponentType.join(","),
    };

    switch (config.AuthConfig.AuthType) {
      case AuthType.AAD:
        const aadAuthConfig = config.AuthConfig as AADAuthConfig;
        properties[Telemetry.properties.reuseTeamsApp] = aadAuthConfig.ReuseTeamsApp
          ? Telemetry.valueYes
          : Telemetry.valueNo;
        break;
      case AuthType.APIKEY:
        const authConfig = config.AuthConfig as APIKeyAuthConfig;
        properties[Telemetry.properties.keyLocation] = authConfig.Location;
        break;
      default:
        break;
    }
    return properties;
  }
}
