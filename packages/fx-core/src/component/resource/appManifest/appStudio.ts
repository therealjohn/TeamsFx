// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  AppPackageFolderName,
  BuildFolderName,
  err,
  FxError,
  InputsWithProjectPath,
  M365TokenProvider,
  ok,
  Result,
  TeamsAppManifest,
  TokenProvider,
  v2,
  v3,
} from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import * as path from "path";
import { v4 } from "uuid";
import isUUID from "validator/lib/isUUID";
import {
  AppStudioScopes,
  compileHandlebarsTemplateString,
  getAppDirectory,
} from "../../../common/tools";
import { AppStudioClient } from "../../../plugins/resource/appstudio/appStudio";
import { Constants } from "../../../plugins/resource/appstudio/constants";
import { AppStudioError } from "../../../plugins/resource/appstudio/errors";
import { AppDefinition } from "../../../plugins/resource/appstudio/interfaces/appDefinition";
import { AppStudioResultFactory } from "../../../plugins/resource/appstudio/results";
import { convertToAppDefinition } from "../../../plugins/resource/appstudio/utils/utils";
import { ComponentNames } from "../../constants";
import { readAppManifest } from "./utils";

/**
 * not support the scenario: user provide app package
 */
export async function createOrUpdateTeamsApp(
  ctx: v2.Context,
  inputs: InputsWithProjectPath,
  envInfo: v3.EnvInfoV3,
  tokenProvider: TokenProvider
): Promise<Result<string, FxError>> {
  const appStudioTokenRes = await tokenProvider.m365TokenProvider.getAccessToken({
    scopes: AppStudioScopes,
  });
  if (appStudioTokenRes.isErr()) {
    return err(appStudioTokenRes.error);
  }
  const appStudioToken = appStudioTokenRes.value;
  if (inputs.appPackagePath) {
    return await createOrUpdateTeamsAppByZip(inputs.appPackagePath, ctx, tokenProvider);
  }
  const teamsAppId = envInfo.state[ComponentNames.AppManifest]?.teamsAppId;
  let create = true;
  if (teamsAppId) {
    try {
      await AppStudioClient.getApp(teamsAppId, appStudioToken, ctx.logProvider);
      create = false;
    } catch (error) {}
  }
  if (create) {
    // create teams app
    try {
      const buildPackage = await buildTeamsAppPackage(inputs.projectPath, envInfo!, true);
      if (buildPackage.isErr()) {
        return err(buildPackage.error);
      }
      const archivedFile = await fs.readFile(buildPackage.value);
      const appDefinition = await AppStudioClient.createApp(
        archivedFile,
        appStudioTokenRes.value,
        ctx.logProvider
      );
      ctx.logProvider.info(`teams app created: ${appDefinition.teamsAppId!}`);
      return ok(appDefinition.teamsAppId!);
    } catch (e: any) {
      return err(
        AppStudioResultFactory.SystemError(
          AppStudioError.TeamsAppCreateFailedError.name,
          AppStudioError.TeamsAppCreateFailedError.message(e)
        )
      );
    }
  } else {
    //update teams app
    const buildPackage = await buildTeamsAppPackage(inputs.projectPath, envInfo!);
    if (buildPackage.isErr()) {
      return err(buildPackage.error);
    }
    const archivedFile = await fs.readFile(buildPackage.value);
    const zipEntries = new AdmZip(archivedFile).getEntries();
    const manifestFile = zipEntries.find((x) => x.entryName === Constants.MANIFEST_FILE);
    if (!manifestFile) {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.FileNotFoundError.name,
          AppStudioError.FileNotFoundError.message(Constants.MANIFEST_FILE)
        )
      );
    }
    const manifestString = manifestFile.getData().toString();
    const manifest = JSON.parse(manifestString) as TeamsAppManifest;
    const appDefinition = convertToAppDefinition(manifest);

    const colorIconContent = zipEntries
      .find((x) => x.entryName === manifest.icons.color)
      ?.getData()
      .toString("base64");
    const outlineIconContent = zipEntries
      .find((x) => x.entryName === manifest.icons.outline)
      ?.getData()
      .toString("base64");
    try {
      const app = await AppStudioClient.updateApp(
        manifest.id,
        appDefinition,
        appStudioTokenRes.value,
        undefined,
        colorIconContent,
        outlineIconContent
      );
      ctx.logProvider.info(`teams app updated: ${app.teamsAppId!}`);
      return ok(app.teamsAppId!);
    } catch (e: any) {
      return err(
        AppStudioResultFactory.SystemError(
          AppStudioError.TeamsAppUpdateFailedError.name,
          AppStudioError.TeamsAppUpdateFailedError.message(manifest.id)
        )
      );
    }
  }
}

/**
 * not support the scenario: user provide app package
 */
export async function createOrUpdateTeamsAppByZip(
  zipFilePath: string,
  ctx: v2.Context,
  tokenProvider: TokenProvider
): Promise<Result<string, FxError>> {
  if (!(await fs.pathExists(zipFilePath))) {
    return err(
      AppStudioResultFactory.UserError(
        AppStudioError.FileNotFoundError.name,
        AppStudioError.FileNotFoundError.message(zipFilePath)
      )
    );
  }
  const appStudioTokenRes = await tokenProvider.m365TokenProvider.getAccessToken({
    scopes: AppStudioScopes,
  });
  if (appStudioTokenRes.isErr()) {
    return err(appStudioTokenRes.error);
  }
  const appStudioToken = appStudioTokenRes.value;
  const archivedFileBuffer = await fs.readFile(zipFilePath);
  const zipEntries = new AdmZip(archivedFileBuffer).getEntries();
  const manifestFile = zipEntries.find((x) => x.entryName === Constants.MANIFEST_FILE);
  if (!manifestFile) {
    return err(
      AppStudioResultFactory.UserError(
        AppStudioError.FileNotFoundError.name,
        AppStudioError.FileNotFoundError.message(Constants.MANIFEST_FILE)
      )
    );
  }
  const manifestString = manifestFile.getData().toString();
  const manifest = JSON.parse(manifestString) as TeamsAppManifest;
  const teamsAppId = manifest.id;
  let create = true;
  if (teamsAppId) {
    try {
      await AppStudioClient.getApp(teamsAppId, appStudioToken, ctx.logProvider);
      create = false;
    } catch (error) {}
  }
  if (create) {
    // create teams app
    try {
      const appDefinition = await AppStudioClient.createApp(
        archivedFileBuffer,
        appStudioTokenRes.value,
        ctx.logProvider
      );
      return ok(appDefinition.teamsAppId!);
    } catch (e: any) {
      return err(
        AppStudioResultFactory.SystemError(
          AppStudioError.TeamsAppCreateFailedError.name,
          AppStudioError.TeamsAppCreateFailedError.message(e)
        )
      );
    }
  } else {
    //update teams app
    const appDefinition = convertToAppDefinition(manifest);
    const colorIconContent = zipEntries
      .find((x) => x.entryName === manifest.icons.color)
      ?.getData()
      .toString("base64");
    const outlineIconContent = zipEntries
      .find((x) => x.entryName === manifest.icons.outline)
      ?.getData()
      .toString("base64");
    try {
      const app = await AppStudioClient.updateApp(
        manifest.id,
        appDefinition,
        appStudioTokenRes.value,
        undefined,
        colorIconContent,
        outlineIconContent
      );
      return ok(app.teamsAppId!);
    } catch (e: any) {
      return err(
        AppStudioResultFactory.SystemError(
          AppStudioError.TeamsAppUpdateFailedError.name,
          AppStudioError.TeamsAppUpdateFailedError.message(manifest.id)
        )
      );
    }
  }
}

export async function publishTeamsApp(
  ctx: v2.Context,
  inputs: InputsWithProjectPath,
  envInfo: v3.EnvInfoV3,
  tokenProvider: M365TokenProvider
): Promise<Result<{ appName: string; publishedAppId: string; update: boolean }, FxError>> {
  let archivedFile;
  // User provided zip file
  if (inputs.appPackagePath) {
    if (await fs.pathExists(inputs.appPackagePath)) {
      archivedFile = await fs.readFile(inputs.appPackagePath);
    } else {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.FileNotFoundError.name,
          AppStudioError.FileNotFoundError.message(inputs.appPackagePath)
        )
      );
    }
  } else {
    const buildPackage = await buildTeamsAppPackage(inputs.projectPath, envInfo!);
    if (buildPackage.isErr()) {
      return err(buildPackage.error);
    }
    archivedFile = await fs.readFile(buildPackage.value);
  }

  const zipEntries = new AdmZip(archivedFile).getEntries();

  const manifestFile = zipEntries.find((x) => x.entryName === Constants.MANIFEST_FILE);
  if (!manifestFile) {
    return err(
      AppStudioResultFactory.UserError(
        AppStudioError.FileNotFoundError.name,
        AppStudioError.FileNotFoundError.message(Constants.MANIFEST_FILE)
      )
    );
  }
  const manifestString = manifestFile.getData().toString();
  const manifest = JSON.parse(manifestString) as TeamsAppManifest;

  // manifest.id === externalID
  const appStudioTokenRes = await tokenProvider.getAccessToken({ scopes: AppStudioScopes });
  if (appStudioTokenRes.isErr()) {
    return err(appStudioTokenRes.error);
  }
  const existApp = await AppStudioClient.getAppByTeamsAppId(manifest.id, appStudioTokenRes.value);
  if (existApp) {
    let executePublishUpdate = false;
    let description = `The app ${existApp.displayName} has already been submitted to tenant App Catalog.\nStatus: ${existApp.publishingState}\n`;
    if (existApp.lastModifiedDateTime) {
      description =
        description + `Last Modified: ${existApp.lastModifiedDateTime?.toLocaleString()}\n`;
    }
    description = description + "Do you want to submit a new update?";
    const res = await ctx.userInteraction.showMessage("warn", description, true, "Confirm");
    if (res?.isOk() && res.value === "Confirm") executePublishUpdate = true;

    if (executePublishUpdate) {
      const appId = await AppStudioClient.publishTeamsAppUpdate(
        manifest.id,
        archivedFile,
        appStudioTokenRes.value
      );
      return ok({ publishedAppId: appId, appName: manifest.name.short, update: true });
    } else {
      throw AppStudioResultFactory.SystemError(
        AppStudioError.TeamsAppPublishCancelError.name,
        AppStudioError.TeamsAppPublishCancelError.message(manifest.name.short)
      );
    }
  } else {
    const appId = await AppStudioClient.publishTeamsApp(
      manifest.id,
      archivedFile,
      appStudioTokenRes.value
    );
    return ok({ publishedAppId: appId, appName: manifest.name.short, update: false });
  }
}

/**
 * Build appPackage.{envName}.zip
 * @returns Path for built Teams app package
 */
export async function buildTeamsAppPackage(
  projectPath: string,
  envInfo: v3.EnvInfoV3,
  withEmptyCapabilities = false
): Promise<Result<string, FxError>> {
  const buildFolderPath = path.join(projectPath, BuildFolderName, AppPackageFolderName);
  await fs.ensureDir(buildFolderPath);
  const appDefinitionRes = await getAppDefinitionAndManifest(projectPath, envInfo);
  if (appDefinitionRes.isErr()) {
    return err(appDefinitionRes.error);
  }
  const manifest: TeamsAppManifest = appDefinitionRes.value[1];
  if (!isUUID(manifest.id)) {
    manifest.id = v4();
  }
  if (withEmptyCapabilities) {
    manifest.bots = [];
    manifest.composeExtensions = [];
    manifest.configurableTabs = [];
    manifest.staticTabs = [];
    manifest.webApplicationInfo = undefined;
  }
  const appDirectory = await getAppDirectory(projectPath);
  const colorFile = path.join(appDirectory, manifest.icons.color);
  if (!(await fs.pathExists(colorFile))) {
    return err(
      AppStudioResultFactory.UserError(
        AppStudioError.FileNotFoundError.name,
        AppStudioError.FileNotFoundError.message(colorFile)
      )
    );
  }

  const outlineFile = path.join(appDirectory, manifest.icons.outline);
  if (!(await fs.pathExists(outlineFile))) {
    return err(
      AppStudioResultFactory.UserError(
        AppStudioError.FileNotFoundError.name,
        AppStudioError.FileNotFoundError.message(outlineFile)
      )
    );
  }

  const zip = new AdmZip();
  zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest, null, 4)));

  // outline.png & color.png, relative path
  let dir = path.dirname(manifest.icons.color);
  zip.addLocalFile(colorFile, dir === "." ? "" : dir);
  dir = path.dirname(manifest.icons.outline);
  zip.addLocalFile(outlineFile, dir === "." ? "" : dir);

  const zipFileName = path.join(buildFolderPath, `appPackage.${envInfo.envName}.zip`);
  zip.writeZip(zipFileName);

  const manifestFileName = path.join(buildFolderPath, `manifest.${envInfo.envName}.json`);
  if (await fs.pathExists(manifestFileName)) {
    await fs.chmod(manifestFileName, 0o777);
  }
  await fs.writeFile(manifestFileName, JSON.stringify(manifest, null, 4));
  await fs.chmod(manifestFileName, 0o444);

  return ok(zipFileName);
}

/**
 * Validate manifest
 * @returns an array of validation error strings
 */
export async function validateManifest(
  manifest: TeamsAppManifest
): Promise<Result<string[], FxError>> {
  // TODO: import teamsfx-manifest package
  return ok([]);
}

async function getAppDefinitionAndManifest(
  projectPath: string,
  envInfo: v3.EnvInfoV3
): Promise<Result<[AppDefinition, TeamsAppManifest], FxError>> {
  // Read template
  const manifestTemplateRes = await readAppManifest(projectPath);
  if (manifestTemplateRes.isErr()) {
    return err(manifestTemplateRes.error);
  }
  let manifestString = JSON.stringify(manifestTemplateRes.value);

  // Render mustache template with state and config
  const view = {
    config: envInfo.config,
    state: envInfo.state,
  };
  manifestString = compileHandlebarsTemplateString(manifestString, view);

  const manifest: TeamsAppManifest = JSON.parse(manifestString);
  const appDefinition = convertToAppDefinition(manifest);

  return ok([appDefinition, manifest]);
}
