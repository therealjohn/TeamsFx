// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";

import * as constants from "./constants";
import * as commonUtils from "./commonUtils";
import { Json, ProductName, ProjectSettings, v2, VsCodeEnv } from "@microsoft/teamsfx-api";
import {
  FolderName,
  isConfigUnifyEnabled,
  isMicrosoftTunnelingEnabled,
  LocalEnvManager,
} from "@microsoft/teamsfx-core";
import { VSCodeDepsChecker } from "./depsChecker/vscodeChecker";
import { vscodeLogger } from "./depsChecker/vscodeLogger";
import { vscodeTelemetry } from "./depsChecker/vscodeTelemetry";
import VsCodeLogInstance from "../commonlib/log";
import { detectVsCodeEnv, showError } from "../handlers";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  DepsType,
  ITaskDefinition,
  ProgrammingLanguage,
  TaskDefinition,
} from "@microsoft/teamsfx-core";
import { vscodeHelper } from "./depsChecker/vscodeHelper";
import { createMicrosoftTunnelingTask } from "./microsoftTunnelingTask";

export class TeamsfxTaskProvider implements vscode.TaskProvider {
  public static readonly type: string = ProductName;

  public async provideTasks(token?: vscode.CancellationToken | undefined): Promise<vscode.Task[]> {
    const tasks: vscode.Task[] = [];
    if (vscode.workspace.workspaceFolders) {
      const workspaceFolder: vscode.WorkspaceFolder = vscode.workspace.workspaceFolders[0];
      const workspacePath: string = workspaceFolder.uri.fsPath;
      if (!(await commonUtils.isFxProject(workspacePath))) {
        return tasks;
      }

      const localEnvManager = new LocalEnvManager(VsCodeLogInstance, ExtTelemetry.reporter);
      let projectSettings: ProjectSettings;
      let localSettings: Json | undefined;
      let localEnvInfo: v2.EnvInfoV2 | undefined;
      let localEnv: { [key: string]: string } | undefined;

      try {
        projectSettings = await localEnvManager.getProjectSettings(workspacePath);
        localSettings = await localEnvManager.getLocalSettings(workspacePath, {
          projectId: projectSettings.projectId,
        });
        if (isConfigUnifyEnabled()) {
          localEnvInfo = await localEnvManager.getLocalEnvInfo(workspacePath, {
            projectId: projectSettings.projectId,
          });
        }
        localEnv = await localEnvManager.getLocalDebugEnvs(
          workspacePath,
          projectSettings,
          localSettings,
          localEnvInfo
        );
      } catch (error: any) {
        showError(error);
        return tasks;
      }

      const programmingLanguage = projectSettings?.programmingLanguage;

      // Always provide the following tasks no matter whether it is defined in tasks.json
      const frontendRoot = await commonUtils.getProjectRoot(workspacePath, FolderName.Frontend);
      if (frontendRoot) {
        tasks.push(await this.createFrontendStartTask(workspaceFolder, localEnv));
      }

      const backendRoot = await commonUtils.getProjectRoot(workspacePath, FolderName.Function);
      if (backendRoot) {
        tasks.push(
          await this.createBackendStartTask(workspaceFolder, programmingLanguage, localEnv)
        );
        if (programmingLanguage === ProgrammingLanguage.typescript) {
          tasks.push(await this.createBackendWatchTask(workspaceFolder));
        }
      }

      const authRoot = commonUtils.getAuthServicePath(localEnv);
      if (authRoot) {
        tasks.push(await this.createAuthStartTask(workspaceFolder, authRoot, localEnv));
      }

      const botRoot = await commonUtils.getProjectRoot(workspacePath, FolderName.Bot);
      if (botRoot) {
        const skipNgrok = !vscodeHelper.isNgrokCheckerEnabled();
        const tunnelingTasks = await this.createTunnelingStartTasks(
          projectSettings,
          workspaceFolder,
          botRoot,
          skipNgrok
        );
        for (const task of tunnelingTasks) {
          tasks.push(task);
        }
        const silent: boolean = frontendRoot !== undefined;
        tasks.push(
          await this.createBotStartTask(workspaceFolder, programmingLanguage, localEnv, silent)
        );
      }

      const vscodeEnv = detectVsCodeEnv();
      const isCodeSpaceEnv =
        vscodeEnv === VsCodeEnv.codespaceBrowser || vscodeEnv === VsCodeEnv.codespaceVsCode;
      if (isCodeSpaceEnv) {
        const localTeamsAppId = localSettings?.teamsApp?.teamsAppId as string;
        const debugConfig = { appId: localTeamsAppId };
        tasks.push(await this.createOpenTeamsWebClientTask(workspaceFolder, debugConfig));
      }
    }
    return tasks;
  }

  public async resolveTask(
    task: vscode.Task,
    token?: vscode.CancellationToken | undefined
  ): Promise<vscode.Task | undefined> {
    // Return undefined since all tasks are provided and fully resolved
    return undefined;
  }

  private async createFrontendStartTask(
    workspaceFolder: vscode.WorkspaceFolder,
    localEnv: { [key: string]: string } | undefined,
    definition?: vscode.TaskDefinition,
    problemMatchers?: string | string[]
  ): Promise<vscode.Task> {
    return createTask(
      TaskDefinition.frontendStart(workspaceFolder.uri.fsPath),
      workspaceFolder,
      commonUtils.getFrontendLocalEnv(localEnv),
      definition,
      problemMatchers || constants.frontendProblemMatcher
    );
  }

  private async createBackendStartTask(
    workspaceFolder: vscode.WorkspaceFolder,
    programmingLanguage: string | undefined,
    localEnv: { [key: string]: string } | undefined,
    definition?: vscode.TaskDefinition,
    problemMatchers?: string | string[]
  ): Promise<vscode.Task> {
    const depsChecker = new VSCodeDepsChecker(vscodeLogger, vscodeTelemetry);
    const funcCoreTools = await depsChecker.getDepsStatus(DepsType.FuncCoreTools);

    return createTask(
      TaskDefinition.backendStart(
        workspaceFolder.uri.fsPath,
        programmingLanguage,
        funcCoreTools.command,
        true
      ),
      workspaceFolder,
      commonUtils.getBackendLocalEnv(localEnv),
      definition,
      problemMatchers || constants.backendProblemMatcher,
      true
    );
  }

  private async createAuthStartTask(
    workspaceFolder: vscode.WorkspaceFolder,
    authRoot: string,
    localEnv: { [key: string]: string } | undefined,
    definition?: vscode.TaskDefinition
  ): Promise<vscode.Task> {
    const depsChecker = new VSCodeDepsChecker(vscodeLogger, vscodeTelemetry);
    const dotnet = await depsChecker.getDepsStatus(DepsType.Dotnet);
    return createTask(
      TaskDefinition.authStart(dotnet.command, authRoot),
      workspaceFolder,
      commonUtils.getAuthLocalEnv(localEnv),
      definition,
      constants.authProblemMatcher,
      true
    );
  }

  private async createTunnelingStartTasks(
    projectSettings: ProjectSettings,
    workspaceFolder: vscode.WorkspaceFolder,
    projectRoot: string,
    isSkipped: boolean,
    definition?: vscode.TaskDefinition
  ): Promise<vscode.Task[]> {
    // prepare PATH to execute `ngrok`
    const depsChecker = new VSCodeDepsChecker(vscodeLogger, vscodeTelemetry);
    const ngrok = await depsChecker.getDepsStatus(DepsType.Ngrok);

    const _createTunnelingSkippedTask = (taskName: string) => {
      return createTask(
        TaskDefinition.tunnelingSkippedStart(taskName, workspaceFolder.uri.fsPath),
        workspaceFolder,
        undefined,
        definition,
        constants.ngrokProblemMatcher
      );
    };

    const _createMicrosoftTunnelingTask = (taskName: string) => {
      return createMicrosoftTunnelingTask(taskName, projectSettings, workspaceFolder);
    };

    const _createNgrokTunnelingTask = (taskName: string) => {
      return createTask(
        TaskDefinition.tunnelingStart(
          taskName,
          workspaceFolder.uri.fsPath,
          ngrok.details.binFolders
        ),
        workspaceFolder,
        undefined,
        definition,
        constants.ngrokProblemMatcher
      );
    };

    const _createTask = isSkipped
      ? _createTunnelingSkippedTask
      : isMicrosoftTunnelingEnabled()
      ? _createMicrosoftTunnelingTask
      : _createNgrokTunnelingTask;

    // Also provide ngrok task for backward compatibility.
    // For existing projects, even though it is called `ngrok start`, it may use Microsoft tunneling according to settings.
    const taskNames = [TaskDefinition.TunnelingStartTaskName, TaskDefinition.NgrokStartTaskName];

    return Promise.all(taskNames.map((taskName) => _createTask(taskName)));
  }

  private async createBotStartTask(
    workspaceFolder: vscode.WorkspaceFolder,
    programmingLanguage: string | undefined,
    localEnv: { [key: string]: string } | undefined,
    silent: boolean,
    definition?: vscode.TaskDefinition
  ): Promise<vscode.Task> {
    return createTask(
      TaskDefinition.botStart(workspaceFolder.uri.fsPath, programmingLanguage, true),
      workspaceFolder,
      commonUtils.getBotLocalEnv(localEnv),
      definition,
      constants.botProblemMatcher,
      silent
    );
  }

  private async createOpenTeamsWebClientTask(
    workspaceFolder: vscode.WorkspaceFolder,
    debugConfig: any,
    definition?: vscode.TaskDefinition
  ): Promise<vscode.Task> {
    const command: string = constants.openWenClientCommand;
    definition = definition || { type: TeamsfxTaskProvider.type, command };

    const localTeamsAppId: string | undefined = debugConfig?.appId;
    const commandLine = `npx open-cli https://teams.microsoft.com/_#/l/app/${localTeamsAppId}?installAppPackage=true`;

    const task = new vscode.Task(
      definition,
      workspaceFolder,
      command,
      TeamsfxTaskProvider.type,
      new vscode.ShellExecution(commandLine)
    );

    return task;
  }

  private async createBackendWatchTask(
    workspaceFolder: vscode.WorkspaceFolder,
    definition?: vscode.TaskDefinition
  ): Promise<vscode.Task> {
    return createTask(
      TaskDefinition.backendWatch(workspaceFolder.uri.fsPath),
      workspaceFolder,
      undefined,
      definition,
      constants.tscWatchProblemMatcher,
      true
    );
  }
}

export async function createTask(
  taskDefinition: ITaskDefinition,
  workspaceFolder: vscode.WorkspaceFolder,
  env?: { [key: string]: string } | undefined,
  definition?: vscode.TaskDefinition,
  problemMatchers?: string | string[],
  isSilent?: boolean
): Promise<vscode.Task> {
  definition = definition || {
    type: TeamsfxTaskProvider.type,
    command: taskDefinition.name,
  };

  const options: vscode.ShellExecutionOptions = {
    cwd: taskDefinition.cwd,
    env: env ?? taskDefinition.env,
    // avoid powershell execution policy issue
    executable: taskDefinition.execOptions.needCmd ? "cmd.exe" : undefined,
    shellArgs: taskDefinition.execOptions.needCmd ? ["/c"] : undefined,
  };

  const execution = taskDefinition.execOptions.needShell
    ? new vscode.ShellExecution(taskDefinition.command, options)
    : new vscode.ProcessExecution(taskDefinition.command, taskDefinition.args ?? [], options);

  const task = new vscode.Task(
    definition,
    workspaceFolder,
    taskDefinition.name,
    TeamsfxTaskProvider.type,
    execution,
    problemMatchers
  );
  task.isBackground = taskDefinition.isBackground;
  if (isSilent) {
    task.presentationOptions.reveal = vscode.TaskRevealKind.Silent;
  }
  return task;
}
