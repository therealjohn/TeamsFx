// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ProductName, UserError } from "@microsoft/teamsfx-api";
import * as uuid from "uuid";
import * as vscode from "vscode";
import {
  DebugNoSessionId,
  endLocalDebugSession,
  getLocalDebugSession,
  getLocalDebugSessionId,
  getNpmInstallLogInfo,
} from "./commonUtils";
import * as globalVariables from "../globalVariables";
import { TelemetryEvent, TelemetryProperty } from "../telemetry/extTelemetryEvents";
import { Correlator, getHashedEnv, isValidProject } from "@microsoft/teamsfx-core";
import * as path from "path";
import {
  errorDetail,
  Hub,
  issueChooseLink,
  issueLink,
  issueTemplate,
  m365AppsPrerequisitesHelpLink,
} from "./constants";
import * as util from "util";
import VsCodeLogInstance from "../commonlib/log";
import { ExtensionSurvey } from "../utils/survey";
import { TreatmentVariableValue } from "../exp/treatmentVariables";
import { TeamsfxDebugConfiguration } from "./teamsfxDebugProvider";
import { localize } from "../utils/localizeUtils";
import { VS_CODE_UI } from "../extension";
import { localTelemetryReporter, sendDebugAllEvent } from "./localTelemetryReporter";
import { ExtensionErrors, ExtensionSource } from "../error";

export const allRunningTeamsfxTasks: Map<string, number> = new Map<string, number>();
export const allRunningDebugSessions: Set<string> = new Set<string>();
const activeNpmInstallTasks = new Set<string>();

/**
 * This EventEmitter is used to track all running tasks called by `runTask`.
 * Each task executed by `runTask` will have an internal task id.
 * Event emitters use this id to identify each tracked task, and `runTask` matches this id
 * to determine whether a task is terminated or not.
 */
export let taskEndEventEmitter: vscode.EventEmitter<{
  id: string;
  name: string;
  exitCode?: number;
}>;
let taskStartEventEmitter: vscode.EventEmitter<string>;
export const trackedTasks = new Map<string, string>();

function getTaskKey(task: vscode.Task): string {
  if (task === undefined) {
    return "";
  }

  // "source|name|scope"
  const scope = (task.scope as vscode.WorkspaceFolder)?.uri?.toString() || task.scope?.toString();
  return `${task.source}|${task.name}|${scope}`;
}

function isNpmInstallTask(task: vscode.Task): boolean {
  if (task) {
    return task.name.trim().toLocaleLowerCase().endsWith("npm install");
  }

  return false;
}

function isTeamsfxTask(task: vscode.Task): boolean {
  // teamsfx: xxx start / xxx watch
  if (task) {
    if (
      task.source === ProductName &&
      (task.name.trim().toLocaleLowerCase().endsWith("start") ||
        task.name.trim().toLocaleLowerCase().endsWith("watch"))
    ) {
      // provided by toolkit
      return true;
    }

    if (task.definition && task.definition.type === ProductName) {
      // defined by launch.json
      const command = task.definition.command as string;
      return (
        command !== undefined &&
        (command.trim().toLocaleLowerCase().endsWith("start") ||
          command.trim().toLocaleLowerCase().endsWith("watch"))
      );
    }

    // dev:teamsfx and watch:teamsfx
    let commandLine: string | undefined;
    if (task.execution && <vscode.ShellExecution>task.execution) {
      const execution = <vscode.ShellExecution>task.execution;
      commandLine =
        execution.commandLine || `${execution.command} ${(execution.args || []).join(" ")}`;
    }
    if (commandLine !== undefined) {
      return /(npm|yarn)[\s]+(run )?[\s]*[^:\s]+:teamsfx/i.test(commandLine);
    }
  }

  return false;
}

function displayTerminal(taskName: string): boolean {
  const terminal = vscode.window.terminals.find((t) => t.name === taskName);
  if (terminal !== undefined && terminal !== vscode.window.activeTerminal) {
    terminal.show(true);
    return true;
  }

  return false;
}

export async function runTask(task: vscode.Task): Promise<number | undefined> {
  if (task.definition.teamsfxTaskId === undefined) {
    task.definition.teamsfxTaskId = uuid.v4();
  }

  const taskId = task.definition.teamsfxTaskId;
  let started = false;

  return new Promise<number | undefined>((resolve, reject) => {
    // corner case but need to handle - somehow the task does not start
    const startTimer = setTimeout(() => {
      if (!started) {
        reject(new Error("Task start timeout"));
      }
    }, 30000);

    const startListener = taskStartEventEmitter.event((result) => {
      if (taskId === result) {
        clearTimeout(startTimer);
        started = true;
        startListener.dispose();
      }
    });

    vscode.tasks.executeTask(task);

    const endListener = taskEndEventEmitter.event((result) => {
      if (taskId === result.id) {
        endListener.dispose();
        resolve(result.exitCode);
      }
    });
  });
}

// TODO: move to local debug prerequisites checker
async function checkCustomizedPort(component: string, componentRoot: string, checklist: RegExp[]) {
  /*
  const devScript = await loadTeamsFxDevScript(componentRoot);
  if (devScript) {
    let showWarning = false;
    for (const check of checklist) {
      if (!check.test(devScript)) {
        showWarning = true;
        break;
      }
    }

    if (showWarning) {
      VsCodeLogInstance.info(`Customized port detected in ${component}.`);
      if (!globalStateGet(constants.PortWarningStateKeys.DoNotShowAgain, false)) {
        const doNotShowAgain = "Don't Show Again";
        const editPackageJson = "Edit package.json";
        const learnMore = "Learn More";
        vscode.window
          .showWarningMessage(
            util.format(
              localize("teamstoolkit.localDebug.portWarning"),
              component,
              path.join(componentRoot, "package.json")
            ),
            doNotShowAgain,
            editPackageJson,
            learnMore
          )
          .then(async (selected) => {
            if (selected === doNotShowAgain) {
              await globalStateUpdate(constants.PortWarningStateKeys.DoNotShowAgain, true);
            } else if (selected === editPackageJson) {
              vscode.commands.executeCommand(
                "vscode.open",
                vscode.Uri.file(path.join(componentRoot, "package.json"))
              );
            } else if (selected === learnMore) {
              vscode.commands.executeCommand(
                "vscode.open",
                vscode.Uri.parse(constants.localDebugHelpDoc)
              );
            }
          });
      }
    }
  }
  */
}

function onDidStartTaskHandler(event: vscode.TaskStartEvent): void {
  const taskId = event.execution.task?.definition?.teamsfxTaskId;
  if (taskId !== undefined) {
    trackedTasks.set(taskId, event.execution.task.name);
    taskStartEventEmitter.fire(taskId as string);
  }
}

function onDidEndTaskHandler(event: vscode.TaskEndEvent): void {
  const taskId = event.execution.task?.definition?.teamsfxTaskId;
  if (taskId !== undefined && trackedTasks.has(taskId as string)) {
    trackedTasks.delete(taskId as string);
    taskEndEventEmitter.fire({
      id: taskId as string,
      name: event.execution.task.name,
      exitCode: undefined,
    });
  }
}

async function onDidStartTaskProcessHandler(event: vscode.TaskProcessStartEvent): Promise<void> {
  if (globalVariables.workspaceUri && isValidProject(globalVariables.workspaceUri.fsPath)) {
    const task = event.execution.task;
    if (task.scope !== undefined && isTeamsfxTask(task)) {
      allRunningTeamsfxTasks.set(getTaskKey(task), event.processId);
      localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugServiceStart, {
        [TelemetryProperty.DebugServiceName]: task.name,
      });
    } else if (isNpmInstallTask(task)) {
      localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugNpmInstallStart, {
        [TelemetryProperty.DebugNpmInstallName]: task.name,
      });

      if (TreatmentVariableValue.isEmbeddedSurvey) {
        // Survey triggering point
        const survey = ExtensionSurvey.getInstance();
        survey.activate();
      }

      activeNpmInstallTasks.add(task.name);
    }
  }
}

async function onDidEndTaskProcessHandler(event: vscode.TaskProcessEndEvent): Promise<void> {
  const timestamp = new Date();
  const task = event.execution.task;
  const activeTerminal = vscode.window.activeTerminal;

  const taskId = task?.definition?.teamsfxTaskId;
  if (taskId !== undefined) {
    trackedTasks.delete(taskId as string);
    taskEndEventEmitter.fire({
      id: taskId as string,
      name: event.execution.task.name,
      exitCode: event.exitCode,
    });
  }

  if (task.scope !== undefined && isTeamsfxTask(task)) {
    if (event.exitCode !== 0) {
      const currentSession = getLocalDebugSession();
      currentSession.failedServices.push({ name: task.name, exitCode: event.exitCode });
    }
    allRunningTeamsfxTasks.delete(getTaskKey(task));
    localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugService, {
      [TelemetryProperty.DebugServiceName]: task.name,
      [TelemetryProperty.DebugServiceExitCode]: event.exitCode + "",
    });
  } else if (isNpmInstallTask(task)) {
    try {
      activeNpmInstallTasks.delete(task.name);
      if (activeTerminal?.name === task.name && event.exitCode === 0) {
        // when the task in active terminal is ended successfully.
        for (const hiddenTaskName of activeNpmInstallTasks) {
          // display the first hidden terminal.
          if (displayTerminal(hiddenTaskName)) {
            return;
          }
        }
      } else if (activeTerminal?.name !== task.name && event.exitCode !== 0) {
        // when the task in hidden terminal failed to execute.
        displayTerminal(task.name);
      }

      const cwdOption = (task.execution as vscode.ShellExecution).options?.cwd;
      let cwd: string | undefined;
      if (cwdOption !== undefined) {
        cwd = cwdOption.replace("${workspaceFolder}", globalVariables.workspaceUri!.fsPath);
      }
      const npmInstallLogInfo = await getNpmInstallLogInfo();
      let validNpmInstallLogInfo = false;
      if (
        cwd !== undefined &&
        npmInstallLogInfo?.cwd !== undefined &&
        path.relative(npmInstallLogInfo.cwd, cwd).length === 0 &&
        event.exitCode !== undefined &&
        npmInstallLogInfo.exitCode === event.exitCode
      ) {
        const timeDiff = timestamp.getTime() - npmInstallLogInfo.timestamp.getTime();
        if (timeDiff >= 0 && timeDiff <= 20000) {
          validNpmInstallLogInfo = true;
        }
      }
      const properties: { [key: string]: string } = {
        [TelemetryProperty.DebugNpmInstallName]: task.name,
        [TelemetryProperty.DebugNpmInstallExitCode]: event.exitCode + "", // "undefined" or number value
      };
      if (validNpmInstallLogInfo) {
        properties[TelemetryProperty.DebugNpmInstallNodeVersion] =
          npmInstallLogInfo?.nodeVersion + ""; // "undefined" or string value
        properties[TelemetryProperty.DebugNpmInstallNpmVersion] =
          npmInstallLogInfo?.npmVersion + ""; // "undefined" or string value
        properties[TelemetryProperty.DebugNpmInstallErrorMessage] =
          npmInstallLogInfo?.errorMessage?.join("\n") + ""; // "undefined" or string value
      }
      if (event.exitCode !== 0 || properties[TelemetryProperty.DebugNpmInstallErrorMessage]) {
        localTelemetryReporter.sendTelemetryErrorEvent(
          TelemetryEvent.DebugNpmInstall,
          new UserError({ name: ExtensionErrors.DebugNpmInstallError, source: ExtensionSource }),
          properties,
          {},
          [TelemetryProperty.DebugNpmInstallErrorMessage]
        );
      } else {
        localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugNpmInstall, properties);
      }

      if (cwd !== undefined && event.exitCode !== undefined && event.exitCode !== 0) {
        // Do not show this hint message for prerequisites check and automatic npm install
        if (taskId === undefined) {
          let url: string;
          if (validNpmInstallLogInfo) {
            url = `${issueLink}title=new+bug+report: Task '${
              task.name
            }' failed&body=${issueTemplate}${errorDetail}${JSON.stringify(
              npmInstallLogInfo,
              undefined,
              4
            )}`;
          } else {
            url = issueChooseLink;
          }
          const issue = {
            title: localize("teamstoolkit.handlers.reportIssue"),
            run: async (): Promise<void> => {
              vscode.commands.executeCommand("vscode.open", vscode.Uri.parse(url));
            },
          };
          vscode.window
            .showErrorMessage(
              util.format(
                localize("teamstoolkit.localDebug.npmInstallFailedHintMessage"),
                task.name,
                task.name
              ),
              issue
            )
            .then(async (button) => {
              await button?.run();
            });
          await VsCodeLogInstance.error(
            util.format(
              localize("teamstoolkit.localDebug.npmInstallFailedHintMessage"),
              task.name,
              task.name
            )
          );
        }
        terminateAllRunningTeamsfxTasks();
      }
    } catch {
      // ignore any error
    }
  }
}

async function onDidStartDebugSessionHandler(event: vscode.DebugSession): Promise<void> {
  if (globalVariables.workspaceUri && isValidProject(globalVariables.workspaceUri.fsPath)) {
    const debugConfig = event.configuration as TeamsfxDebugConfiguration;
    if (
      debugConfig &&
      debugConfig.name &&
      (debugConfig.url || debugConfig.port) && // it's from launch.json
      !debugConfig.postRestartTask
    ) {
      allRunningDebugSessions.add(event.id);

      // show M365 tenant hint message for Outlook and Office
      if (debugConfig.teamsfxHub === Hub.outlook || debugConfig.teamsfxHub === Hub.office) {
        VS_CODE_UI.showMessage(
          "info",
          localize("teamstoolkit.localDebug.m365TenantHintMessage"),
          false,
          localize("teamstoolkit.localDebug.learnMore")
        ).then(async (result) => {
          if (result.isOk() && result.value === localize("teamstoolkit.localDebug.learnMore")) {
            await VS_CODE_UI.openUrl(m365AppsPrerequisitesHelpLink);
          }
        });
      }

      // and not a restart one
      // send f5 event telemetry
      let env = "";
      if (debugConfig.teamsfxEnv) {
        env = getHashedEnv(debugConfig.teamsfxEnv);
      }

      localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugStart, {
        [TelemetryProperty.DebugSessionId]: event.id,
        [TelemetryProperty.DebugType]: debugConfig.type,
        [TelemetryProperty.DebugRequest]: debugConfig.request,
        [TelemetryProperty.DebugPort]: debugConfig.port + "",
        [TelemetryProperty.DebugRemote]: debugConfig.teamsfxIsRemote + "",
        [TelemetryProperty.DebugAppId]: debugConfig.teamsfxAppId + "",
        [TelemetryProperty.Env]: env,
      });
      // This is the launch browser local debug session.
      if (debugConfig.request === "launch" && !debugConfig.teamsfxIsRemote) {
        // Handle cases that some services failed immediately after start.
        const currentSession = getLocalDebugSession();
        if (currentSession.id !== DebugNoSessionId && currentSession.failedServices.length > 0) {
          terminateAllRunningTeamsfxTasks();
          await vscode.debug.stopDebugging();
          sendDebugAllEvent(
            debugConfig.teamsfxIsRemote,
            new UserError({
              source: ExtensionSource,
              name: ExtensionErrors.DebugServiceFailedBeforeStartError,
            }),
            {
              [TelemetryProperty.DebugFailedServices]: JSON.stringify(
                currentSession.failedServices
              ),
            }
          );
          endLocalDebugSession();
          return;
        }

        await sendDebugAllEvent(debugConfig.teamsfxIsRemote);
      }
    }
  }
}

export function terminateAllRunningTeamsfxTasks(): void {
  for (const task of allRunningTeamsfxTasks) {
    try {
      process.kill(task[1], "SIGTERM");
    } catch (e) {
      // ignore and keep killing others
    }
  }
  allRunningTeamsfxTasks.clear();
}

function onDidTerminateDebugSessionHandler(event: vscode.DebugSession): void {
  if (allRunningDebugSessions.has(event.id)) {
    // a valid debug session
    // send stop-debug event telemetry
    localTelemetryReporter.sendTelemetryEvent(TelemetryEvent.DebugStop, {
      [TelemetryProperty.DebugSessionId]: event.id,
    });

    terminateAllRunningTeamsfxTasks();

    allRunningDebugSessions.delete(event.id);
    if (allRunningDebugSessions.size == 0) {
      endLocalDebugSession();
    }
    allRunningTeamsfxTasks.clear();
  }
}

export function registerTeamsfxTaskAndDebugEvents(): void {
  taskEndEventEmitter = new vscode.EventEmitter<{ id: string; name: string; exitCode?: number }>();
  taskStartEventEmitter = new vscode.EventEmitter<string>();
  globalVariables.context.subscriptions.push({
    dispose() {
      taskEndEventEmitter.dispose();
      taskStartEventEmitter.dispose();
      trackedTasks.clear();
    },
  });

  globalVariables.context.subscriptions.push(vscode.tasks.onDidStartTask(onDidStartTaskHandler));
  globalVariables.context.subscriptions.push(vscode.tasks.onDidEndTask(onDidEndTaskHandler));

  globalVariables.context.subscriptions.push(
    vscode.tasks.onDidStartTaskProcess((event: vscode.TaskProcessStartEvent) =>
      Correlator.runWithId(getLocalDebugSessionId(), onDidStartTaskProcessHandler, event)
    )
  );

  globalVariables.context.subscriptions.push(
    vscode.tasks.onDidEndTaskProcess((event: vscode.TaskProcessEndEvent) =>
      Correlator.runWithId(getLocalDebugSessionId(), onDidEndTaskProcessHandler, event)
    )
  );

  // debug session handler use correlation-id from event.configuration.teamsfxCorrelationId
  // to minimize concurrent debug session affecting correlation-id
  globalVariables.context.subscriptions.push(
    vscode.debug.onDidStartDebugSession((event: vscode.DebugSession) =>
      Correlator.runWithId(
        // fallback to retrieve correlation id from the global variable.
        event.configuration.teamsfxCorrelationId || getLocalDebugSessionId(),
        onDidStartDebugSessionHandler,
        event
      )
    )
  );
  globalVariables.context.subscriptions.push(
    vscode.debug.onDidTerminateDebugSession((event: vscode.DebugSession) =>
      Correlator.runWithId(
        event.configuration.teamsfxCorrelationId || getLocalDebugSessionId(),
        onDidTerminateDebugSessionHandler,
        event
      )
    )
  );
}
