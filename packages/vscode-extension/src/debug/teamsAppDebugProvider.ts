import {
  CancellationToken,
  DebugConfiguration,
  DebugConfigurationProvider,
  WorkspaceFolder,
  TaskProvider,
  ProviderResult,
  Task,
  workspace,
  FileSystemWatcher,
  ShellExecution,
  TaskScope,
} from "vscode";
import { generateAccountHint } from "./teamsfxDebugProvider";
import * as yaml from "js-yaml";
import * as fs from "fs-extra";
import * as path from "path";

export interface TeamsAppDebugConfiguration extends DebugConfiguration {
  fxTemplate: string;
}

type FxDebugStep = {
  name: string;
  template: string;
  parameter: Record<string, unknown>;
};

type FxDebugTemplate = {
  apiVersion: string;
  kind: "debug";
  metadata: {
    name: string;
  };
  spec: {
    steps: FxDebugStep[];
  };
};

async function parseYamlTemplate(path: string): Promise<FxDebugTemplate> {
  return yaml.load(await fs.readFile(path, "utf-8")) as FxDebugTemplate;
}

export class TeamsAppDebugProvider implements DebugConfigurationProvider {
  public async resolveDebugConfiguration?(
    folder: WorkspaceFolder | undefined,
    debugConfiguration: TeamsAppDebugConfiguration,
    token?: CancellationToken
  ): Promise<DebugConfiguration | undefined> {
    if (!debugConfiguration.fxTemplate || !folder) {
      return debugConfiguration;
    }

    const yamlPath = path.join(folder.uri.fsPath, debugConfiguration.fxTemplate);
    const fxTpl = await parseYamlTemplate(yamlPath);
    if (fxTpl.kind !== "debug") {
      return debugConfiguration;
    }
    const teamsAppId = fxTpl.spec.steps[1].parameter.teamsAppId;
    const accountHint = await generateAccountHint(true);

    return {
      name: `${debugConfiguration.name}-${fxTpl.metadata.name}`,
      type: "pwa-msedge",
      request: "launch",
      url: `https://teams.microsoft.com/l/app/${teamsAppId}?installAppPackage=true&webjoin=true&${accountHint}`,
      preLaunchTask: `teamsfx: ${fxTpl.spec.steps[0].name}`,
    };
  }
}

export class TeamsFxDebugTaskProvider implements TaskProvider {
  static TYPE = "teamsfx";
  private taskPromise?: Promise<Task[]>;
  private workspaceRoot: string;
  constructor(workspaceRoot: string) {
    this.workspaceRoot = workspaceRoot;

    const file = path.join(workspaceRoot, "debug.yaml");
    const fileWatcher: FileSystemWatcher = workspace.createFileSystemWatcher(file);
    fileWatcher.onDidChange(() => {
      this.taskPromise = undefined;
    });
    fileWatcher.onDidCreate(() => {
      this.taskPromise = undefined;
    });
    fileWatcher.onDidDelete(() => {
      this.taskPromise = undefined;
    });
  }

  provideTasks(token: CancellationToken): ProviderResult<Task[]> {
    if (!this.taskPromise) {
      this.taskPromise = this.getTasks();
    }
    return this.taskPromise;
  }

  private async getTasks(): Promise<Task[]> {
    const yamlPath = path.join(this.workspaceRoot, "debug.yaml");
    const fxTpl = await parseYamlTemplate(yamlPath);
    const tasks = fxTpl.spec.steps
      .filter((step) => step.template === "script-task")
      .map((step) => {
        const task = new Task(
          { type: TeamsFxDebugTaskProvider.TYPE, task: step.name },
          TaskScope.Workspace,
          step.name,
          TeamsFxDebugTaskProvider.TYPE,
          new ShellExecution(step.parameter["cmd"] as string, {
            cwd: path.join(this.workspaceRoot, step.parameter["folder"] as string),
          }),
          "$teamsfx-task-watch"
        );
        task.isBackground = step.parameter["isBackground"] as boolean;
        return task;
      });
    return tasks;
  }

  resolveTask(task: Task, token: CancellationToken): ProviderResult<Task> {
    return undefined;
  }
}
