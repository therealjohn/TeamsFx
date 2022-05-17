// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import os from "os";
import fs from "fs-extra";
import { Argv } from "yargs";
import { FxError, err, ok, Result, UserError, Void } from "@microsoft/teamsfx-api";
import { YargsCommand } from "../yargsCommand";
import { NotValidInputValue } from "../error";
import { cliSource, teamsManifestSchema, teamsManifestVersion } from "../constants";

export default class UpgradeManifest extends YargsCommand {
  public readonly commandHead = "upgrade-manifest";
  public readonly command = this.commandHead;
  public readonly description = "Upgrade Teams manifest to support Outlook and Office apps";

  public builder(yargs: Argv): Argv<any> {
    yargs
      .options("in", {
        description: `The path to the Teams manifest file.`,
        type: "string",
      })
      .options("out", {
        description: `The path to the output manifest file.`,
        type: "string",
      });
    return yargs.version(false);
  }

  public async runCommand(args: { [argName: string]: string }): Promise<Result<null, FxError>> {
    if (!args.in) {
      return err(NotValidInputValue("in", "The path to the Teams manifest file is invalid"));
    }
    if (!args.out) {
      return err(NotValidInputValue("out", "The path to the output manifest file is invalid"));
    }
    const in_ = args.in;
    const out = args.out;

    return (await this.upgradeManifest(in_, out)).map(() => null);
  }

  public async upgradeManifest(
    sourcePath: string,
    outPath: string
  ): Promise<Result<Void, FxError>> {
    try {
      const manifest = await fs.readJSON(sourcePath);
      manifest["$schema"] = teamsManifestSchema;
      manifest["manifestVersion"] = teamsManifestVersion;

      // TODO: migrate Teams App Resource-specific consent
      if (!!manifest?.webApplicationInfo?.applicationPermissions) {
        manifest.webApplicationInfo.applicationPermissions = undefined;
      }
      await fs.writeJSON(outPath, manifest, { spaces: 4, EOL: os.EOL });
      return ok(Void);
    } catch (e: any) {
      return err(
        new UserError({
          error: e,
          source: cliSource,
          name: "UpgradeManifestError",
        })
      );
    }
  }
}
