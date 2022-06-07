import "mocha";
import chai from "chai";
import { convertToValidProjectSettingsAppName } from "../../src/common/utils";

describe("convert to valid AppName in ProjectSetting", () => {
  it("convert app name", () => {
    const appName = "app.123";
    const expected = "app123";
    const projectSettingsName = convertToValidProjectSettingsAppName(appName);

    chai.assert.equal(projectSettingsName, expected);
  });

  it("convert app name", () => {
    const appName = "app.1@@2！3";
    const expected = "app123";
    const projectSettingsName = convertToValidProjectSettingsAppName(appName);

    chai.assert.equal(projectSettingsName, expected);
  });
});
