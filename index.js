import { subscribe } from "@parcel/watcher";
import ExcelJS from "exceljs";
import simpleGit from "simple-git";
import path from "path";
import last from "lodash/last.js";
import fs from "fs-extra";
import chalk from "chalk";

const directoryToWatch = "C:\\Projects\\sample-excel-repo\\";

const git = simpleGit(directoryToWatch);

async function getCommitId(params) {
  const commitId = await git.revparse(["HEAD"]);
  return commitId;
}

const EXCEL_TEMP_PATTERNS = [
  /^~\$.+\.xlsx?$/, // temporary files
  /^.+\.tmp$/,
  /^\..+\.xlsx?\.tmp$/,
];

export const isExcelTempFile = (filename) => {
  return EXCEL_TEMP_PATTERNS.some((pattern) => {
    return pattern.test(path.basename(filename));
  });
};

const artifactFileRegex =
  /^[\w-]+_(EngData|DataDictionary)(\d+)?\.(?:xlsx?|xlsm|xlsb|xltx?|xltm|xlam)$/i;

const isArtifactFile = (filename) => {
  if (filename) return artifactFileRegex.test(path.basename(filename));
};

async function watchArtifacts() {
  try {
    console.log(chalk.cyan("watching artifacts directory..."));

    const subscriber = subscribe(
      directoryToWatch,
      async (error, events) => {
        if (error) {
          console.error(`Error: ${error}`);
          return;
        }

        for (const event of events) {
          const { path: filepath, type } = event;
          const etx = path.extname(filepath);

          const fileName = last(filepath.split("\\"));
          console.log(fileName);

          if (
            isArtifactFile(fileName) &&
            type === "update" &&
            !isExcelTempFile(filepath)
          ) {
            console.time("File pre-processing");
            console.log(chalk.yellow(`File changed: ${filepath}`));
            const folderExtension = filepath.split(directoryToWatch)[1];
            console.log(chalk.green(folderExtension));
            console.log(chalk.yellowBright(filepath));
            console.log(chalk.yellow(folderExtension));

            console.timeEnd("File pre-processing");

            console.time("File Copy");

            const copyPath = path.join(
              directoryToWatch,
              "dist",
              folderExtension
            );

            fs.copy(filepath, copyPath, (err) => {
              if (err) {
                console.log(chalk.red(`Error copying file: ${err.message}`));
              }
              console.log(`file copied successfully...`);
            });

            console.timeEnd(`File Copy`);

            console.time(`Fetch Commit Hash`);
            const commitId = await getCommitId();
            console.log(chalk.magenta(commitId.trim()));
            console.timeEnd(`Fetch Commit Hash`);

            console.time(`Load Workbook`);
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(copyPath);
            console.timeEnd(`Load Workbook`);

            console.log(chalk.yellow(`loaded workbook at copypath...`));

            console.time(`Get Version Sheet`);
            const worksheet = workbook.getWorksheet("Version");

            if (!worksheet) {
              console.log("No such worksheet... adding new worksheet");
              return;
            }

            console.log(chalk.magentaBright(worksheet?.name));
            console.timeEnd(`Get Version Sheet`);

            console.time(`Write to Version Sheet`);
            worksheet.getCell("G3").value = `GIT_COMMMIT_ID: ${commitId}`;
            await workbook.xlsx.writeFile(copyPath);
            console.timeEnd(`Write to Version Sheet`);

            console.log(
              chalk.green(
                `Complete copying file from one location to another location`
              )
            );
          }
        }
      },
      {
        ignore: [".git", "dist"],
      }
    );
  } catch (err) {
    console.error(`Error watching artifatcs: ${err}`);
  }
}

watchArtifacts();
