require("dotenv").config();
const PUPPETEER = require("puppeteer");
const PATH = require("path");
const XLSX = require("xlsx");
const FS = require("fs");

const { MONEY_FORWARD_ID, MONEY_FORWARD_PW } = process.env;
if (!MONEY_FORWARD_ID || !MONEY_FORWARD_PW)
  throw new Error("ログイン情報がありません");

const windowSet = (page: any, name: string, value: any) =>
  page.evaluateOnNewDocument(`
    Object.defineProperty(window, '${name}', {
      get() {
        return '${value}'
      }
    })
  `);

interface RowProp {
  date: string;
  start?: string;
  end?: string;
}
interface TimeHistoryRowProp {
  date: string;
  day: string;
  start?: string;
  end?: string;
  hasStart: boolean;
  hasEnd: boolean;
}

const init = async () => {
  let browser = await PUPPETEER.launch({
    headless: false,
    slowMo: 50,
    //args: ["--window-size-1920,1000"],
  });

  let page = await browser.newPage();
  return page;
};

const login = async (page: any) => {
  await page.goto(`https://attendance.moneyforward.com/my_page`);

  await page.waitForTimeout(5000);

  let classByAcceptButton =
    ".attendance-button-mfid.attendance-button-link.attendance-button-size-wide";

  await page.click(classByAcceptButton);

  let id = MONEY_FORWARD_ID;
  let pw = MONEY_FORWARD_PW;
  await page.waitForTimeout(1000);

  await page.focus("input.inputItem");
  for await (const char of id.split("")) {
    await page.keyboard.press(char);
  }
  await page.keyboard.press("Enter");

  await page.waitForTimeout(3000);

  await page.focus("input.inputItem");

  for await (const char of pw.split("")) {
    await page.keyboard.press(char);
  }
  await page.keyboard.press("Enter");

  await page.waitForTimeout(3000);
};

const writeReport = async (page: any, month: string, rows: RowProp[]) => {
  let editPage: string = `https://attendance.moneyforward.com/my_page/bulk_attendances/2023-${month}-01/edit`;

  await page.goto(editPage);

  await page.waitForTimeout(3000);

  try {
    await page.click(
      ".attendance-button-primary.attendance-button-size-small.attendance-button-fullwidth"
    );
  } catch (e) {
    try {
      await page.click(
        ".attendance-button-mfid.attendance-button-link.attendance-button-size-wide"
      );

      await page.waitForTimeout(3000);

      await page.click("input[type='submit'].submitBtn");

      await page.waitForTimeout(3000);
    } catch (e) {}
  }

  await page.goto(editPage);

  await page.waitForTimeout(3000);

  let timeHistroyRows: TimeHistoryRowProp[] = rows.map((el: any) => {
    return {
      ...el,
      day: el.date.split("/")[1],
      hasStart: !!el.start,
      hasEnd: !!el.end,
    };
  });

  var i = -1;
  for await (const excelRow of timeHistroyRows) {
    i++;
    await page.focus(
      `.attendance-table-contents tr:nth-child(${
        i + 1
      }) input.attendance-input-field-small`
    );

    if (excelRow.hasStart || excelRow.hasEnd) {
      if (excelRow.hasStart) {
        for await (const char of (excelRow.start as string).split("")) {
          await page.keyboard.press(char);
        }
      }
      await page.waitForTimeout(50);
      await page.keyboard.press("Tab");
      await page.waitForTimeout(50);
      await page.keyboard.press("Tab");
      await page.waitForTimeout(50);
      if (excelRow.hasEnd) {
        for await (const char of (excelRow.end as string).split("")) {
          await page.keyboard.press(char);
        }
      }
      await page.waitForTimeout(50);
      await page.keyboard.press("Tab");
      await page.waitForTimeout(50);
      await page.keyboard.press("Tab");
      await page.waitForTimeout(50);

      let hourByStartStr: any = excelRow.start?.split(":")[0];
      let hourByEndStr: any = excelRow.end?.split(":")[0];

      let hourByStart: number = +hourByStartStr;
      let hourByEnd: number = +hourByEndStr;

      if (hourByStart < 12 && 12 < hourByEnd) {
        for await (const char of "12:00".split("")) {
          await page.keyboard.press(char);
        }
        await page.waitForTimeout(50);
        await page.keyboard.press("Tab");
        await page.waitForTimeout(50);
        await page.keyboard.press("Tab");
        await page.waitForTimeout(50);
        for await (const char of "13:00".split("")) {
          await page.keyboard.press(char);
        }
      }
    }
  }

  page.$eval(
    "input[type='submit'].attendance-button-primary.attendance-button-size-medium",
    (submit: HTMLInputElement) => {
      submit.click();
    }
  );
  await page.waitForTimeout(5000);
};

const readReport = async (fileName: string) => {
  let source = PATH.join(__dirname, `resources/${fileName}`);
  let workbook = await XLSX.readFile(source);
  let sheet = await workbook.Sheets[`sheet1`];
  let rows = await XLSX.utils.sheet_to_json(sheet);
  return Array.from(rows);
};

const main = async () => {
  let source = PATH.join(__dirname, `resources/`);

  let page = await init();
  try {
    await login(page);

    for await (const xlsxFile of FS.readdirSync(source)) {
      /* xlsxFile Ex) 01.xlsx */
      let rows: any = await readReport(xlsxFile);
      //console.log(JSON.stringify(rows, null, 2));
      console.log(`file name : ${xlsxFile}`);
      await writeReport(page, `${xlsxFile.split(".")[0]}`, rows);
    }
  } catch (e: any) {
    console.error(`Error : ${e}`);
  }
  await page.close();
};

main();
