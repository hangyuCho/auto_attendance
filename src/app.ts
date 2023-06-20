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

const rows_04: RowProp[] = [
  { date: "4/1" },
  { date: "4/2" },
  { date: "4/3", start: "8:55", end: "18:31" },
  { date: "4/4", start: "8:50", end: "18:01" },
  { date: "4/5", start: "8:56", end: "18:13" },
  { date: "4/6", start: "8:50", end: "18:26" },
  { date: "4/7", start: "8:50", end: "18:38" },
  { date: "4/8" },
  { date: "4/9" },
  { date: "4/10", start: "8:53", end: "18:06" },
  { date: "4/11", start: "8:50", end: "18:53" },
  { date: "4/12", start: "8:50", end: "18:14" },
  { date: "4/13", start: "8:52", end: "18:11" },
  { date: "4/14", start: "8:50", end: "18:16" },
  { date: "4/15" },
  { date: "4/16" },
  { date: "4/17" },
  { date: "4/18", start: "8:53", end: "18:26" },
  { date: "4/19", start: "8:00", end: "18:08" },
  { date: "4/20", start: "8:00", end: "18:18" },
  { date: "4/21", start: "8:51", end: "12:02" },
  { date: "4/22" },
  { date: "4/23" },
  { date: "4/24", start: "8:51", end: "18:37" },
  { date: "4/25", start: "8:53", end: "19:38" },
  { date: "4/26", start: "8:49", end: "18:28" },
  { date: "4/27", start: "8:53", end: "18:35" },
  { date: "4/28", start: "8:50", end: "18:00" },
  { date: "4/29" },
  { date: "4/30" },
];

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
