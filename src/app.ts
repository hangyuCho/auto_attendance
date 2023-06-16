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

const writeReport = async (month: string, rows: RowProp[]) => {
  let browser = await PUPPETEER.launch({
    headless: false,
    slowMo: 50,
    args: ["--window-size-1920,1000"],
  });

  let page = await browser.newPage();

  await page.goto(`https://attendance.moneyforward.com/my_page`);

  await page.waitForTimeout(5000);
  await page.click(
    ".attendance-button-mfid.attendance-button-link.attendance-button-size-wide"
  );

  await page.waitForSelector("input.inputItem");

  await page.$eval("input.inputItem", (el: any) => {
    console.log(JSON.stringify(el.window, null, 2));
  });
  let id = MONEY_FORWARD_ID;
  let pw = MONEY_FORWARD_PW;
  await page.waitForTimeout(1000);
  await page.focus("input.inputItem");
  for await (const char of id.split("")) {
    await page.keyboard.press(char);
  }
  await page.keyboard.press("Enter");

  await page.waitForTimeout(5000);

  await page.waitForTimeout(5000);

  await page.waitForTimeout(1000);
  await page.focus("input.inputItem");
  for await (const char of pw.split("")) {
    await page.keyboard.press(char);
  }
  await page.keyboard.press("Enter");

  await page.waitForTimeout(5000);

  await page.goto(
    `https://attendance.moneyforward.com/my_page/bulk_attendances/2023-04-01/edit`
  );
  await page.waitForTimeout(5000);

  await page.close();
};

const readReport = async (fileName: string) => {
  let source = PATH.join(__dirname, `resources/${fileName}`);
  let workbook = XLSX.readFile(source);
  let sheet = workbook.Sheets[`sheet1`];
  let rows = XLSX.utils.sheet_to_json(sheet);

  console.log(JSON.stringify(rows, null, 2));
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
  { date: "4/13", start: "8.52", end: "18:11" },
  { date: "4/14", start: "8:50", end: "18:16" },
  { date: "4/15" },
  { date: "4/16" },
  { date: "4/17" },
  { date: "4/18", start: "8'53", end: "18:26" },
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
  FS.readdirSync(source).forEach((xlsxFile: string) => {
    /* xlsxFile Ex) 01.xlsx */
    // readReport(xlsxFile);
    //writeReport(`04`, rows_04);
  });
  writeReport(`04`, rows_04);
};

main();
