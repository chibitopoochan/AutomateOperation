/*
 * Tool name "operation automate tools"
 * How to Use:
 * 1. Install Node.js
 * 2. Execute node automation_tools.js
 */

// load modules
const xlsx = require('xlsx');
const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs').promises;
require('dotenv').config();

// get .env informations
const MY_ID = process.env.SFDC_ID;
const MY_PW = process.env.SFDC_PW;
const STEP_SHEET = process.env.SHEET_NAME;

let headlessMode;

class Step {
  constructor(action, target, value, loop) {
    this.action = action;
    this.target = target;
    this.value = value;
    this.loop = loop;
  }
}

async function loadExcelFile() {
  const steps = [];

  // search excel files
  const filePath = path.join('.', 'input');
  const files = await fs.readdir(filePath);
  const fileName = files.pop();
  console.log(`loading... ${filePath}\\${fileName}`);

  // open excel file
  const book = xlsx.readFile(path.join(filePath, fileName));
  console.log('Excel opened');

  // open sheet
  const sheet = book.Sheets[STEP_SHEET];
  console.log(`select ${STEP_SHEET} sheet`);

  // get headless mode
  headlessMode = sheet.H2.v;
  console.log(`headless mode : ${headlessMode}`);

  // define get cell function
  const getValue = cell => ((typeof cell !== 'undefined' && cell.v !== 'undefined') ? cell.v : '');
  const getActionCell = row => getValue(sheet[xlsx.utils.encode_cell({ c: 1, r: row })]);
  const getTargetCell = row => getValue(sheet[xlsx.utils.encode_cell({ c: 2, r: row })]);
  const getValueCell = row => getValue(sheet[xlsx.utils.encode_cell({ c: 3, r: row })]);
  const getLoopCell = row => getValue(sheet[xlsx.utils.encode_cell({ c: 4, r: row })]);

  // collect to step objects
  let row = 3;
  while (getActionCell(row) !== '') {
    const step = new Step(
      getActionCell(row),
      getTargetCell(row),
      getValueCell(row),
      getLoopCell(row),
    );
    steps.push(step);
    console.log(`step ${row} ${JSON.stringify(step)}`);

    row += 1;
  }
  return steps;
}

/*
 * Screenshot
 */
async function screenshot(page, target) {
  console.log(`Screenshot : save to ${target}`);
  await page.screenshot({ path: path.join('output', target), fullPage: true });
}

/*
 * Go to Action
 */
async function gotoAction(no, page, target) {
  console.log(`GoTo : ${target}`);
  await screenshot(page, `screenshot-${no}.png`);
  await page.goto(target, { waitUntil: 'networkidle2' });
}

/*
 * Type Action
 */
async function typeAction(no, page, target, value) {
  const transID = param => (param === '$ID' ? MY_ID : param);
  const transPW = param => (param === '$PW' ? MY_PW : param);
  const typeValue = transID(transPW(value));
  console.log(`Type : ${typeValue} to ${target} `);
  await page.waitForSelector(target);
  await screenshot(page, `screenshot-${no}.png`);
  await page.type(target, typeValue);
}

/*
 * Click Action
 */
async function clickAction(no, page, target) {
  console.log(`Click : ${target}`);
  await page.waitForSelector(target);
  await screenshot(page, `screenshot-${no}.png`);
  await page.click(target);
}

/*
 * Select Action
 */
async function selectAction(no, page, target, value) {
  console.log(`Select : ${value} on ${target}`);
  await page.waitForSelector(`select[name="${target}"]`);
  await screenshot(page, `screenshot-${no}.png`);
  await page.select(`select[name="${target}"]`, value);
}

/*
 * Download Action
 */
async function downloadAction(no, page, target) {
  console.log(`Download Setting : ${target}`);
  const client = await page.target().createCDPSession();
  await client.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: 'output' });
  await clickAction(no, page, target);
}

/*
 * Tool boot
 */
(async () => {
  // Load excel file
  const steps = [];
  await loadExcelFile().then(array => steps.push(...array));
  console.log(steps);

  const browser = await puppeteer.launch({
    headless: headlessMode,
    slowMo: 50,
  });

  const page = await browser.newPage();
  if (headlessMode) {
    await page.setViewport({ width: 1200, height: 800 });
  }

  console.log('Execute steps');
  // eslint-disable-next-line no-restricted-syntax
  for (let i = 0; i < steps.length; i += 1) {
    const step = steps[i];
    console.log(`Step : ${step.action}`);
    switch (step.action) {
      case 'GoTo':
        // eslint-disable-next-line no-await-in-loop
        await gotoAction(i, page, step.target);
        break;
      case 'Type':
        // eslint-disable-next-line no-await-in-loop
        await typeAction(i, page, step.target, step.value);
        break;
      case 'Click':
        // eslint-disable-next-line no-await-in-loop
        await clickAction(i, page, step.target);
        break;
      case 'Select':
        // eslint-disable-next-line no-await-in-loop
        await selectAction(i, page, step.target, step.value);
        break;
      case 'PrtScn':
        // eslint-disable-next-line no-await-in-loop
        await screenshot(i, page, step.target);
        break;
      case 'DL':
        // eslint-disable-next-line no-await-in-loop
        await downloadAction(i, page, step.target);
        break;
      case 'UL':
      default:
        break;
    }
  }

  await browser.close();
})();
