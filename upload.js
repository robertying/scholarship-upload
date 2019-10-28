const puppeteer = require("puppeteer");
const Xlsx = require("xlsx");
const ProgressBar = require("progress");

const debug = true;

const username = process.argv[2];
const password = process.argv[3];
const filename = process.argv[4];

if (!username || !password) {
  console.error("Username and password are required");
  process.exit(1);
}

if (!filename) {
  console.error("Excel file required");
  process.exit(1);
}

const workbook = Xlsx.readFile(filename);
const firstWorksheet = workbook.Sheets[workbook.SheetNames[0]];

const applications = Xlsx.utils
  .sheet_to_json(firstWorksheet, {
    header: 1
  })
  .filter(i => i.length !== 0);
applications.shift();

const bar = new ProgressBar("Uploading [:bar] :current/:total :etas", {
  total: applications.length,
  width: 50
});

bar.interrupt("Begin to upload...");

const honors = {
  综合优秀奖: 0,
  学业优秀奖: 1,
  社会工作优秀奖: 2,
  体育优秀奖: 4,
  文艺优秀奖: 5,
  社会实践优秀奖: 6,
  科技创新优秀奖: 7,
  无校级荣誉: 8,
  学习进步奖: 9,
  志愿公益优秀奖: 10,
  好读书奖: 11
};

(async () => {
  const browser = await puppeteer.launch({
    headless: !debug,
    slowMo: 500
  });
  const page = await browser.newPage();
  page.setViewport({ width: 1920, height: 1080 });

  await page.goto("https://info.tsinghua.edu.cn/");
  const usernameField = await page.$('input[name="userName"]');
  const passwordField = await page.$('input[name="password"]');
  const loginButton = await page.$('input[type="image"]');
  await usernameField.type(username);
  await passwordField.type(password);
  await loginButton.click();

  await page.waitFor(3000); // adjust this value based on how long it takes to log in
  await page.evaluate(() => {
    document.querySelector(".hot_nav_left").childNodes[7].firstChild.click();
  });

  await page.waitForNavigation();
  await page.goto(
    "http://jxxxfw.cic.tsinghua.edu.cn/roam.do3?type=srch&id=1024"
  );

  await page.waitFor(1000);
  const scholarshipTab = (await page.$$(".fir_li"))[1];
  await scholarshipTab.click();

  await page.waitFor(1000);
  const universityManagedTab = await page.waitForSelector(
    'a[data="/f/bksjzd/jxj/v_bj_jxj_xflrb/xsgzz/beforePageList"]'
  );
  await universityManagedTab.click();

  let i = 0;
  for (const application of applications) {
    try {
      const studentId = application[0].toString();
      const honor = application[3].trim();
      const scholarshipCode = application[5].trim();
      await fillForm(page, studentId, honor, scholarshipCode);
      i++;
      bar.tick();
    } catch {
      bar.terminate();
      console.error("Upload failed at:");
      console.error(application);
    }
  }

  bar.complete();
  await browser.close();
})();

const fillForm = async (page, studentId, honorName, scholarshipCode) => {
  const addButton = await page.waitForSelector("#addbtn");
  await addButton.click();
  const idInput = await page.waitForSelector("#q_xh");
  await idInput.evaluate((e, studentId) => (e.value = studentId), studentId);
  const searchButton = await page.$(".a_btn");
  await searchButton.click();
  await page.waitFor(500);
  const honorSelection = await page.$("#select2-jllxm-container");
  await honorSelection.tap();
  const honorList = await page.$$(".select2-results__option");
  await honorList[honors[honorName]].tap();
  const scholarshipSelection = await page.$("#select2-dm-container");
  await scholarshipSelection.tap();
  const scholarshipList = await page.$$(".select2-results__option");
  const scholarshipListResult = await Promise.all(
    scholarshipList.map(e =>
      e.evaluate(
        (e, scholarshipCode) => e.id.includes(scholarshipCode),
        scholarshipCode
      )
    )
  );
  const scholarship =
    scholarshipList[scholarshipListResult.findIndex(r => r === true)];
  await scholarship.tap();
  const submitButton = await page.$(".btn.btn-sub");
  await submitButton.click();
  await page.waitFor(500);
  const closeButton = (await page.$$(".btn.btn-sub"))[1];
  await closeButton.click();
};
