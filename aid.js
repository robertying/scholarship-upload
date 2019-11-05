const puppeteer = require("puppeteer");
const Xlsx = require("xlsx");
const ProgressBar = require("progress");

const debug = true; // headless mode strangely not working
const delay = 5000; // up to how fast your computer loads pages

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

const bar = new ProgressBar("Uploading [:bar] :current/:total ETA: :etas", {
  total: applications.length,
  width: 50
});

bar.interrupt("Begin to upload...");

const aids = {
  "清华之友——怀庄助学金": [
    { code: "Z2052032", amount: 3200, type: "校管校分" }
  ],
  清华大学生活费助学金: [{ code: "Z2062020", amount: 2000, type: "校管校分" }],
  "清华之友——励志助学金": [
    { code: "Z2072040", amount: 4000, type: "校管校分" }
  ],
  "恒大集团助学基金(一)": [
    { code: "Z2122065", amount: 6500, type: "校管校分" }
  ],
  恒大集团助学基金: [{ code: "Z2132065", amount: 6500, type: "校管校分" }],
  "清华之友——黄俞助学金": [
    { code: "Z2152120", amount: 12000, type: "校管校分" }
  ],
  龙门希望工程助学金: [{ code: "Z2322050", amount: 5000, type: "校管校分" }],
  清华伟新励学金: [{ code: "Z2352040", amount: 4000, type: "校管校分" }],
  "清华之友——赵敏意助学金": [
    { code: "Z2432100", amount: 10000, type: "校管校分" }
  ],
  "清华之友——咏芳助学金": [
    { code: "Z2492040", amount: 4000, type: "校管校分" }
  ],
  "清华大学——昱鸿助学金": [
    { code: "Z2552050", amount: 5000, type: "校管校分" },
    { code: "Z2552100", amount: 10000, type: "校管校分" }
  ],
  "清华之友——张明为助学金": [
    { code: "Z2612050", amount: 5000, type: "校管校分" }
  ],
  "清华之友——一汽丰田助学金": [
    { code: "Z2682050", amount: 5000, type: "校管校分" }
  ],
  "清华校友——河南校友会励学金": [
    { code: "Z4262130", amount: 13000, type: "校管校分" }
  ],
  "清华校友——传信励学基金": [
    { code: "Z4312010", amount: 1000, type: "校管校分" },
    { code: "Z4312020", amount: 2000, type: "校管校分" },
    { code: "Z4312030", amount: 3000, type: "校管校分" }
  ],
  "清华校友——德强励学金": [
    { code: "Z4372050", amount: 5000, type: "校管校分" }
  ],
  "清华校友——孟昭英励学基金": [
    { code: "Z4492010", amount: 1000, type: "校管校分" },
    { code: "Z4492020", amount: 2000, type: "校管校分" },
    { code: "Z4492030", amount: 3000, type: "校管校分" },
    { code: "Z4492060", amount: 6000, type: "校管校分" }
  ],
  "清华校友——常迵励学基金": [
    { code: "Z4502005", amount: 500, type: "校管校分" },
    { code: "Z4502010", amount: 1000, type: "校管校分" },
    { code: "Z4502015", amount: 1500, type: "校管校分" },
    { code: "Z4502020", amount: 2000, type: "校管校分" },
    { code: "Z4502050", amount: 5000, type: "校管校分" }
  ],
  "清华校友——张维国励学基金": [
    { code: "Z4612010", amount: 1000, type: "校管校分" },
    { code: "Z4612060", amount: 6000, type: "校管校分" },
    { code: "Z4612080", amount: 8000, type: "校管校分" },
    { code: "Z4612100", amount: 10000, type: "校管校分" }
  ],
  "清华校友——凌复云·马晓云励学基金": [
    { code: "Z4642060", amount: 6000, type: "校管校分" }
  ],
  "清华之友——华硕励学基金": [
    { code: "Z4922030", amount: 3000, type: "校管校分" },
    { code: "Z4922050", amount: 5000, type: "校管校分" }
  ],
  清华江西校友励学基金: [
    { code: "Z5392030", amount: 3000, type: "校管校分" },
    { code: "Z5392060", amount: 6000, type: "校管校分" }
  ],
  清华校友零零励学基金: [{ code: "Z5412060", amount: 6000, type: "校管校分" }],
  清华校友励学金: [{ code: "Z5562100", amount: 10000, type: "校管校分" }],
  "清华校友——吴道怀史常忻励学基金": [
    { code: "Z5712040", amount: 4000, type: "校管校分" }
  ],
  "清华校友——广州校友会励学金（周进波）": [
    { code: "Z5819060", amount: 6000, type: "校管校分" }
  ],
  "清华校友——代贻榘励学基金": [
    { code: "Z6002060", amount: 6000, type: "校管校分" }
  ],
  "清华校友——山西校友会励学基金": [
    { code: "Z6022060", amount: 6000, type: "校管校分" }
  ],
  "清华校友——李志坚励学基金": [
    { code: "Z6102050", amount: 5000, type: "校管校分" },
    { code: "Z6102100", amount: 10000, type: "校管校分" }
  ],
  珠海市得理慈善基金会清华励学金: [
    { code: "Z6142050", amount: 5000, type: "校管校分" },
    { code: "Z6152100", amount: 10000, type: "校管校分" }
  ],
  清华78届雷四班校友及苏宁电器励学基金: [
    { code: "Z6182050", amount: 5000, type: "校管校分" }
  ],
  "清华校友励学金（任向军）": [
    { code: "Z6242120", amount: 12000, type: "校管校分" }
  ],
  国家助学金: [
    { code: "Z2012020", amount: 2000, type: "校管院分" },
    { code: "Z2012030", amount: 3000, type: "校管院分" },
    { code: "Z2012050", amount: 5000, type: "校管院分" }
  ]
};

const schoolManagedAids = applications.filter(
  i => aids[i[3].trim()][0].type === "校管校分"
);
const departmentManagedAids = applications.filter(
  i => aids[i[3].trim()][0].type === "校管院分"
);

(async () => {
  const browser = await puppeteer.launch({
    headless: !debug
  });
  const page = await browser.newPage();
  await page.setViewport({ width: 1920, height: 1080 });

  bar.interrupt("[progress] logging in...");
  await page.goto("https://info.tsinghua.edu.cn/");
  const usernameField = await page.$('input[name="userName"]');
  const passwordField = await page.$('input[name="password"]');
  const loginButton = await page.$('input[type="image"]');
  await usernameField.type(username);
  await passwordField.type(password);
  await loginButton.click();

  bar.interrupt("[progress] navigating to management site...");
  await page.waitFor(delay);
  await page.evaluate(() => {
    document.querySelector(".hot_nav_left").childNodes[7].firstChild.click();
  });
  await page.waitForNavigation();
  await page.goto(
    "http://jxxxfw.cic.tsinghua.edu.cn/roam.do3?type=srch&id=1024"
  );

  await page.waitFor(1000);
  const aidTab = (await page.$$(".fir_li"))[2];
  await aidTab.click();
  await page.waitFor(1000);

  bar.interrupt("[progress] start to upload 校管校分...");
  const schoolManagedTab = await page.waitForSelector(
    'a[data="/f/bksjzd/zxj/v_bj_zxj_xflrb/xsgzz/beforePageList"]'
  );
  await schoolManagedTab.click();
  await page.waitFor(1000);

  for (const aid of schoolManagedAids) {
    try {
      const studentId = aid[0].toString();
      const code = aid[4].trim();
      await fillForm(page, studentId, code);
      bar.tick();
    } catch (e) {
      console.error("Upload failed at:");
      console.error(aid);
    }
  }

  bar.interrupt("[progress] start to upload 校管院分...");
  let closeButton = (await page.$$(".btn"))[5];
  await closeButton.click();
  const departmentManagedTab = await page.waitForSelector(
    'a[data="/f/bksjzd/zxj/v_bj_zxj_yflrb/xsgzz/beforePageList"]'
  );
  await departmentManagedTab.click();
  await page.waitFor(1000);

  for (const aid of departmentManagedAids) {
    try {
      const studentId = aid[0].toString();
      const code = aid[4].trim();
      await fillForm(page, studentId, code);
      bar.tick();
    } catch (e) {
      console.error("Upload failed at:");
      console.error(aid);
    }
  }

  bar.interrupt("Upload finished!");
  await browser.close();
})();

const fillForm = async (page, studentId, aidCode) => {
  const addButton = await page.waitForSelector("#addbtn");
  await addButton.click();
  const idInput = await page.waitForSelector("#q_xh");
  await idInput.evaluate((e, studentId) => (e.value = studentId), studentId);
  const searchButton = await page.$(".a_btn");
  await searchButton.click();

  await page.waitFor(500);
  const aidSelection = await page.$("#select2-jllxm-container");
  await aidSelection.tap();
  await aidSelection.tap();
  await aidSelection.tap();
  await page.waitFor(".select2-results__option");
  const aidList = await page.$$(".select2-results__option");
  const aidListResult = await Promise.all(
    aidList.map(e =>
      e.evaluate((e, aidCode) => e.id.includes(aidCode), aidCode)
    )
  );
  const aid = aidList[aidListResult.findIndex(r => r === true)];
  await aid.tap();

  const submitButton = await page.$(".btn.btn-sub");
  await submitButton.click();
  await page.waitFor(500);
  const closeButton = (await page.$$(".btn.btn-sub"))[1];
  await closeButton.click();
};
