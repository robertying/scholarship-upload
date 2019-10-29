const puppeteer = require("puppeteer");
const Xlsx = require("xlsx");
const ProgressBar = require("progress");

const debug = true; // headless mode strangely not working
const delay = 10000; // up to how fast your computer loads pages

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

const honorsForSchoolManaged = {
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

const honorsForDepartmentManaged = {
  综合优秀奖: 0,
  学业优秀奖: 1,
  社会工作优秀奖: 2,
  体育优秀奖: 3,
  文艺优秀奖: 4,
  社会实践优秀奖: 5,
  科技创新优秀奖: 6,
  学习进步奖: 7,
  志愿公益优秀奖: 8
};

const honorsForDepartmentDistributed = {
  综合优秀奖: 0,
  学业优秀奖: 1,
  社会工作优秀奖: 2,
  体育优秀奖: 3,
  文艺优秀奖: 4,
  社会实践优秀奖: 5,
  科技创新优秀奖: 6,
  无校级荣誉: 7,
  学习进步奖: 8,
  志愿公益优秀奖: 9
};

const scholarships = {
  "清华之友——华为奖学金": [
    { code: "J2022050", amount: 5000, type: "校管校分" }
  ],
  "清华之友——张明为奖学金": [
    { code: "J2072040", amount: 4000, type: "校管校分" }
  ],
  "中国宋庆龄基金会·中芯国际孟宁奖助学金": [
    {
      code: "J2102100",
      amount: 10000,
      type: "校管校分"
    }
  ],
  "清华之友——董氏东方奖学金": [
    { code: "J2162050", amount: 5000, type: "校管校分" }
  ],
  "清华之友——周惠琪奖学金": [
    { code: "J2202050", amount: 5000, type: "校管校分" }
  ],
  "清华之友——POSCO奖学金": [
    { code: "J2282070", amount: 7000, type: "校管校分" }
  ],
  "清华之友——黄乾亨奖学金": [
    { code: "J2362020", amount: 2000, type: "校管校分" }
  ],
  "清华之友——苏州工业园区奖学金": [
    { code: "J2462080", amount: 8000, type: "校管校分" }
  ],
  "清华之友——恒大奖学金": [
    { code: "J2532050", amount: 5000, type: "校管校分" }
  ],
  工商银行奖学金: [{ code: "J2542100", amount: 10000, type: "校管校分" }],
  "清华之友——深交所奖学金": [
    { code: "J2562050", amount: 5000, type: "校管校分" }
  ],
  国家奖学金: [{ code: "J2602080", amount: 8000, type: "校管校分" }],
  "清华之友——丰田奖学金": [
    { code: "J2612030", amount: 3000, type: "校管校分" },
    { code: "J2612050", amount: 5000, type: "校管校分" }
  ],
  "清华之友——SK奖学金": [{ code: "J2622060", amount: 6000, type: "校管校分" }],
  "清华之友——三星奖学金": [
    { code: "J2652050", amount: 5000, type: "校管校分" }
  ],
  "清华之友——郑格如奖学金": [
    { code: "J2722020", amount: 2000, type: "校管校分" }
  ],
  "清华之友——广药集团奖学金": [
    { code: "J2782030", amount: 3000, type: "校管校分" },
    { code: "J2782050", amount: 5000, type: "校管校分" }
  ],
  "清华之友——黄奕聪伉俪奖助学金": [
    { code: "J2802040", amount: 4000, type: "校管校分" }
  ],
  国家励志奖学金: [{ code: "J2892050", amount: 5000, type: "校管校分" }],
  "清华之友——渠玉芝奖学金": [
    { code: "J2902020", amount: 2000, type: "校管校分" }
  ],
  蒋南翔奖学金: [{ code: "J3012150", amount: 15000, type: "校管校分" }],
  "一二·九奖学金": [{ code: "J3022150", amount: 15000, type: "校管校分" }],
  好读书奖学金: [
    { code: "J3032030", amount: 3000, type: "校管校分" },
    { code: "J3032080", amount: 8000, type: "校管校分" }
  ],
  "清华校友——孟昭英奖学金": [
    { code: "J3122030", amount: 3000, type: "校管校分" }
  ],
  电子系97级校友奖学金: [{ code: "J7237030", amount: 1800, type: "院管院分" }],
  电子系1998级校友奖学基金: [
    { code: "J7232020", amount: 1800, type: "院管院分" }
  ],
  常锋奖学金: [{ code: "J7235020", amount: 2000, type: "院管院分" }],
  校设奖学金: [
    { code: "J1022000", amount: 0, type: "校管院分" },
    { code: "J1022010", amount: 1000, type: "校管院分" },
    { code: "J1022020", amount: 2000, type: "校管院分" },
    { code: "J1022030", amount: 3000, type: "校管院分" },
    { code: "J1022040", amount: 4000, type: "校管院分" },
    { code: "J1022050", amount: 5000, type: "校管院分" }
  ],
  "2018级新生一等奖学金": [
    { code: "J1142125", amount: 12500, type: "校管校分" }
  ],
  "2018级新生二等奖学金": [
    { code: "J1142050", amount: 5000, type: "校管校分" }
  ],
  "2017级新生一等奖学金": [
    { code: "J1152125", amount: 12500, type: "校管校分" }
  ],
  "2017级新生二等奖学金": [
    { code: "J1152050", amount: 5000, type: "校管校分" }
  ],
  "2016级新生一等奖学金": [
    { code: "J1162125", amount: 12500, type: "校管校分" }
  ],
  "2016级新生二等奖学金": [
    { code: "J1162050", amount: 5000, type: "校管校分" }
  ],
  "2019级新生一等奖学金": [
    { code: "J1172125", amount: 12500, type: "校管校分" }
  ],
  "2019级新生二等奖学金": [{ code: "J1172050", amount: 5000, type: "校管校分" }]
};

const schoolManagedApplications = applications.filter(
  i => scholarships[i[4].trim()][0].type === "校管校分"
);
const departmentManagedApplications = applications.filter(
  i => scholarships[i[4].trim()][0].type === "校管院分"
);
const departmentDistributedApplications = applications.filter(
  i => scholarships[i[4].trim()][0].type === "院管院分"
);

(async () => {
  const browser = await puppeteer.launch({
    headless: !debug
  });
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 720 });

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
  const scholarshipTab = (await page.$$(".fir_li"))[1];
  await scholarshipTab.click();
  await page.waitFor(1000);

  bar.interrupt("[progress] start to upload 校管校分...");
  const schoolManagedTab = await page.waitForSelector(
    'a[data="/f/bksjzd/jxj/v_bj_jxj_xflrb/xsgzz/beforePageList"]'
  );
  await schoolManagedTab.click();
  await page.waitFor(1000);

  let i = 0;
  for (const application of schoolManagedApplications) {
    try {
      const studentId = application[0].toString();
      const honor = application[3].trim();
      const scholarshipCode = application[5].trim();
      await fillForm(page, "校管校分", studentId, honor, scholarshipCode);
      i++;
      bar.tick();
      break;
    } catch (e) {
      console.error("Upload failed at:");
      console.error(application);
      throw e;
    }
  }

  bar.interrupt("[progress] start to upload 校管院分...");
  let closeButton = (await page.$$(".btn"))[5];
  await closeButton.click();
  const departmentManagedTab = await page.waitForSelector(
    'a[data="/f/bksjzd/jxj/v_bj_jxj_yflrb/xsgzz/beforePageList"]'
  );
  await departmentManagedTab.click();
  await page.waitFor(1000);

  for (const application of departmentManagedApplications) {
    try {
      const studentId = application[0].toString();
      const honor = application[3].trim();
      const scholarshipCode = application[5].trim();
      await fillForm(page, "校管院分", studentId, honor, scholarshipCode);
      i++;
      bar.tick();
    } catch (e) {
      console.error("Upload failed at:");
      console.error(application);
      throw e;
    }
  }

  bar.interrupt("[progress] start to upload 院管院分...");
  closeButton = (await page.$$(".btn"))[5];
  await closeButton.click();
  const departmentDistributedTab = await page.waitForSelector(
    'a[data="/f/bksjzd/jxj/v_bj_jxj_yglrb/xsgzz/beforePageList"]'
  );
  await departmentDistributedTab.click();
  await page.waitFor(1000);

  for (const application of departmentDistributedApplications) {
    try {
      const studentId = application[0].toString();
      const honor = application[3].trim();
      const scholarshipCode = application[5].trim();
      await fillForm(page, "院管院分", studentId, honor, scholarshipCode);
      i++;
      bar.tick();
    } catch (e) {
      console.error("Upload failed at:");
      console.error(application);
      throw e;
    }
  }

  bar.interrupt("Upload finished!");
  await browser.close();
})();

const fillForm = async (page, type, studentId, honorName, scholarshipCode) => {
  const addButton = await page.waitForSelector("#addbtn");
  await addButton.click();
  const idInput = await page.waitForSelector("#q_xh");
  await idInput.evaluate((e, studentId) => (e.value = studentId), studentId);
  const searchButton = await page.$(".a_btn");
  await searchButton.click();

  await page.waitFor(500);
  const honorSelection = await page.$("#select2-jllxm-container");
  await honorSelection.tap();
  await page.waitForSelector(".select2-results__option");
  const honorList = await page.$$(".select2-results__option");
  const honors =
    type === "校管校分"
      ? honorsForSchoolManaged
      : type === "校管院分"
      ? honorsForDepartmentManaged
      : honorsForDepartmentDistributed;
  await honorList[honors[honorName]].tap();

  const scholarshipSelection = await page.$("#select2-dm-container");
  await scholarshipSelection.tap();
  await page.waitFor(".select2-results__option");
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
