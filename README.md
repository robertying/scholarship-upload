# scholarship-upload

Upload scholarship & financial aid results to http://sa.tsinghua.edu.cn using Puppeteer

清华大学学生奖助资助系统上传脚本

## Usage

### Scholarships

Upload Excel file format:

| 学号 | 姓名 | 班级 | 荣誉 | 奖学金 | 代码 | 金额 |
| :--: | :--: | :--: | :--: | :----: | :--: | :--: |


First make sure you are connected to the campus network, then use the script:

```shell
node scholarship.js username password filename.xlsx
```

### Financial Aids

Upload Excel file format:

| 学号 | 姓名 | 班级 | 助学金 | 代码 | 金额 |
| :--: | :--: | :--: | :----: | :--: | :--: |


First make sure you are connected to the campus network, then use the script:

```shell
node aid.js username password filename.xlsx
```

## Caveats

You may need to keep Chromium active by not hiding it as a background window to reduce possible failure
