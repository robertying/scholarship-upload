# scholarship-upload

Upload scholarship results to http://sa.tsinghua.edu.cn using Puppeteer

清华大学学生奖助资助系统上传脚本

## Usage

Upload Excel file format:

| 学号 | 姓名 | 班级 | 荣誉 | 奖学金 | 代码 | 金额 |
| :--: | :--: | :--: | :--: | :----: | :--: | :--: |


First make sure you are connected to the campus network, then use the script:

```shell
node upload.js username password filename.xlsx
```
