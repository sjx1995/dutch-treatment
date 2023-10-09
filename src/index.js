/*
 * @Description:
 * @Author: Sunly
 * @Date: 2023-09-28 08:11:07
 */
import fs from "fs";
import path from "path";
import excelJs from "exceljs";
import dayjs from "dayjs";
import { ALIPAY_PATH, readAliPayCSV } from "./reduceAliPay.js";
import { WECHAT_PATH, readWeChatCSV } from "./reduceWeChat.js";

function isExist(filePath) {
  return fs.existsSync(filePath);
}

function getTotalPerson() {
  const persons = process.argv.slice(2, 3);
  if (!persons.length) {
    throw new Error("请输入人数");
  }
  return parseInt(persons[0]);
}

// async function readExcel() {
//   const workbook = new excelJs.Workbook();
//   await workbook.xlsx.readFile(file, { parserOptions: { encoding: "utf8" } });

//   const worksheet = workbook.worksheets[0];

//   return worksheet;
// }

// 处理数据
async function reduceData(data) {
  let total = 0;
  const types = {};
  const excelData = [];

  data.forEach((row) => {
    let time = row[0];
    const type = row[1];
    const seller = row[2];
    const goods = row[3];
    let money = row[4];

    if (typeof money !== "number") {
      money = parseFloat(money);
    }

    // 处理类别
    if (!types[type]) {
      types[type] = 0;
    }
    types[type] += money;

    // 格式化时间
    time = dayjs(time).format("YYYY-MM-DD HH:mm");

    // 总金额
    total += money;

    excelData.push([time, type, seller, goods, money]);
  });

  await generateExcel(excelData);

  return { types, total };
}

async function generateExcel(data) {
  const workbook = new excelJs.Workbook();
  const worksheet = workbook.addWorksheet("账单");

  worksheet.columns = [
    { header: "交易时间", key: "time", width: 25 },
    { header: "交易类型", key: "type", width: 12 },
    { header: "交易对象", key: "seller", width: 40 },
    { header: "商品", key: "goods", width: 40 },
    { header: "金额", key: "money", width: 10 },
  ];

  // 时间正序
  data.sort((a, b) => {
    return dayjs(a[0]).valueOf() - dayjs(b[0]).valueOf();
  });

  data.forEach((row) => {
    worksheet.addRow(row);
  });

  await workbook.xlsx.writeFile("bill.xlsx");
}

(async function main() {
  // 读取总人数
  const totalPerson = getTotalPerson();

  // 读取对应的文件
  const res = [];
  if (isExist(ALIPAY_PATH)) {
    res.push(...(await readAliPayCSV()));
  }
  if (isExist(WECHAT_PATH)) {
    res.push(...(await readWeChatCSV()));
  }

  if (!res.length) {
    throw new Error("没有找到账单，请确保账单存在且csv文件命名正确");
  }

  // 处理excel
  const { types, total } = await reduceData(res);

  // 输出
  for (const type in types) {
    console.log(
      `[${type}] ${types[type].toFixed(2)}  (${(
        (types[type] / total) *
        100
      ).toFixed(2)}%)`
    );
  }
  console.log(`\n[总金额] ${total.toFixed(2)}`);
  console.log(`[💰人均] ${(total / totalPerson).toFixed(2)}`);
})();
