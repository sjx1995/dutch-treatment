/*
 * @Description: 处理支付宝账单
 * @Author: Sunly
 * @Date: 2023-09-28 09:45:07
 */
import path from "path";
import excelJs from "exceljs";
import iconv from "iconv-lite";

const file = path.resolve(process.cwd(), "alipay.csv");

async function readAliPayCSV() {
  let data = [];
  const workbook = new excelJs.Workbook();
  const content = await workbook.csv.readFile(file, {
    parserOptions: { encoding: "binary" },
  });
  content.eachRow((row, i) => {
    if (i !== 1) {
      const rowData = [];
      row.eachCell(({ value }, j) => {
        if (j === 1 || j === 2 || j === 3 || j === 5 || j === 7) {
          if (typeof value === "string") {
            value = gb2312ToUtf8(value);
          }
          rowData.push(value);
        }
      });
      data.push(rowData);
    }
  });
  return data;
}

function gb2312ToUtf8(str) {
  const buf = Buffer.from(str, "binary");
  const str1 = iconv.decode(buf, "gb2312");
  const res = iconv.encode(str1, "utf8").toString();
  return res;
}

export { file as ALIPAY_PATH, readAliPayCSV };
