/*
 * @Description: 处理微信账单
 * @Author: Sunly
 * @Date: 2023-10-08 03:29:13
 */
import path from "path";
import excelJs from "exceljs";

const file = path.resolve(process.cwd(), "wechat.csv");

async function readWeChatCSV() {
  let data = [];
  const workbook = new excelJs.Workbook();
  const content = await workbook.csv.readFile(file, {
    parserOptions: { encoding: "utf8" },
  });
  content.eachRow((row, i) => {
    if (i !== 1) {
      const rowData = [];
      row.eachCell(({ value }, j) => {
        if (j === 1 || j === 2 || j === 3 || j === 4) {
          rowData.push(value);
        }
        if (j === 6) {
          if (value.startsWith("¥")) {
            value = value.slice(1);
          }
          rowData.push(parseFloat(value));
        }
      });
      data.push(rowData);
    }
  });
  return data;
}

export { file as WECHAT_PATH, readWeChatCSV };
