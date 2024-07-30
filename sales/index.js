const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const path = require("path");
const cors = require("cors");
const fs = require("fs");
const { v4: uuidv4 } = require('uuid');

const app = express();
const upload = multer({ dest: "uploads/" });

// 使用cors中间件
app.use(cors());

app.post("/upload", upload.single("file"), (req, res) => {
  const filePath = path.join(__dirname, req.file.path);
  const workbook = xlsx.readFile(filePath);

  const sheetsData = {};
  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    
    // 检查 worksheet['!ref'] 是否存在
    if (!worksheet['!ref']) {
      return;
    }

    // 获取双层表头信息
    const headers = [[], []];
    const range = xlsx.utils.decode_range(worksheet['!ref']);
    for (let row = 0; row < 2; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = xlsx.utils.encode_cell({ r: range.s.r + row, c: col });
        const cell = worksheet[cellAddress];
        headers[row].push(cell ? cell.v : undefined);
      }
    }

    // 将Excel中的数据转换为JSON，并去掉表头
    const jsonSheet = xlsx.utils.sheet_to_json(worksheet, {
      raw: false,
      dateNF: "yyyy-mm-dd",
      cellDates: true,
      header: 1,
    }).slice(2); // 移除前两行表头

    const formattedSheet = jsonSheet.map(row => {
      const rowData = {};
      headers[0].forEach((header1, index) => {
        const header2 = headers[1][index];
        const key = header1 || header2;
        rowData[key] = row[index];
        if (Object.prototype.toString.call(row[index]) === "[object Date]") {
          rowData[key] = row[index].toISOString().split("T")[0];
        }
      });
      rowData["id"] = uuidv4();
      return rowData;
    });

    sheetsData[sheetName] = { data: formattedSheet };
  });

  res.json(sheetsData);

  // 删除上传的文件
  fs.unlink(filePath, (err) => {
    if (err) {
      console.error(`Error deleting file: ${err}`);
    } else {
      console.log(`File ${filePath} deleted successfully.`);
    }
  });
});

app.listen(3001, () => {
  console.log("Server started on http://localhost:3001");
});
