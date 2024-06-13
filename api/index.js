const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const path = require("path");
const cors = require("cors");
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
    const jsonSheet = xlsx.utils.sheet_to_json(worksheet, {
      raw: false,
      dateNF: "yyyy-mm-dd",
      cellDates: true,
    });

    // 将Excel中的日期转换为字符串
    jsonSheet.forEach((row) => {
      for (const key in row) {
        if (Object.prototype.toString.call(row[key]) === "[object Date]") {
          row[key] = row[key].toISOString().split("T")[0];
        }
        row["id"] = uuidv4()
      }
    });

    sheetsData[sheetName] = { data: jsonSheet };
  });

  res.json(sheetsData);
});

app.listen(3001, () => {
  console.log("Server started on http://localhost:3001");
});
