const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const session = require('express-session');
const path = require('path');

const app = express();
app.use(cors());

// 🟢 增加 JSON 请求体大小限制，避免 413 报错
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

app.use(express.static('public'));

app.use(session({
  secret: 'excelSecret',
  resave: false,
  saveUninitialized: true
}));

// ✅ 使用 Render 支持写入的临时目录
const upload = multer({ dest: '/tmp/' });

const USER = { username: 'admin', password: '123456' };

function checkAuth(req, res, next) {
  if (req.session.user) next();
  else res.status(401).send('请登录');
}

// 登录接口
app.post('/login', (req, res) => {
  const { username, password } = req.body;
  if (username === USER.username && password === USER.password) {
    req.session.user = username;
    res.sendStatus(200);
  } else {
    res.status(401).send('登录失败');
  }
});

let currentExcelPath = '';

// 上传 Excel 并读取内容
app.post('/upload', checkAuth, upload.single('file'), (req, res) => {
  try {
    currentExcelPath = req.file.path;
    const wb = xlsx.readFile(currentExcelPath);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    res.json({ data });
  } catch (err) {
    console.error('读取 Excel 出错:', err);
    res.status(500).send('读取失败：' + err.message);
  }
});

// 保存修改后的 Excel
app.post('/save', checkAuth, (req, res) => {
  try {
    const newData = req.body.data;

    if (!newData || !Array.isArray(newData)) {
      throw new Error('提交的数据格式无效');
    }

    const ws = xlsx.utils.aoa_to_sheet(newData);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');

    xlsx.writeFile(wb, currentExcelPath);
    console.log('✅ 保存成功:', currentExcelPath);
    res.sendStatus(200);
  } catch (err) {
    console.error('❌ 保存失败:', err);
    res.status(500).send('保存失败：' + err.message);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
