const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const session = require('express-session');
const path = require('path');

const app = express();
app.use(cors());

// 🟡 放大 json 数据体限制，避免 413 报错
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// 🟡 设置静态资源目录
app.use(express.static('public'));

// ✅ Render 环境 session 设置：适配 HTTPS
app.use(session({
  secret: 'excelSecret',
  resave: false,
  saveUninitialized: false,
  cookie: {
    sameSite: 'lax',
    secure: false  // 如果你用的是强制 HTTPS，可设为 true
  }
}));

// ✅ multer 使用 Render 可写目录
const upload = multer({ dest: '/tmp/' });

// 👤 假设只有一个账户
const USER = { username: 'admin', password: '123456' };

// 🔐 权限中间件
function checkAuth(req, res, next) {
  if (req.session.user) next();
  else res.status(401).send('未登录');
}

// 🔍 登录状态检查接口
app.get('/check-auth', (req, res) => {
  if (req.session.user) res.sendStatus(200);
  else res.sendStatus(401);
});

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

// 上传文件
app.post('/upload', checkAuth, upload.single('file'), (req, res) => {
  try {
    currentExcelPath = req.file.path;
    const wb = xlsx.readFile(currentExcelPath);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    res.json({ data });
  } catch (err) {
    console.error('❌ 上传失败:', err);
    res.status(500).send('上传失败: ' + err.message);
  }
});

// 保存编辑
app.post('/save', checkAuth, (req, res) => {
  try {
    const newData = req.body.data;
    if (!newData || !Array.isArray(newData)) {
      throw new Error('数据格式不合法');
    }

    const ws = xlsx.utils.aoa_to_sheet(newData);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');

    xlsx.writeFile(wb, currentExcelPath);
    console.log('✅ 保存成功:', currentExcelPath);
    res.sendStatus(200);
  } catch (err) {
    console.error('❌ 保存失败:', err);
    res.status(500).send('保存失败: ' + err.message);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
