const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const session = require('express-session');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

app.use(session({
  secret: 'excelSecret',
  resave: false,
  saveUninitialized: true
}));

const upload = multer({ dest: path.join(__dirname, 'uploads/') });

const USER = { username: 'admin', password: '123456' }; // 需要时可扩展到数据库

function checkAuth(req, res, next) {
  if (req.session.user) next();
  else res.status(401).send('请登录');
}

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

app.post('/upload', checkAuth, upload.single('file'), (req, res) => {
  currentExcelPath = req.file.path;
  const wb = xlsx.readFile(currentExcelPath);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  res.json({ data });
});

app.post('/save', checkAuth, (req, res) => {
  const newData = req.body.data;
  const ws = xlsx.utils.aoa_to_sheet(newData);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
  xlsx.writeFile(wb, currentExcelPath);
  res.sendStatus(200);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
