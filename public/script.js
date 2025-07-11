let currentData = [];

function redirectIfNotLogged() {
  fetch('/upload', { method: 'POST' })
    .then(res => { if (res.status === 401) window.location = '/login.html'; });
}
redirectIfNotLogged();

function uploadExcel() {
  const f = document.getElementById('excelFile').files[0];
  if (!f) return alert('请选择文件');
  const fd = new FormData();
  fd.append('file', f);
  fetch('/upload', { method: 'POST', body: fd })
    .then(res => res.json())
    .then(j => { currentData = j.data; renderTable(); });
}

function renderTable() {
  const c = document.getElementById('tableContainer');
  c.innerHTML = '';
  const t = document.createElement('table');
  currentData.forEach((row, i) => {
    const tr = document.createElement('tr');
    row.forEach((cell, j) => {
      const td = document.createElement('td');
      td.contentEditable = "true";
      td.textContent = cell;
      td.addEventListener('input', () => currentData[i][j] = td.textContent);
      tr.appendChild(td);
    });
    t.appendChild(tr);
  });
  c.appendChild(t);
}

function saveExcel() {
  fetch('/save', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ data: currentData })
  })
  .then(res => res.ok ? alert('保存成功！') : alert('保存失败'));
}
