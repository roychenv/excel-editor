let currentData = [];

function uploadExcel() {
  const file = document.getElementById('excelFile').files[0];
  if (!file) return alert('请选择文件');

  const formData = new FormData();
  formData.append('file', file);

  fetch('/upload', {
    method: 'POST',
    body: formData
  })
  .then(res => {
    if (!res.ok) throw new Error('上传失败');
    return res.json();
  })
  .then(res => {
    currentData = res.data;
    renderTable();
  })
  .catch(err => alert('❌ 上传失败: ' + err.message));
}

function renderTable() {
  const container = document.getElementById('tableContainer');
  container.innerHTML = '';
  const table = document.createElement('table');

  currentData.forEach((row, i) => {
    const tr = document.createElement('tr');
    row.forEach((cell, j) => {
      const td = document.createElement('td');
      td.contentEditable = "true";
      td.textContent = cell;
      td.addEventListener('input', () => {
        currentData[i][j] = td.textContent;
      });
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });

  container.appendChild(table);
}

function saveExcel() {
  fetch('/save', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ data: currentData })
  })
  .then(async res => {
    if (res.ok) {
      alert('✅ 保存成功！');
    } else {
      const errorText = await res.text();
      alert('❌ ' + errorText);
    }
  })
  .catch(err => alert('❌ 请求错误: ' + err.message));
}
