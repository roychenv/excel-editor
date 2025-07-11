let currentData = [];

function checkLoginStatus() {
  return fetch('/check-auth')
    .then(res => {
      if (res.status === 401) {
        window.location.href = '/login.html';
        throw new Error('未登录');
      }
    });
}

function uploadExcel() {
  const file = document.getElementById('excelFile').files[0];
  if (!file) return alert('请选择文件');

  checkLoginStatus().then(() => {
    const formData = new FormData();
    formData.append('file', file);

    fetch('/upload', {
      method: 'POST',
      body: formData
    })
    .then(async res => {
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    })
    .then(res => {
      currentData = res.data;
      renderTable();
    })
    .catch(err => {
      alert('❌ 上传失败: ' + err.message);
      console.error(err);
    });
  });
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
  checkLoginStatus().then(() => {
    fetch('/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data: currentData })
    })
    .then(async res => {
      if (!res.ok) throw new Error(await res.text());
      alert('✅ 保存成功！');
    })
    .catch(err => {
      alert('❌ 保存失败: ' + err.message);
      console.error(err);
    });
  });
}

// ✅ 导出 Excel 到本地
function downloadExcel() {
  if (!currentData || currentData.length === 0) {
    alert("表格为空，无法导出！");
    return;
  }

  const ws = XLSX.utils.aoa_to_sheet(currentData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

  XLSX.writeFile(wb, "编辑结果.xlsx");
}
