
let tplPath = null;
let table = null;

openExcel.onclick = async () => {
  const res = await window.api.openExcel();
  if (!res) return;
  document.getElementById('excelName').textContent = res.fileName || '';
  document.getElementById('previewCard').classList.remove('d-none');
  if (table) { table.destroy(); $('#preview').empty(); }
  table = $('#preview').DataTable({
    data: res.rows,
    columns: res.headers.map(h => ({ title: h, data: h }))
  });

  // Populate and reveal the File Grouping select
  const sel = document.getElementById('fileGrouping');
  sel.innerHTML = '<option value="">(none — all files in one folder)</option>';
  res.headers.forEach(h => {
    const opt = document.createElement('option');
    opt.value = h;
    opt.textContent = h;
    sel.appendChild(opt);
  });
  document.getElementById('fileGroupingRow').classList.remove('d-none');
};

validate.onclick = async () => {
  const res = await window.api.validateSchema(tplPath);
  if (res.ok) {
    showStatus('success', 'Validation passed — all column headings match template placeholders.');
  } else {
    showStatus('danger', res.issues.map(i => '&#8226; ' + i).join('<br>'));
  }
};

selectTemplate.onclick = async () => {
  tplPath = await window.api.selectTemplate();
  document.getElementById('templateName').textContent =
    tplPath ? tplPath.split('/').pop() : 'No template selected';
};

merge.onclick = async () => {
  if (!tplPath) return showStatus('warning', 'Select a template first.');
  document.getElementById('progressCard').classList.remove('d-none');
  const groupingCol = document.getElementById('fileGrouping').value || null;
  const res = await window.api.mergeDocs(tplPath, groupingCol);
  if (!res || res.canceled) {
    document.getElementById('progressCard').classList.add('d-none');
    return;
  }
  let msg = 'Merge complete! The output folder has been opened.';
  if (res.foldersCreated > 0) {
    msg += ` ${res.foldersCreated} sub-folder${res.foldersCreated === 1 ? '' : 's'} created.`;
  }
  showStatus('success', msg);
};

window.api.onProgress(p => {
  document.getElementById('progressPct').textContent = p;
  const bar = document.getElementById('progress');
  bar.style.width = p + '%';
  bar.setAttribute('aria-valuenow', p);
  if (p >= 100) {
    setTimeout(() => document.getElementById('progressCard').classList.add('d-none'), 1500);
  }
});

function showStatus(type, message) {
  document.getElementById('statusMessage').innerHTML =
    `<div class="alert alert-${type} alert-dismissible mb-3" role="alert">
      <div>${message}</div>
      <a class="btn-close" data-bs-dismiss="alert" aria-label="close"></a>
    </div>`;
}
