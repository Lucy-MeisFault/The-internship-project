  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const fileChosen = document.getElementById('fileChosen');
  const fileName = document.getElementById('fileName');
  const browseBtn = document.getElementById('browseBtn');
  const clearFile = document.getElementById('clearFile');

  browseBtn.onclick = (e) => { e.stopPropagation(); fileInput.click(); };
  dropZone.onclick = () => fileInput.click();

  fileInput.onchange = () => {
    if (fileInput.files[0]) showFile(fileInput.files[0].name);
  };

  dropZone.ondragover = (e) => { e.preventDefault(); dropZone.classList.add('drag'); };
  dropZone.ondragleave = () => dropZone.classList.remove('drag');
dropZone.ondrop = (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag');

    if (e.dataTransfer.files[0]) {
        fileInput.files = e.dataTransfer.files;
        showFile(e.dataTransfer.files[0].name);
    }
};

  function showFile(name) {
    fileName.textContent = name;
    fileChosen.style.display = 'flex';
    dropZone.style.display = 'none';
  }

  clearFile.onclick = () => {
    fileInput.value = '';
    fileChosen.style.display = 'none';
    dropZone.style.display = 'flex';
  };

document.getElementById('generateBtn').onclick = () => {
if (fileInput.value === '') {
    return;
}


  const file = document.getElementById('fileInput').files[0];
  const formData = new FormData();
  formData.append('file', file);

fetch('/upload', {
    method: 'POST',
    body: formData
})
.then(r => r.blob())
.then(blob => {
    const url = window.URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'report.pptx';

    document.body.appendChild(a);
    a.click();
    a.remove();

    window.URL.revokeObjectURL(url);
});
}