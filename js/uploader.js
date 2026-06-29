// uploader.js — drag-and-drop Excel file upload handler
(function () {
  'use strict';

  document.addEventListener('DOMContentLoaded', function () {
    var dropZone = document.getElementById('dropZone');
    var fileInput = document.getElementById('fileInput');
    if (!dropZone || !fileInput) return;

    dropZone.addEventListener('click', function () { fileInput.click(); });
    dropZone.addEventListener('dragover', function (e) { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', function () { dropZone.classList.remove('dragover'); });
    dropZone.addEventListener('drop', function (e) {
      e.preventDefault();
      dropZone.classList.remove('dragover');
      var files = e.dataTransfer.files;
      if (files.length) fileInput.files = files;
      fileInput.dispatchEvent(new Event('change'));
    });

    fileInput.addEventListener('change', function () {
      var file = fileInput.files[0];
      if (!file) return;
      if (window.handleFileUpload) window.handleFileUpload(file);
      dropZone.style.display = 'none';
    });
  });
})();
