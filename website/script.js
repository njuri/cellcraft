import { processData } from "./input-procesor.js";

document.addEventListener("DOMContentLoaded", function () {
  document.getElementById("processButton").addEventListener("click", function () {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput.files[0];

    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);

      // Writing back to binary
      const workbook = processData(data);
      const processedWorkbook = XLSX.write(workbook, { bookType: "xlsx", type: "binary" });

      function s2ab(s) {
        const buffer = new ArrayBuffer(s.length);
        const view = new Uint8Array(buffer);
        for (let i = 0; i < s.length; i++) {
          view[i] = s.charCodeAt(i) & 0xff;
        }
        return buffer;
      }

      const processedData = s2ab(processedWorkbook);
      const fileName = file.name.replace(/\.[^/.]+$/, "");
      createAndDownloadFile(processedData, `${fileName}_report.xlsx`);
    };
    reader.readAsArrayBuffer(file);
  });
});

function createAndDownloadFile(data, fileName) {
  const blob = new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const url = URL.createObjectURL(blob);
  const downloadLink = document.getElementById("downloadLink");
  downloadLink.style.display = "block";

  downloadLink.addEventListener("click", function () {
    const tempLink = document.createElement("a");
    tempLink.href = url;
    tempLink.download = fileName;
    tempLink.style.display = "none";
    document.body.appendChild(tempLink);
    tempLink.click();
    document.body.removeChild(tempLink);
  });
}
