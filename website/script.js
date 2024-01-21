import { processData } from "./input-procesor.js";

document.addEventListener("DOMContentLoaded", function () {
  document.getElementById("processButton").addEventListener("click", async function () {
    const dataFileInput = document.getElementById("dataFileInput");
    const imageFileInput = document.getElementById("imageFileInput");

    const dataFile = dataFileInput.files[0];
    const imageFile = imageFileInput.files[0];

    if (!dataFile || !imageFile) {
      console.error("Both files must be selected");
      return;
    }

    const dataFileArrayBuffer = await readFileAsArrayBuffer(dataFile);
    const imageFileArrayBuffer = await readFileAsArrayBuffer(imageFile);

    showSpinner(true);

    try {
      await processDataFiles(dataFileArrayBuffer, imageFileArrayBuffer, dataFile);
    } finally {
      showSpinner(false);
    }
  });
});

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(new Uint8Array(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

async function processDataFiles(dataFileArrayBuffer, imageFileArrayBuffer, dataFile) {
  const workbook = await processData(dataFileArrayBuffer, imageFileArrayBuffer);
  const buffer = await workbook.xlsx.writeBuffer();

  const fileName = dataFile.name.replace(/\.[^/.]+$/, "");
  createAndDownloadFile(buffer, `${fileName}_report.xlsx`);
}

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

function showSpinner(isProcessing) {
  const processButton = document.getElementById("processButton");
  processButton.innerHTML = isProcessing ? '<div class="spinner"></div>' : "Создать отчёт";
}
