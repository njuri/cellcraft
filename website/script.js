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

  const numAnimals = 1;

  for (let i = 0; i < numAnimals; i++) {
    const animal = document.createElement("div");
    animal.className = "kitten";
    document.body.appendChild(animal); // Append to body for full window movement

    // Starting position
    let x = Math.random() * (window.innerWidth - 70); // Adjust for animal size
    let y = Math.random() * (window.innerHeight - 70);

    // Random velocity
    let velocityX = Math.random() * 0.75; // Speed range -2 to 2
    let velocityY = Math.random() * 0.75;

    function updatePosition() {
      // Update position
      x += velocityX;
      y += velocityY;

      // Reflect off edges of the window
      if (x <= 0 || x >= window.innerWidth - 70) velocityX *= -1; // Adjust for animal size
      if (y <= 0 || y >= window.innerHeight - 70) velocityY *= -1;

      animal.style.left = x + "px";
      animal.style.top = y + "px";

      requestAnimationFrame(updatePosition);
    }

    updatePosition();
  }
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
  processButton.innerHTML = isProcessing ? '<div class="spinner"></div>' : '<img src="./happy-cat.gif" />Отчёт готов!';
}
