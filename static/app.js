const fileInput = document.querySelector("#file-input");
const dropzone = document.querySelector("#dropzone");
const fileTitle = document.querySelector("#file-title");
const fileMeta = document.querySelector("#file-meta");
const thresholdInput = document.querySelector("#threshold");
const includeEmailsInput = document.querySelector("#include-emails");
const generateButton = document.querySelector("#generate-btn");
const message = document.querySelector("#message");
const apiStatus = document.querySelector("#api-status");
const resultTitle = document.querySelector("#result-title");
const resultMeta = document.querySelector("#result-meta");
const downloadLink = document.querySelector("#download-link");

let selectedFile = null;
let downloadUrl = null;

function setFile(file) {
  selectedFile = file;
  fileTitle.textContent = file.name;
  fileMeta.textContent = `${(file.size / 1024).toFixed(1)} KB`;
  generateButton.disabled = false;
}

function setMessage(text, type = "") {
  message.textContent = text;
  message.className = `message ${type}`.trim();
}

fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files;
  if (file) setFile(file);
});

["dragenter", "dragover"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragging");
  });
});

dropzone.addEventListener("drop", (event) => {
  const [file] = event.dataTransfer.files;
  if (file) setFile(file);
});

generateButton.addEventListener("click", async () => {
  if (!selectedFile) return;

  generateButton.disabled = true;
  generateButton.textContent = "Generating report...";
  setMessage("Generating PDF. This can take a little while.");

  const formData = new FormData();
  formData.append("file", selectedFile);
  formData.append("threshold", thresholdInput.value || "75");
  formData.append("include_emails", includeEmailsInput.checked ? "true" : "false");

  try {
    const response = await fetch("/reports", { method: "POST", body: formData });
    if (!response.ok) {
      const payload = await response.json();
      throw new Error(payload.detail || "Report generation failed.");
    }

    const blob = await response.blob();
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    downloadUrl = URL.createObjectURL(blob);

    const pdfName = `${selectedFile.name.replace(/\.xlsx$/i, "")}.pdf`;
    downloadLink.href = downloadUrl;
    downloadLink.download = pdfName;
    downloadLink.hidden = false;
    resultTitle.textContent = pdfName;
    resultMeta.textContent = `${(blob.size / 1024).toFixed(1)} KB PDF ready`;
    setMessage("PDF ready to download.", "success");
  } catch (error) {
    setMessage(error.message, "error");
  } finally {
    generateButton.disabled = false;
    generateButton.innerHTML = `
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <path d="M12 3v12"></path>
        <path d="m7 10 5 5 5-5"></path>
        <path d="M5 21h14"></path>
      </svg>
      Generate PDF
    `;
  }
});

fetch("/health")
  .then((response) => {
    if (!response.ok) throw new Error();
    apiStatus.textContent = "API ready";
    apiStatus.classList.add("ready");
  })
  .catch(() => {
    apiStatus.textContent = "API unavailable";
  });
