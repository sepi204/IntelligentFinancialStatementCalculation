const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");
const browseBtn = document.getElementById("browse-btn");
const uploadBtn = document.getElementById("upload-btn");
const progressBar = document.getElementById("progress-bar");
const progressText = document.getElementById("progress-text");
const alertBox = document.getElementById("alert");
const downloadLink = document.getElementById("download-link");

let selectedFile = null;

const resetAlert = () => {
    alertBox.classList.add("hidden");
    alertBox.classList.remove("success", "error");
    alertBox.textContent = "";
};

const showAlert = (message, type) => {
    alertBox.textContent = message;
    alertBox.classList.remove("hidden");
    alertBox.classList.toggle("success", type === "success");
    alertBox.classList.toggle("error", type === "error");
};

const setProgress = (value, text) => {
    progressBar.style.width = `${value}%`;
    progressText.textContent = text;
};

const setControlsEnabled = (enabled) => {
    uploadBtn.disabled = !enabled;
    browseBtn.disabled = !enabled;
};

const handleFiles = (files) => {
    if (!files || !files.length) {
        return;
    }
    selectedFile = files[0];
    downloadLink.classList.add("hidden");
    resetAlert();
    setProgress(0, `فایل انتخاب شد: ${selectedFile.name}`);
};

browseBtn.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", (e) => handleFiles(e.target.files));

dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));

dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    handleFiles(e.dataTransfer.files);
});

uploadBtn.addEventListener("click", () => {
    if (!selectedFile) {
        showAlert("لطفاً ابتدا فایل اکسل را انتخاب کنید.", "error");
        return;
    }

    const formData = new FormData();
    formData.append("UploadedFile", selectedFile);

    resetAlert();
    downloadLink.classList.add("hidden");
    setControlsEnabled(false);
    setProgress(0, "در حال شروع آپلود...");

    const xhr = new XMLHttpRequest();
    xhr.open("POST", "?handler=Upload", true);
    xhr.responseType = "blob";

    xhr.upload.onprogress = (event) => {
        if (event.lengthComputable) {
            const percent = Math.round((event.loaded / event.total) * 100);
            setProgress(percent, `آپلود ${percent}%`);
        }
    };

    xhr.onload = () => {
        setControlsEnabled(true);
        if (xhr.status >= 200 && xhr.status < 300) {
            const blob = xhr.response;
            const url = window.URL.createObjectURL(blob);
            const header = xhr.getResponseHeader("Content-Disposition") ?? "";
            const match = header.match(/filename="?(.+)"?/i);
            const fileName = match ? decodeURIComponent(match[1]) : "Reports.xlsx";

            downloadLink.href = url;
            downloadLink.download = fileName;
            downloadLink.classList.remove("hidden");
            setProgress(100, "پردازش کامل شد.");
            showAlert("پردازش با موفقیت انجام شد. فایل را دریافت کنید.", "success");
        } else {
            const reader = new FileReader();
            reader.onload = () => {
                showAlert(`خطا: ${reader.result || "مشکل در پردازش"}`, "error");
                setProgress(0, "خطا رخ داد.");
            };
            reader.readAsText(xhr.response);
        }
    };

    xhr.onerror = () => {
        setControlsEnabled(true);
        showAlert("خطا در ارتباط با سرور رخ داد.", "error");
        setProgress(0, "خطا رخ داد.");
    };

    xhr.send(formData);
});

