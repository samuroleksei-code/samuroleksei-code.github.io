const modal = document.getElementById("preview-modal");
const previewFrame = document.getElementById("preview-frame");
const previewTitle = document.getElementById("preview-title");
const previewDownload = document.getElementById("preview-download");
const closeControls = document.querySelectorAll("[data-close-preview]");
const previewTriggers = document.querySelectorAll(".preview-trigger");

function closePreview() {
    if (!modal) {
        return;
    }

    modal.hidden = true;
    document.body.classList.remove("preview-open");
    previewFrame.setAttribute("src", "about:blank");
}

function openPreview({ title, previewSrc, downloadSrc, downloadName }) {
    if (!modal) {
        return;
    }

    previewTitle.textContent = title;
    previewDownload.setAttribute("href", downloadSrc);
    previewDownload.setAttribute("download", downloadName || "");
    previewFrame.setAttribute("src", previewSrc);
    modal.hidden = false;
    document.body.classList.add("preview-open");
}

previewTriggers.forEach((trigger) => {
    trigger.addEventListener("click", () => {
        openPreview({
            title: trigger.dataset.previewTitle,
            previewSrc: trigger.dataset.previewSrc,
            downloadSrc: trigger.dataset.downloadSrc,
            downloadName: trigger.dataset.downloadName
        });
    });
});

closeControls.forEach((control) => {
    control.addEventListener("click", closePreview);
});

window.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && modal && !modal.hidden) {
        closePreview();
    }
});
