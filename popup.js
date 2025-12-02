// global: display info from manifest
(function () {
    const manifest = chrome.runtime.getManifest(); // get manifest object
    const versionElement = document.getElementById("ext-ver");
    if (versionElement) versionElement.textContent = manifest.version; // extension version
    const nameElement = document.getElementById("ext-name");
    if (nameElement) nameElement.textContent = manifest.name; // extension name
    document.title = manifest.name; // set popup HTML title
    const repoLink = document.getElementById("ext-repo");
    if (repoLink && manifest.homepage_url) repoLink.href = manifest.homepage_url; // extension repo URL
})();

// global: selectors
const scrapeStatus = document.getElementById("qd-status");
const scrapeBtn = document.getElementById("qd-scrapeBtn");
const imgCheckbox = document.getElementById("qd-outFileIncludeImage");
const swapCheckbox = document.getElementById("qd-outFileContentOrder");
const filenameInput = document.getElementById("qd-outFileNamePattern");
const formatSelect = document.getElementById("qd-outFileNameExt");
const formatEnableCheckbox = document.getElementById("qd-enableCustomFormat");

// global: load saved configuration from storage
document.addEventListener("DOMContentLoaded", () => {
    chrome.storage.local.get(["savedFilenamePattern", "savedOutputFormat", "savedIncludeImage", "savedSwapState"], (result) => {
        if (result.savedFilenamePattern) filenameInput.value = result.savedFilenamePattern;
        if (result.savedOutputFormat) formatSelect.value = result.savedOutputFormat;
        if (result.savedIncludeImage !== undefined) imgCheckbox.checked = result.savedIncludeImage;
        if (result.savedSwapState !== undefined) swapCheckbox.checked = result.savedSwapState;
        if (result.savedEnableCustomFormat !== undefined) {
            formatEnableCheckbox.checked = result.savedEnableCustomFormat;
            toggleFormatSelect();
        }
    });
});

// global: save configuration to storage
function saveOptions() {
    chrome.storage.local.set({
        savedFilenamePattern: filenameInput.value,
        savedOutputFormat: formatSelect.value,
        savedIncludeImage: imgCheckbox.checked,
        savedSwapState: swapCheckbox.checked,
        savedEnableCustomFormat: formatEnableCheckbox.checked,
    });
}

function toggleFormatSelect() {
    formatSelect.disabled = !formatEnableCheckbox.checked;
}

filenameInput.addEventListener("input", saveOptions);
formatSelect.addEventListener("change", saveOptions);
imgCheckbox.addEventListener("change", saveOptions);
swapCheckbox.addEventListener("change", saveOptions);
formatEnableCheckbox.addEventListener("change", () => {
    toggleFormatSelect();
    saveOptions();
});

// global: process scraping and download result
if (scrapeBtn) {
    scrapeBtn.addEventListener("click", async () => {
        let [tab] = await chrome.tabs.query({ active: true, currentWindow: true }); // get current active tab

        // validate Quizlet page URL
        if (!tab.url.includes("quizlet.com")) {
            scrapeStatus.textContent = "Error: Please open a Quizlet set page.";
            scrapeStatus.style.color = "crimson";
            return;
        }

        scrapeStatus.textContent = "Scraping data...";
        scrapeStatus.style.color = "yellow";

        const isSwapped = swapCheckbox ? swapCheckbox.checked : false; // check swap option status
        const includeImg = imgCheckbox ? imgCheckbox.checked : true; // check image option status

        try {
            const response = await chrome.tabs.sendMessage(tab.id, {
                action: "scrape_quizlet",
                swap: isSwapped,
            });

            if (response && response.status === "success") {
                const payload = response.payload;
                const outputFormat = formatEnableCheckbox.checked ? formatSelect.value : "json";

                const filenamePattern = filenameInput ? filenameInput.value : "{quizSetTitle}_YYYY-MM-DD_HH-mm-ss_{swapState}";
                const finalFilename = generateFilename(filenamePattern, payload.info, outputFormat);

                const { content, mimeType } = await formatData(payload, outputFormat, includeImg);
                await downloadFile(content, finalFilename, mimeType);

                scrapeStatus.textContent = `Done! Found ${payload.info.numberOfQuizzes} quizzes.`;
                scrapeStatus.style.color = "forestgreen";
            } else {
                scrapeStatus.textContent = "Error: No data found or script failed.";
                scrapeStatus.style.color = "crimson";
            }
        } catch (error) {
            console.error(error);
            scrapeStatus.textContent = "Error: Page not ready or blocked.";
            scrapeStatus.style.color = "crimson";
        }
    });
}

// global: generate filename based on pattern
function generateFilename(pattern, info, extension) {
    const protectedPattern = pattern.replace(/(\{.*?\})/g, "[$1]"); // avoid moment.js formatting placeholder
    let resultName = moment().format(protectedPattern);

    // process {quizSetTitle}
    let rawTitle = info.quizSetTitle || "quizlet_set";
    rawTitle = rawTitle.replace(/\s+/g, " "); // normalize whitespace
    rawTitle = rawTitle.replace(/ /g, "_"); // replace whitespace with underscore
    rawTitle = rawTitle.replace(/[^\p{L}\p{N}_-]/gu, "-"); // replace invalid characters
    resultName = resultName.replace("{quizSetTitle}", rawTitle);

    // process {swapState}
    const swapStr = info.swapped ? "DT" : "TD"; // DT: Definition - Term, TD: Term - Definition
    resultName = resultName.replace("{swapState}", swapStr);

    // process file extension
    if (!resultName.toLowerCase().endsWith("." + extension)) {
        resultName += "." + extension;
    }

    return resultName;
}

// global: data formatting
async function formatData(payload, format, includeImg) {
    if (format === "csv") {
        return {
            content: generateCSV(payload, includeImg),
            mimeType: "text/csv;charset=utf-8",
        };
    } else if (format === "docx") {
        const blob = await generateDOCX(payload, includeImg);
        return {
            content: blob,
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        };
    } else {
        return {
            content: generateJSON(payload, includeImg),
            mimeType: "application/json;charset=utf-8",
        };
    }
}

// global: download file utility
async function downloadFile(content, filename, mimeType) {
    let finalContent = content;

    if (content instanceof Blob) {
        finalContent = content;
    } else if (mimeType && mimeType.includes("csv")) {
        if (typeof content === "string" && !content.startsWith("\uFEFF")) {
            finalContent = "\uFEFF" + content; // add BOM for CSV files (UTF-8)
        }
    } else if (typeof content === "object") {
        finalContent = JSON.stringify(content, null, 4); // pretty-print JSON if needed
    }

    const blob = finalContent instanceof Blob ? finalContent : new Blob([finalContent], { type: mimeType });
    const url = URL.createObjectURL(blob);

    try {
        await chrome.downloads.download({
            url: url,
            filename: filename,
            saveAs: true,
        });
    } catch (e) {
        console.error("Download failed:", e);
    }
}

// global: combine text and image URL (JSON/CSV)
function getDefinitionStr(defObj, includeImg) {
    let str = defObj.text || "";
    if (includeImg && defObj.image) {
        if (str.length > 0) str += " - ";
        str += `(${defObj.image})`;
    }
    return str;
}

// global: generate JSON content from payload
function generateJSON(payload, includeImg) {
    const cleanString = (str) => {
        if (typeof str !== "string") return str;
        return str.replace(/[ \t\r\f\v]+/g, " ").trim();
    };

    const newPayload = JSON.parse(JSON.stringify(payload));

    for (const key in newPayload.info) {
        newPayload.info[key] = cleanString(newPayload.info[key]);
    }

    for (const key in newPayload.quizData) {
        newPayload.quizData[key].termPart = cleanString(newPayload.quizData[key].termPart);
        const defObj = newPayload.quizData[key].definitionPart;
        defObj.text = cleanString(defObj.text);
        newPayload.quizData[key].definitionPart = getDefinitionStr(defObj, includeImg);
    }

    return JSON.stringify(newPayload, null, 4);
}

// global: generate CSV content from payload
function generateCSV(payload, includeImg) {
    const info = payload.info;
    const data = payload.quizData;
    const rows = [];

    // wrap csv cell content properly
    const escape = (text) => {
        if (!text) return "";
        const str = String(text);
        if (str.includes(",") || str.includes("\n") || str.includes('"')) {
            return `"${str.replace(/"/g, '""')}"`;
        }
        return str;
    };

    // section info
    rows.push(`Key,Value`);
    rows.push(`Quiz Set Title,${escape(info.quizSetTitle)}`);
    rows.push(`Quiz Set URL,${escape(info.quizSetURL)}`);
    rows.push(`Creator Name,${escape(info.creatorName)}`);
    rows.push(`Creator URL,${escape(info.creatorURL)}`);
    rows.push(`Date Scraped,${escape(info.dateScraped)}`);
    rows.push(`Number of Quizzes,${info.numberOfQuizzes}`);
    rows.push(`,`); // empty separator row

    const headerLeft = info.swapped ? "Definition Part" : "Term Part";
    const headerRight = info.swapped ? "Term Part" : "Definition Part";
    rows.push(`${headerLeft},${headerRight}`);

    // section data
    const sortedKeys = Object.keys(data).sort();
    sortedKeys.forEach((key) => {
        const item = data[key];
        const defStr = getDefinitionStr(item.definitionPart, includeImg);
        const termStr = item.termPart;

        const col1Text = info.swapped ? defStr : termStr;
        const col2Text = info.swapped ? termStr : defStr;

        rows.push(`${escape(col1Text)},${escape(col2Text)}`);
    });

    return rows.join("\n");
}

// global: generate DOCX content from payload
async function generateDOCX(payload, includeImg) {
    // check if docx library is loaded
    if (typeof docx === "undefined") {
        throw new Error("DOCX library not loaded");
    }

    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ExternalHyperlink, AlignmentType, ImageRun, BorderStyle } = docx;
    const info = payload.info;
    const data = payload.quizData;
    const FONT_NAME = "Calibri";

    // section info
    const createInfoLine = (key, value) => {
        return new Paragraph({
            children: [new TextRun({ text: key, bold: true, size: 24, font: FONT_NAME }), new TextRun({ text: String(value), size: 24, font: FONT_NAME })],
            spacing: { after: 100 },
        });
    };

    const createHyperlinkLine = (key, value) => {
        return new Paragraph({
            children: [
                new TextRun({ text: key, bold: true, size: 24, font: FONT_NAME }),
                new ExternalHyperlink({
                    children: [new TextRun({ text: value, style: "Hyperlink", size: 24, font: FONT_NAME, color: "0563C1", underline: { type: "single" } })],
                    link: value,
                }),
            ],
            spacing: { after: 100 },
        });
    };

    const infoSection = [
        createInfoLine("Quiz Set Title: ", info.quizSetTitle),
        createHyperlinkLine("Quiz Set URL: ", info.quizSetURL),
        createInfoLine("Creator Name: ", info.creatorName),
        createHyperlinkLine("Creator URL: ", info.creatorURL),
        createInfoLine("Date Scraped: ", info.dateScraped),
        createInfoLine("Number of Quizzes: ", info.numberOfQuizzes),
        new Paragraph({ text: "", spacing: { after: 200 } }),
    ];

    // section data
    // header row
    const headerLeft = info.swapped ? "Definition Part" : "Term Part";
    const headerRight = info.swapped ? "Term Part" : "Definition Part";

    const tableRows = [
        new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: headerLeft, bold: true, size: 24, font: FONT_NAME })], alignment: AlignmentType.CENTER })],
                    width: { size: 30, type: WidthType.PERCENTAGE },
                }),
                new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: headerRight, bold: true, size: 24, font: FONT_NAME })], alignment: AlignmentType.CENTER })],
                    width: { size: 70, type: WidthType.PERCENTAGE },
                }),
            ],
        }),
    ];

    // data rows
    const createTextRunsWithNewlines = (text, size, bold = false) => {
        if (!text) return [new TextRun({ text: "", size: size, font: FONT_NAME })];

        const lines = text.split("\n");
        return lines.map((line, index) => {
            return new TextRun({ text: line, break: index > 0 ? 1 : 0, size: size, bold: bold, font: FONT_NAME });
        });
    };

    const fetchImage = async (url) => {
        try {
            const response = await fetch(url);
            if (!response.ok) return null;
            return await response.arrayBuffer();
        } catch (error) {
            console.warn("Failed to fetch image:", url, error);
            return null;
        }
    };

    const sortedKeys = Object.keys(data).sort();
    for (const key of sortedKeys) {
        const item = data[key];

        // content term text
        const termCellChildren = [new Paragraph({ children: createTextRunsWithNewlines(item.termPart, 22) })];

        // content definition text
        const defCellChildren = [new Paragraph({ children: createTextRunsWithNewlines(item.definitionPart.text, 22) })];

        // content definition image
        if (includeImg && item.definitionPart.image) {
            const imageBuffer = await fetchImage(item.definitionPart.image);
            if (imageBuffer) {
                defCellChildren.push(new Paragraph({ text: "", spacing: { before: 100 } }));

                const imgRun = new ImageRun({ data: imageBuffer, transformation: { width: 200, height: 150 } });
                const linkedImage = new ExternalHyperlink({ children: [imgRun], link: item.definitionPart.image });
                const imageTable = new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph({ children: [linkedImage], alignment: AlignmentType.CENTER })],
                                    margins: { top: 0, bottom: 0, left: 0, right: 0 },
                                    borders: {
                                        top: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
                                        bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
                                        left: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
                                        right: { style: BorderStyle.SINGLE, size: 6, color: "000000" },
                                    },
                                }),
                            ],
                        }),
                    ],
                    width: { size: 100, type: WidthType.PERCENTAGE },
                });

                defCellChildren.push(imageTable);
            }
        }

        // handle swap position
        const cell1Children = info.swapped ? defCellChildren : termCellChildren;
        const cell2Children = info.swapped ? termCellChildren : defCellChildren;

        tableRows.push(new TableRow({ children: [new TableCell({ children: cell1Children }), new TableCell({ children: cell2Children })] }));
    }

    const table = new Table({
        rows: tableRows,
        width: { size: 100, type: WidthType.PERCENTAGE },
    });

    // create document object
    const doc = new Document({
        sections: [
            {
                properties: {},
                children: [...infoSection, table],
            },
        ],
    });

    return await Packer.toBlob(doc);
}
