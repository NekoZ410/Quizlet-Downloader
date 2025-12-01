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

// global: process scraping and download result
const statusDisplay = document.getElementById("qd-status");
const scrapeBtn = document.getElementById("qd-scrapeBtn");
const swapCheckbox = document.getElementById("qd-outFileContentOrder");
const filenameInput = document.getElementById("qd-outFileNamePattern");
const formatSelect = document.getElementById("qd-outFileNameExt");

if (scrapeBtn) {
    scrapeBtn.addEventListener("click", async () => {
        let [tab] = await chrome.tabs.query({ active: true, currentWindow: true }); // get current active tab

        // validate Quizlet page URL
        if (!tab.url.includes("quizlet.com")) {
            statusDisplay.textContent = "Error: Please open a Quizlet set page.";
            statusDisplay.style.color = "crimson";
            return;
        }

        statusDisplay.textContent = "Scraping data...";
        statusDisplay.style.color = "yellow";

        const isSwapped = swapCheckbox ? swapCheckbox.checked : false; // check swap option status

        try {
            const response = await chrome.tabs.sendMessage(tab.id, {
                action: "scrape_quizlet",
                swap: isSwapped,
            });

            if (response && response.status === "success") {
                const payload = response.payload;
                const outputFormat = formatSelect ? formatSelect.value : "json";

                const filenamePattern = filenameInput ? filenameInput.value : "{quizSetTitle}_YYYY-MM-DD_HH-mm-ss_{swapState}";
                const finalFilename = generateFilename(filenamePattern, payload.info, outputFormat);

                const { content, mimeType } = await formatData(payload, outputFormat);
                await downloadFile(content, finalFilename, mimeType);

                statusDisplay.textContent = `Done! Found ${payload.info.numberOfQuizzes} quizzes.`;
                statusDisplay.style.color = "forestgreen";
            } else {
                statusDisplay.textContent = "Error: No data found or script failed.";
                statusDisplay.style.color = "crimson";
            }
        } catch (error) {
            console.error(error);
            statusDisplay.textContent = "Error: Page not ready or blocked.";
            statusDisplay.style.color = "crimson";
        }
    });
}

// global: generate filename based on pattern
function generateFilename(pattern, info, extension) {
    const protectedPattern = pattern.replace(/(\{.*?\})/g, "[$1]"); // avoid moment.js formatting placeholder
    let resultName = moment().format(protectedPattern);

    // process {quizSetTitle}
    const safeQuizSetTitle = (info.quizSetTitle || "quizlet_set").replace(/\s+/g, "_");
    resultName = resultName.replace("{quizSetTitle}", safeQuizSetTitle);

    // process {swapState}
    const swapStr = info.swapped ? "BS" : "SB"; // BS: big first small later; SB: small first big later (default)
    resultName = resultName.replace("{swapState}", swapStr);

    // process file extension
    if (!resultName.toLowerCase().endsWith("." + extension)) {
        resultName += "." + extension;
    }

    return resultName;
}

// global: data formatting
async function formatData(payload, format) {
    if (format === "csv") {
        return {
            content: generateCSV(payload),
            mimeType: "text/csv;charset=utf-8",
        };
    } else if (format === "docx") {
        const blob = await generateDOCX(payload);
        return {
            content: blob,
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        };
    } else {
        return {
            content: generateJSON(payload),
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

// global: generate JSON content from payload
function generateJSON(payload) {
    const cleanString = (str) => {
        if (typeof str !== "string") return str;
        return str.replace(/[ \t\r\f\v]+/g, " ").trim();
    };

    const newPayload = JSON.parse(JSON.stringify(payload));

    for (const key in newPayload.info) {
        newPayload.info[key] = cleanString(newPayload.info[key]);
    }

    for (const key in newPayload.quizData) {
        newPayload.quizData[key].partSmall = cleanString(newPayload.quizData[key].partSmall);
        newPayload.quizData[key].partBig = cleanString(newPayload.quizData[key].partBig);
    }

    return JSON.stringify(newPayload, null, 4);
}

// global: generate CSV content from payload
function generateCSV(payload) {
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

    const headerLeft = info.swapped ? "Big Part" : "Small Part";
    const headerRight = info.swapped ? "Small Part" : "Big Part";
    rows.push(`${headerLeft},${headerRight}`);

    // section data
    const sortedKeys = Object.keys(data).sort();
    sortedKeys.forEach((key) => {
        const item = data[key];
        if (info.swapped) {
            rows.push(`${escape(item.partBig)},${escape(item.partSmall)}`); // big first, small later
        } else {
            rows.push(`${escape(item.partSmall)},${escape(item.partBig)}`); // small first, big later
        }
    });

    return rows.join("\n");
}

// global: generate DOCX content from payload
async function generateDOCX(payload) {
    // check if docx library is loaded
    if (typeof docx === "undefined") {
        throw new Error("DOCX library not loaded");
    }

    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, ExternalHyperlink, AlignmentType } = docx;
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
                    children: [
                        new TextRun({
                            text: value,
                            style: "Hyperlink",
                            size: 24,
                            font: FONT_NAME,
                            color: "0563C1",
                            underline: { type: "single" },
                        }),
                    ],
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
    const headerLeft = info.swapped ? "Big Part" : "Small Part";
    const headerRight = info.swapped ? "Small Part" : "Big Part";

    const tableRows = [
        new TableRow({
            children: [
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [new TextRun({ text: headerLeft, bold: true, size: 24, font: FONT_NAME })],
                            alignment: AlignmentType.CENTER,
                        }),
                    ],
                    width: { size: 30, type: WidthType.PERCENTAGE },
                }),
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [new TextRun({ text: headerRight, bold: true, size: 24, font: FONT_NAME })],
                            alignment: AlignmentType.CENTER,
                        }),
                    ],
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
            return new TextRun({
                text: line,
                break: index > 0 ? 1 : 0,
                size: size,
                bold: bold,
                font: FONT_NAME,
            });
        });
    };

    const sortedKeys = Object.keys(data).sort();
    sortedKeys.forEach((key) => {
        const item = data[key];
        const col1Text = info.swapped ? item.partBig : item.partSmall;
        const col2Text = info.swapped ? item.partSmall : item.partBig;

        tableRows.push(
            new TableRow({
                children: [
                    new TableCell({
                        children: [
                            new Paragraph({
                                children: createTextRunsWithNewlines(col1Text, 22),
                            }),
                        ],
                    }),
                    new TableCell({
                        children: [
                            new Paragraph({
                                children: createTextRunsWithNewlines(col2Text, 22),
                            }),
                        ],
                    }),
                ],
            })
        );
    });

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
