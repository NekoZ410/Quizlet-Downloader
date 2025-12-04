// global: html selector
const SELECTORS = {
    QUIZ_SET_TITLE: ".s1ygu81a",
    QUIZ_SET_COUNTS: ".t1hzdbu9 .t18rmeis",
    QUIZ_SET_SHOWMORE_BTN: ".a1qd8xfe.a1q8vbq2",
    QUIZ_SET_CREATOR_NAME: ".u1xtrgf5 .UserLink-content .UILink span",
    QUIZ_SET_CREATOR_URL: ".u1xtrgf5 .UserLink-content .UILink",

    QUIZZES_SELECTOR: ".SetPageTermsList-term .se6rv9p",
    QUIZ_PART_TERM_TEXT: ".s7ascy3",
    QUIZ_PART_DEFINITION: ".l1rpwius",
    QUIZ_PART_DEFINITION_TEXT: ".hdftvph .TermText",
    QUIZ_PART_DEFINITION_IMAGE: ".sumuxuf .SetPageTerm-image",
};

// global: listen to message to scrape Quizlet data
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "scrape_quizlet") {
        const titleEl = document.querySelector(SELECTORS.QUIZ_SET_TITLE); // get quiz set title element
        const rawTitle = titleEl ? titleEl.innerText.trim() : "Untitled Set";

        const creatorNameEl = document.querySelector(SELECTORS.QUIZ_SET_CREATOR_NAME); // get creator name element
        const creatorName = creatorNameEl ? creatorNameEl.innerText.trim() : "Unknown";
        const creatorUrlEl = document.querySelector(SELECTORS.QUIZ_SET_CREATOR_URL); // get creator URL element
        const creatorURL = creatorUrlEl ? creatorUrlEl.href : "";

        const termRows = document.querySelectorAll(SELECTORS.QUIZZES_SELECTOR); // get nodelist of all quizzes
        const count = termRows.length; // total number of quizzes counted
        const countEl = document.querySelector(SELECTORS.QUIZ_SET_COUNTS); // get total number of quizzes element
        if (countEl) {
            const countMatch = countEl.innerText.match(/\d+/);
            if (countMatch) {
                const expectedCount = parseInt(countMatch[0], 10); // check expected count
                if (count !== expectedCount) {
                    sendResponse({
                        status: "fail",
                        message: `Mismatch error: Found ${count} terms but expected ${expectedCount}. Please scroll down/expand all.`,
                    });
                    return true;
                }
            }
        }
        const padLen = String(count).length; // length for zero-padding indices

        const now = new Date();
        const offsetMs = now.getTimezoneOffset() * 60000;
        const localDate = new Date(now.getTime() - offsetMs).toISOString().slice(0, -1) + "Z"; // get local ISO date string

        const imgRegex = /https:\/\/o\..+/; // regex to filter image URLs

        let quizDataObj = {};
        if (count > 0) {
            termRows.forEach((row, index) => {
                const elTerm = row.querySelector(SELECTORS.QUIZ_PART_TERM_TEXT);
                const textTerm = elTerm ? elTerm.innerText.trim().replace(/\n+/g, "\n") : "";

                const elDefContainer = row.querySelector(SELECTORS.QUIZ_PART_DEFINITION);
                let textDef = "";
                let imgSrcDef = "";
                if (elDefContainer) {
                    const elDefText = elDefContainer.querySelector(SELECTORS.QUIZ_PART_DEFINITION_TEXT);
                    const elDefImg = elDefContainer.querySelector(SELECTORS.QUIZ_PART_DEFINITION_IMAGE);

                    textDef = elDefText ? elDefText.innerText.trim().replace(/\n+/g, "\n") : "";
                    if (elDefImg && elDefImg.src) {
                        const rawSrc = elDefImg.src;
                        const match = rawSrc.match(imgRegex);
                        imgSrcDef = match ? match[0] : rawSrc;
                    }
                }

                const idxKey = String(index + 1).padStart(padLen, "0"); // zero-padded index key

                let entry = {};

                const defObj = {
                    text: textDef,
                    image: imgSrcDef,
                };

                if (request.swap) {
                    // term first, definition later
                    entry["definitionPart"] = defObj;
                    entry["termPart"] = textTerm;
                } else {
                    // definition first, term later (default)
                    entry["termPart"] = textTerm;
                    entry["definitionPart"] = defObj;
                }
                quizDataObj[idxKey] = entry;
            });
        }

        const finalPayload = {
            info: {
                quizSetTitle: rawTitle,
                quizSetURL: window.location.href,
                creatorName: creatorName,
                creatorURL: creatorURL,
                dateScraped: localDate,
                numberOfQuizzes: count,
                swapped: request.swap || false,
            },
            quizData: quizDataObj,
        };

        if (count > 0) {
            sendResponse({ status: "success", payload: finalPayload });
        } else {
            sendResponse({ status: "fail", count: 0 });
        }
    }

    return true; // keep message channel open for async response
});

// global: auto expand quiz set
function autoExpandQuizSet() {
    try {
        const countEl = document.querySelector(SELECTORS.QUIZ_SET_COUNTS); // get total count element
        if (!countEl) return;

        // extract total count and check
        const countText = countEl.innerText;
        const totalMatch = countText.match(/\d+/);
        if (!totalMatch) return;

        const totalTerms = parseInt(totalMatch[0], 10);
        if (totalTerms <= 100) return;

        const remainingTerms = totalTerms - 100; // remaining quizzes
        const buttons = document.querySelectorAll(SELECTORS.QUIZ_SET_SHOWMORE_BTN);
        for (const btn of buttons) {
            if (btn.innerText.includes(remainingTerms.toString())) {
                console.log(`[Quizlet Downloader] Found 'Show more' button for ${remainingTerms} terms. Clicking...`);
                btn.click();
                break;
            }
        }
    } catch (e) {
        console.error("[Quizlet Downloader] Auto-expand failed:", e);
    }
}

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", autoExpandQuizSet);
} else {
    autoExpandQuizSet();
}
setTimeout(autoExpandQuizSet, 2000);
