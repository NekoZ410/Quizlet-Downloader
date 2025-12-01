// global: html selector
const SELECTORS = {
    QUIZ_SET_TITLE: ".s1ygu81a",
    QUIZ_SET_CREATOR_NAME: ".u1xtrgf5 .UserLink-content .UILink span",
    QUIZ_SET_CREATOR_URL: ".u1xtrgf5 .UserLink-content .UILink",
    QUIZZES_SELECTOR: ".SetPageTermsList-term .se6rv9p",
    QUIZ_PART_SMALL: ".s7ascy3",
    QUIZ_PART_BIG: ".hdftvph",
};

// global: listen to message to scrape Quizlet data
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "scrape_quizlet") {
        const titleEl = document.querySelector(SELECTORS.QUIZ_SET_TITLE); // get quiz set title element
        const rawTitle = titleEl ? titleEl.innerText.trim() : "Untitled Set";

        const creatorNameEl = document.querySelector(SELECTORS.QUIZ_SET_CREATOR_NAME);
        const creatorUrlEl = document.querySelector(SELECTORS.QUIZ_SET_CREATOR_URL);
        const creatorName = creatorNameEl ? creatorNameEl.innerText.trim() : "Unknown";
        const creatorURL = creatorUrlEl ? creatorUrlEl.href : "";

        const termRows = document.querySelectorAll(SELECTORS.QUIZZES_SELECTOR); // get nodelist of all quizzes
        const count = termRows.length; // total number of quizzes
        const padLen = String(count).length; // length for zero-padding indices

        const now = new Date();
        const offsetMs = now.getTimezoneOffset() * 60000;
        const localDate = new Date(now.getTime() - offsetMs).toISOString().slice(0, -1) + "Z"; // get local ISO date string

        let quizDataObj = {};
        if (count > 0) {
            termRows.forEach((row, index) => {
                const elSmall = row.querySelector(SELECTORS.QUIZ_PART_SMALL);
                const elBig = row.querySelector(SELECTORS.QUIZ_PART_BIG);

                const textSmall = elSmall ? elSmall.innerText.trim().replace(/\n+/g, "\n") : "";
                const textBig = elBig ? elBig.innerText.trim().replace(/\n+/g, "\n") : "";
                const idxKey = String(index + 1).padStart(padLen, "0"); // zero-padded index key

                let entry = {};
                if (request.swap) {
                    // big first, small later
                    entry["partBig"] = textBig;
                    entry["partSmall"] = textSmall;
                } else {
                    // small first, big later (default)
                    entry["partSmall"] = textSmall;
                    entry["partBig"] = textBig;
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
