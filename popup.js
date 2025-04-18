// popup.js
// const scriptURL = "https://script.google.com/macros/s/AKfycbxJXoFlX-ddGzXPzr4ixDstSWPoy1xlRLXgZyhsH75a0T2RldPCCZs97hQRHp5mnBkB/exec";
const scriptURL = "https://script.google.com/macros/s/AKfycbzY__2Tly8XI3mqYOpHVP31u0a54Yucd5OHxRanhucnHV7f1MbbYHHfWlWYVi-81-NC/exec";
// let currentSpreadsheetId = null;

document.addEventListener("DOMContentLoaded", async () => {
    currentSpreadsheetId = await getActiveSheetId();

    const callAppsScript = async (func, data = {}) => {
        if (!currentSpreadsheetId) {
          alert("Please open a Google Sheet first.");
          return;
        }
        const payload = { function: func, spreadsheetId: currentSpreadsheetId, ...data };
        try {
          const response = await fetch(scriptURL, {
            method: 'POST',
            mode: 'cors',
            credentials: 'omit', // Add this line
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
          });
          if (!response.ok) {
            console.error("Apps Script Error:", response);
            alert("Error communicating with Google Apps Script.");
            return;
          }
          const responseData = await response.text(); // Get response as text for debugging
          console.log("Apps Script response (text):", responseData);
          return responseData;
        } catch (error) {
          console.error("Fetch Error:", error);
          alert("Failed to connect to Google Apps Script.");
          return;
        }
      };
              
    async function getActiveSheetId() {
        let [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        if (!tab || !tab.url.includes("docs.google.com/spreadsheets/d/")) {
            alert("Please open a Google Sheet.");
            return null;
        }
        const urlParts = tab.url.split('/d/');
        if (urlParts.length < 2) return null; // Handle unexpected URL format
        const idPart = urlParts[1].split('/')[0];
        return idPart;
    }

    document.getElementById("renameAndUpdateImages").addEventListener("click", () => callAppsScript("renameAndUpdateImages"));

    document.getElementById("processSelectedDatasets").addEventListener("click", async () => {
        await callAppsScript("processSelectedDatasets");
        // No need to handle a response or display UI here
    });

    document.getElementById("renameFilesBasedOnSheet").addEventListener("click", () => callAppsScript("renameFilesBasedOnSheet"));
});
