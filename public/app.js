"use strict";


/* =====================================================
   STATE
   ===================================================== */

let currentTab = "letter";

let currentLetter = "";

let currentNoteFile = "";

let currentKeyPoints = [];

let originalSourceText = "";


/* =====================================================
   ELEMENT HELPER
   ===================================================== */

function $(id) {

  return document.getElementById(id);

}


/* =====================================================
   API REQUEST
   ===================================================== */

async function apiRequest(
  url,
  options = {}
) {

  let response;

  try {

    response =
      await fetch(
        url,
        options
      );

  } catch (error) {

    throw new Error(
      "Unable to connect to the server."
    );

  }


  let data = null;


  try {

    data =
      await response.json();

  } catch (error) {

    throw new Error(
      "Server returned an invalid response."
    );

  }


  if (
    !response.ok ||
    data.success === false
  ) {

    throw new Error(
      data.message ||
      data.error ||
      `Request failed (${response.status}).`
    );

  }


  return data;

}


/* =====================================================
   LOADING
   ===================================================== */

function showLoading(message) {

  $("loadingText").textContent =
    message ||
    "Processing...";


  $("loading")
    .classList
    .add("show");

}


function hideLoading() {

  $("loading")
    .classList
    .remove("show");

}


/* =====================================================
   TOAST
   ===================================================== */

function toast(message) {

  const element =
    $("toast");


  element.textContent =
    message;


  element.style.display =
    "block";


  window.setTimeout(
    () => {

      element.style.display =
        "none";

    },
    3500
  );

}


/* =====================================================
   HEALTH CHECK
   ===================================================== */

async function checkHealth() {

  try {

    const data =
      await apiRequest(
        "/api/health"
      );


    $("statusDot")
      .style.background =
      "#27ae60";


    $("statusText")
      .textContent =
      "Online";


    $("systemInfo").innerHTML = "";


    addSystemLine(
      "Server",
      "Online"
    );


    addSystemLine(
      "Gemini PDF/OCR",
      data.geminiConfigured
        ? "Configured"
        : "Not configured"
    );


    addSystemLine(
      "Gemini Model",
      data.geminiModel ||
      "-"
    );


    addSystemLine(
      "OpenAI Drafting",
      data.openaiConfigured
        ? "Configured"
        : "Not configured"
    );


    addSystemLine(
      "OpenAI Model",
      data.openaiModel ||
      "-"
    );


    addSystemLine(
      "Architecture",
      "PDF → Gemini → OpenAI → Letter + Note File"
    );


  } catch (error) {

    $("statusDot")
      .style.background =
      "#d92d20";


    $("statusText")
      .textContent =
      "Offline";


    $("systemInfo").textContent =
      error.message;

  }

}


/* =====================================================
   SYSTEM INFORMATION
   ===================================================== */

function addSystemLine(
  label,
  value
) {

  const div =
    document.createElement("div");


  const strong =
    document.createElement("strong");


  strong.textContent =
    `${label}: `;


  div.appendChild(
    strong
  );


  div.appendChild(
    document.createTextNode(
      value
    )
  );


  $("systemInfo")
    .appendChild(div);

}


/* =====================================================
   PDF EXTRACTION
   ===================================================== */

async function extractPDF() {

  const input =
    $("pdfFile");


  const file =
    input.files[0];


  if (!file) {

    toast(
      "Please select a PDF file."
    );

    return;

  }


  const isPdf =
    file.type === "application/pdf" ||
    file.name
      .toLowerCase()
      .endsWith(".pdf");


  if (!isPdf) {

    toast(
      "Please select a PDF file."
    );

    return;

  }


  const formData =
    new FormData();


  formData.append(
    "pdf",
    file
  );


  try {

    showLoading(
      "Gemini is reading the PDF / OCR..."
    );


    $("pdfStatus").textContent =
      "Uploading PDF to Gemini...";


    const data =
      await apiRequest(
        "/api/pdf-to-text",
        {
          method:
            "POST",

          body:
            formData
        }
      );


    const extractedText =
      data.text ||
      data.extractedText ||
      "";


    if (!extractedText.trim()) {

      throw new Error(
        "No readable text was extracted from the PDF."
      );

    }


    $("sourceText").value =
      extractedText;


    originalSourceText =
      extractedText;


    const count =
      data.characterCount ||
      extractedText.length;


    $("pdfStatus").textContent =
      `Extraction completed. ${count.toLocaleString()} characters extracted.`;


    toast(
      "PDF text extracted successfully."
    );


  } catch (error) {

    $("pdfStatus").textContent =
      "";


    toast(
      error.message
    );


  } finally {

    hideLoading();

  }

}


/* =====================================================
   GENERATE DRAFT
   ===================================================== */

async function generateDraft() {

  const sourceText =
    $("sourceText")
      .value
      .trim();


  const workCommand =
    $("workCommand")
      .value
      .trim();


  const language =
    $("language")
      .value;


  const recipient =
    $("recipient")
      .value
      .trim();


  const subject =
    $("subject")
      .value
      .trim();


  if (!sourceText) {

    toast(
      "Please enter Source / Received Text or upload a PDF."
    );

    return;

  }


  if (!workCommand) {

    toast(
      "Please enter Work / Command."
    );

    return;

  }


  originalSourceText =
    sourceText;


  try {

    showLoading(
      "OpenAI is preparing the official Letter and Note File..."
    );


    const data =
      await apiRequest(
        "/api/generate",
        {

          method:
            "POST",

          headers: {
            "Content-Type":
              "application/json"
          },

          body:
            JSON.stringify({

              sourceText:
                sourceText,

              workCommand:
                workCommand,

              language:
                language,

              recipient:
                recipient,

              subject:
                subject

            })

        }
      );


    currentLetter =
      data.letter ||
      data.currentLetter ||
      "";


    currentNoteFile =
      data.note_file ||
      data.noteFile ||
      "";


    currentKeyPoints =
      Array.isArray(
        data.key_points
      )
        ? data.key_points
        : [];


    renderKeyPoints();


    showTab(
      "letter"
    );


    toast(
      "Letter and Note File generated successfully."
    );


  } catch (error) {

    toast(
      error.message
    );

  } finally {

    hideLoading();

  }

}


/* =====================================================
   CONTINUE / ALTER
   ===================================================== */

async function continueDraft() {

  const alterCommand =
    $("alterCommand")
      .value
      .trim();


  if (!originalSourceText) {

    toast(
      "Please generate the first draft before using Continue / Alter."
    );

    return;

  }


  if (
    !currentLetter ||
    !currentNoteFile
  ) {

    toast(
      "Previous Letter / Note File is not available."
    );

    return;

  }


  if (!alterCommand) {

    toast(
      "Please enter the Continue / Alter Command."
    );

    return;

  }


  try {

    showLoading(
      "OpenAI is revising the Letter and Note File..."
    );


    const data =
      await apiRequest(
        "/api/continue",
        {

          method:
            "POST",

          headers: {
            "Content-Type":
              "application/json"
          },

          body:
            JSON.stringify({

              sourceText:
                originalSourceText,

              previousLetter:
                currentLetter,

              previousNoteFile:
                currentNoteFile,

              alterCommand:
                alterCommand,

              language:
                $("language").value,

              recipient:
                $("recipient")
                  .value
                  .trim(),

              subject:
                $("subject")
                  .value
                  .trim()

            })

        }
      );


    currentLetter =
      data.letter ||
      data.currentLetter ||
      "";


    currentNoteFile =
      data.note_file ||
      data.noteFile ||
      "";


    currentKeyPoints =
      Array.isArray(
        data.key_points
      )
        ? data.key_points
        : [];


    renderKeyPoints();


    showTab(
      "letter"
    );


    $("alterCommand").value =
      "";


    toast(
      "Draft revised successfully."
    );


  } catch (error) {

    toast(
      error.message
    );

  } finally {

    hideLoading();

  }

}


/* =====================================================
   TABS
   ===================================================== */

function showTab(tab) {

  currentTab =
    tab;


  $("letterTab")
    .classList
    .remove("active");


  $("noteTab")
    .classList
    .remove("active");


  if (
    tab === "letter"
  ) {

    $("letterTab")
      .classList
      .add("active");


    $("output").textContent =
      currentLetter ||
      "Letter will appear here.";


  } else {

    $("noteTab")
      .classList
      .add("active");


    $("output").textContent =
      currentNoteFile ||
      "Note File will appear here.";

  }

}


/* =====================================================
   KEY POINTS
   ===================================================== */

function renderKeyPoints() {

  const box =
    $("keyPoints");


  box.replaceChildren();


  if (
    !currentKeyPoints.length
  ) {

    box.classList
      .add("hidden");

    return;

  }


  box.classList
    .remove("hidden");


  const strong =
    document.createElement(
      "strong"
    );


  strong.textContent =
    "Key Points";


  box.appendChild(
    strong
  );


  const list =
    document.createElement(
      "ul"
    );


  currentKeyPoints.forEach(
    point => {

      const li =
        document.createElement(
          "li"
        );


      li.textContent =
        String(point);


      list.appendChild(
        li
      );

    }
  );


  box.appendChild(
    list
  );

}


/* =====================================================
   COPY
   ===================================================== */

async function copyOutput() {

  const text =
    currentTab === "letter"
      ? currentLetter
      : currentNoteFile;


  if (!text) {

    toast(
      "There is no output to copy."
    );

    return;

  }


  try {

    await navigator
      .clipboard
      .writeText(text);


    toast(
      "Copied to clipboard."
    );


  } catch (error) {

    toast(
      "Copy failed. Please copy the text manually."
    );

  }

}


/* =====================================================
   WORD DOWNLOAD
   ===================================================== */

async function downloadWord() {

  if (
    !currentLetter ||
    !currentNoteFile
  ) {

    toast(
      "Generate the Letter and Note File first."
    );

    return;

  }


  try {

    showLoading(
      "Creating Word document..."
    );


    const response =
      await fetch(
        "/api/download-word",
        {

          method:
            "POST",

          headers: {
            "Content-Type":
              "application/json"
          },

          body:
            JSON.stringify({

              letter:
                currentLetter,

              noteFile:
                currentNoteFile

            })

        }
      );


    if (!response.ok) {

      let message =
        "Word download failed.";


      try {

        const data =
          await response.json();


        message =
          data.message ||
          message;

      } catch (error) {

        // Ignore JSON parsing failure.

      }


      throw new Error(
        message
      );

    }


    const blob =
      await response.blob();


    const url =
      window.URL
        .createObjectURL(blob);


    const link =
      document.createElement("a");


    link.href =
      url;


    link.download =
      "Revenue_Draft.docx";


    document.body
      .appendChild(link);


    link.click();


    link.remove();


    window.URL
      .revokeObjectURL(url);


    toast(
      "Word document created."
    );


  } catch (error) {

    toast(
      error.message
    );


  } finally {

    hideLoading();

  }

}


/* =====================================================
   PRINT
   ===================================================== */

function printOutput() {

  const text =
    currentTab === "letter"
      ? currentLetter
      : currentNoteFile;


  if (!text) {

    toast(
      "There is no output to print."
    );

    return;

  }


  const printWindow =
    window.open(
      "",
      "_blank"
    );


  if (!printWindow) {

    toast(
      "Popup was blocked by the browser."
    );

    return;

  }


  const safeText =
    escapeHtml(text);


  printWindow.document.open();


  printWindow.document.write(`
<!DOCTYPE html>

<html>

<head>

  <meta charset="UTF-8">

  <title>
    Revenue Office Draft
  </title>

  <style>

    body {
      font-family:
        Georgia,
        "Times New Roman",
        serif;

      font-size:
        15px;

      line-height:
        1.7;

      padding:
        30px;

      white-space:
        pre-wrap;
    }

  </style>

</head>

<body>

  <div id="printContent">${safeText}</div>

</body>

</html>
  `);


  printWindow.document.close();


  window.setTimeout(
    () => {

      printWindow.focus();

      printWindow.print();

    },
    300
  );

}


/* =====================================================
   CLEAR
   ===================================================== */

function clearAll() {

  $("sourceText").value =
    "";


  $("workCommand").value =
    "";


  $("alterCommand").value =
    "";


  $("recipient").value =
    "";


  $("subject").value =
    "";


  $("pdfFile").value =
    "";


  $("pdfStatus").textContent =
    "";


  currentLetter =
    "";


  currentNoteFile =
    "";


  currentKeyPoints =
    [];


  originalSourceText =
    "";


  renderKeyPoints();


  showTab(
    "letter"
  );


  toast(
    "Form cleared."
  );

}


/* =====================================================
   ESCAPE HTML
   ===================================================== */

function escapeHtml(value) {

  return String(value)

    .replace(
      /&/g,
      "&amp;"
    )

    .replace(
      /</g,
      "&lt;"
    )

    .replace(
      />/g,
      "&gt;"
    )

    .replace(
      /"/g,
      "&quot;"
    )

    .replace(
      /'/g,
      "&#039;"
    );

}


/* =====================================================
   EVENT LISTENERS
   ===================================================== */

function initializeApp() {

  const buttons = {

    extractPdfBtn:
      extractPDF,

    generateBtn:
      generateDraft,

    clearBtn:
      clearAll,

    continueBtn:
      continueDraft,

    copyBtn:
      copyOutput,

    downloadWordBtn:
      downloadWord,

    printBtn:
      printOutput

  };


  Object.entries(
    buttons
  ).forEach(
    ([id, handler]) => {

      const element =
        $(id);


      if (element) {

        element.addEventListener(
          "click",
          handler
        );

      }

    }
  );


  if ($("letterTab")) {

    $("letterTab")
      .addEventListener(
        "click",
        () => showTab("letter")
      );

  }


  if ($("noteTab")) {

    $("noteTab")
      .addEventListener(
        "click",
        () => showTab("note")
      );

  }


  checkHealth();

}


/* =====================================================
   START
   ===================================================== */

if (
  document.readyState ===
  "loading"
) {

  document.addEventListener(
    "DOMContentLoaded",
    initializeApp
  );

} else {

  initializeApp();

}
