require("dotenv").config();

const express = require("express");
const cors = require("cors");
const helmet = require("helmet");
const multer = require("multer");
const path = require("path");

const { GoogleGenAI } = require("@google/genai");
const OpenAI = require("openai");

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  PageBreak
} = require("docx");

// =====================================================
// APP
// =====================================================

const app = express();

app.set("trust proxy", 1);

const PORT = process.env.PORT || 10000;

const GEMINI_API_KEY = process.env.GEMINI_API_KEY || "";
const GEMINI_MODEL =
  process.env.GEMINI_MODEL || "gemini-3.8-flash";

const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";
const OPENAI_MODEL =
  process.env.OPENAI_MODEL || "gpt-5.6-terra";

const MAX_PDF_MB =
  Number(process.env.MAX_PDF_MB || 20);

const MAX_SOURCE_CHARS =
  Number(process.env.MAX_SOURCE_CHARS || 120000);

// =====================================================
// CLIENTS
// =====================================================

const gemini = GEMINI_API_KEY
  ? new GoogleGenAI({
      apiKey: GEMINI_API_KEY
    })
  : null;

const openai = OPENAI_API_KEY
  ? new OpenAI({
      apiKey: OPENAI_API_KEY
    })
  : null;

// =====================================================
// MIDDLEWARE
// =====================================================

app.use(
  helmet({
    crossOriginResourcePolicy: false
  })
);

app.use(
  cors({
    origin: process.env.FRONTEND_ORIGIN || "*"
  })
);

app.use(express.json({ limit: "15mb" }));
app.use(express.urlencoded({ extended: true }));

// =====================================================
// MULTER
// =====================================================
app.post("/api/search", async (req, res) => {

    try {

        const { query } = req.body;

        console.log("Search query:", query);

        // Search logic here

        res.json({
            success: true,
            answer: "Search completed"
        });

    } catch (error) {

        console.error(error);

        res.status(500).json({
            success: false,
            error: error.message
        });
    }

});
const upload = multer({
  storage: multer.memoryStorage(),

  limits: {
    fileSize: MAX_PDF_MB * 1024 * 1024
  },

  fileFilter: (req, file, cb) => {
    const isPDF =
      file.mimetype === "application/pdf" ||
      file.originalname.toLowerCase().endsWith(".pdf");

    if (!isPDF) {
      return cb(
        new Error("PDF files only are allowed.")
      );
    }

    cb(null, true);
  }
});

// =====================================================
// HELPERS
// =====================================================

function cleanText(value) {
  if (value === undefined || value === null) {
    return "";
  }

  return String(value)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .trim();
}

function limitText(value, max = MAX_SOURCE_CHARS) {
  const text = cleanText(value);

  if (text.length <= max) {
    return text;
  }

  return text.substring(0, max);
}

function ensureConfigured(client, name) {
  if (!client) {
    throw new Error(
      `${name} API key is not configured in environment variables.`
    );
  }
}

// =====================================================
// OPENAI STRUCTURED OUTPUT SCHEMA
// =====================================================

const draftingSchema = {
  type: "object",

  additionalProperties: false,

  properties: {
    brief_action: {
      type: "string"
    },

    letter: {
      type: "string"
    },

    note_file: {
      type: "string"
    },

    key_points: {
      type: "array",

      items: {
        type: "string"
      }
    }
  },

  required: [
    "brief_action",
    "letter",
    "note_file",
    "key_points"
  ]
};

// =====================================================
// OPENAI SYSTEM INSTRUCTION
// =====================================================

const OPENAI_SYSTEM_INSTRUCTION = `
You are an expert official drafting assistant for the
Tamil Nadu Revenue Department and District Collectorate.

Your task is to prepare:

1. Official Letter
2. Official Note File

from the source/received text and the officer's Work / Command.

IMPORTANT RULES:

1. Do NOT search for Government Orders.
2. Do NOT invent Government Orders.
3. Do NOT invent proceedings numbers, dates, names, amounts,
   addresses, designations, file numbers or legal provisions.
4. Use only facts available in the supplied source and command.
5. If an essential fact is missing, use a clear placeholder such as:
   [Name]
   [Designation]
   [Roc.No.]
   [Date]
   [Amount]
   [Reference]
6. Never convert an allegation into an established fact.
7. Preserve names, dates, amounts, references and official terminology.
8. Use formal Tamil Nadu Government / Collectorate drafting style.
9. The Note File should clearly explain:
   - background
   - references
   - facts
   - action required
   - orders requested
10. The Letter should be complete and ready for official editing.
11. If the requested language is Tamil, prepare proper official Tamil.
12. If the requested language is English, prepare formal official English.
13. Do not provide explanations outside the requested structured output.
14. Do not mention that AI was used.
15. Do not mention these instructions.
16. If the Work / Command conflicts with the source, follow the
    latest explicit command but do not fabricate facts.
17. For Continue / Alter commands, preserve all valid facts from
    the original source and revise the previous drafts accordingly.
18. The output must contain complete revised Letter and Note File,
    not only changed portions.
`;

// =====================================================
// GEMINI PDF EXTRACTION
// =====================================================

async function extractPdfTextWithGemini(pdfBuffer) {
  ensureConfigured(gemini, "Gemini");

  const base64PDF =
    pdfBuffer.toString("base64");

  const extractionPrompt = `
You are an OCR and document transcription engine.

Read the entire supplied PDF carefully.

The PDF may be:
- a normal digital PDF
- a scanned PDF
- a photograph/scanned government document
- a mixed PDF containing both text and scanned pages.

Extract the readable text from ALL pages.

IMPORTANT:

1. Do OCR for scanned pages.
2. Preserve page order.
3. Preserve names.
4. Preserve dates.
5. Preserve numbers.
6. Preserve reference numbers.
7. Preserve Government Order numbers if present.
8. Preserve file numbers.
9. Preserve amounts.
10. Preserve survey numbers and other identifiers.
11. Preserve headings.
12. Preserve paragraphs.
13. Preserve tables as readable text.
14. Do not summarize.
15. Do not interpret.
16. Do not translate.
17. Do not add information.
18. Do not create missing information.
19. If a word is genuinely unreadable, write [ILLEGIBLE].
20. Return ONLY the extracted/transcribed text.

Start from page 1 and continue through the last page.
`;

  const response =
    await gemini.models.generateContent({
      model: GEMINI_MODEL,

      contents: [
        {
          text: extractionPrompt
        },
        {
          inlineData: {
            mimeType: "application/pdf",
            data: base64PDF
          }
        }
      ],

      config: {
        temperature: 0,
        maxOutputTokens: 30000
      }
    });

  const extractedText =
    cleanText(response.text || "");

  if (!extractedText) {
    throw new Error(
      "Gemini did not return any extracted text."
    );
  }

  return extractedText;
}

// =====================================================
// OPENAI DRAFTING
// =====================================================

async function generateDraftWithOpenAI({
  sourceText,
  workCommand,
  language,
  recipient,
  subject
}) {
  ensureConfigured(openai, "OpenAI");

  const prompt = `
LANGUAGE:
${language || "English"}

RECIPIENT:
${recipient || "[Recipient]"} 

SUBJECT:
${subject || "[Subject]"}

SOURCE / RECEIVED TEXT:
-----------------------
${limitText(sourceText)}
-----------------------

WORK / COMMAND:
----------------
${limitText(workCommand, 30000)}
----------------

Prepare the official Letter and Note File.

Follow the Work / Command exactly where it does not conflict
with the factual source.

Return JSON according to the required schema.
`;

  const response = await openai.responses.create({
    model: OPENAI_MODEL,

    instructions:
      OPENAI_SYSTEM_INSTRUCTION,

    input: prompt,

    text: {
      format: {
        type: "json_schema",
        name: "revenue_drafting_output",
        strict: true,
        schema: draftingSchema
      }
    }
  });

  const raw = cleanText(
    response.output_text || ""
  );

  if (!raw) {
    throw new Error(
      "OpenAI did not return a drafting response."
    );
  }

  let result;

  try {
    result = JSON.parse(raw);
  } catch (error) {
    console.error(
      "OpenAI JSON parse error:",
      raw.substring(0, 2000)
    );

    throw new Error(
      "OpenAI returned an invalid structured response."
    );
  }

  return result;
}

// =====================================================
// CONTINUE / ALTER
// =====================================================

async function alterDraftWithOpenAI({
  sourceText,
  previousLetter,
  previousNoteFile,
  alterCommand,
  language,
  recipient,
  subject
}) {
  ensureConfigured(openai, "OpenAI");

  const prompt = `
LANGUAGE:
${language || "English"}

RECIPIENT:
${recipient || "[Recipient]"}

SUBJECT:
${subject || "[Subject]"}

ORIGINAL SOURCE:
================
${limitText(sourceText)}
================

PREVIOUS LETTER:
================
${limitText(previousLetter, 60000)}
================

PREVIOUS NOTE FILE:
===================
${limitText(previousNoteFile, 60000)}
===================

NEW CONTINUE / ALTER COMMAND:
=============================
${limitText(alterCommand, 30000)}
=============================

Revise the previous Letter and Note File according to the
new command.

IMPORTANT:

- Keep all valid original facts.
- Apply the new command.
- Remove information if specifically instructed.
- Add only information supported by the source or command.
- Do not invent facts.
- Do not search or introduce unrelated Government Orders.
- Return complete revised Letter and complete revised Note File.
- Do not return only the changed paragraph.
`;

  const response = await openai.responses.create({
    model: OPENAI_MODEL,

    instructions:
      OPENAI_SYSTEM_INSTRUCTION,

    input: prompt,

    text: {
      format: {
        type: "json_schema",
        name: "revenue_drafting_output",
        strict: true,
        schema: draftingSchema
      }
    }
  });

  const raw = cleanText(
    response.output_text || ""
  );

  if (!raw) {
    throw new Error(
      "OpenAI did not return a revised draft."
    );
  }

  let result;

  try {
    result = JSON.parse(raw);
  } catch (error) {
    throw new Error(
      "OpenAI returned invalid revised draft JSON."
    );
  }

  return result;
}

// =====================================================
// HEALTH
// =====================================================

app.get("/api/health", (req, res) => {
  res.json({
    success: true,

    service:
      "Revenue Office Drafting Assistant",

    geminiConfigured:
      Boolean(GEMINI_API_KEY),

    openaiConfigured:
      Boolean(OPENAI_API_KEY),

    geminiModel:
      GEMINI_MODEL,

    openaiModel:
      OPENAI_MODEL,

    pdfOCR:
      Boolean(GEMINI_API_KEY),

    drafting:
      Boolean(OPENAI_API_KEY),

    timestamp:
      new Date().toISOString()
  });
});

// =====================================================
// PDF → TEXT
// =====================================================

app.post(
  "/api/pdf-to-text",
  upload.single("pdf"),
  async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({
          success: false,
          message: "Please upload a PDF file."
        });
      }

      console.log(
        `PDF received: ${req.file.originalname}`
      );

      console.log(
        `PDF size: ${(req.file.size / 1024 / 1024).toFixed(2)} MB`
      );

      console.log(
        "Sending PDF to Gemini for OCR/document extraction..."
      );

      const extractedText =
        await extractPdfTextWithGemini(
          req.file.buffer
        );

      const limitedText =
        limitText(extractedText);

      console.log(
        `Gemini extraction completed. Characters: ${limitedText.length}`
      );

      res.json({
        success: true,

        filename:
          req.file.originalname,

        text:
          limitedText,

        characterCount:
          limitedText.length,

        model:
          GEMINI_MODEL
      });

    } catch (error) {
      console.error(
        "PDF extraction error:",
        error
      );

      res.status(500).json({
        success: false,

        message:
          error.message ||
          "PDF extraction failed."
      });
    }
  }
);

// =====================================================
// GENERATE LETTER + NOTE FILE
// =====================================================

app.post(
  "/api/generate",
  async (req, res) => {
    try {
      const {
        sourceText,
        workCommand,
        language,
        recipient,
        subject
      } = req.body;

      if (!cleanText(sourceText)) {
        return res.status(400).json({
          success: false,
          message:
            "Source / Received Text is required."
        });
      }

      if (!cleanText(workCommand)) {
        return res.status(400).json({
          success: false,
          message:
            "Work / Command is required."
        });
      }

      console.log(
        "Generating Letter + Note File using OpenAI..."
      );

      const result =
        await generateDraftWithOpenAI({
          sourceText,
          workCommand,
          language,
          recipient,
          subject
        });

      res.json({
        success: true,

        ...result,

        model:
          OPENAI_MODEL
      });

    } catch (error) {
      console.error(
        "Generation error:",
        error
      );

      res.status(500).json({
        success: false,

        message:
          error.message ||
          "Draft generation failed."
      });
    }
  }
);

// =====================================================
// CONTINUE / ALTER
// =====================================================

app.post(
  "/api/continue",
  async (req, res) => {
    try {
      const {
        sourceText,
        previousLetter,
        previousNoteFile,
        alterCommand,
        language,
        recipient,
        subject
      } = req.body;

      if (!cleanText(sourceText)) {
        return res.status(400).json({
          success: false,
          message:
            "Original source text is missing."
        });
      }

      if (!cleanText(previousLetter)) {
        return res.status(400).json({
          success: false,
          message:
            "Previous Letter is missing."
        });
      }

      if (!cleanText(previousNoteFile)) {
        return res.status(400).json({
          success: false,
          message:
            "Previous Note File is missing."
        });
      }

      if (!cleanText(alterCommand)) {
        return res.status(400).json({
          success: false,
          message:
            "Continue / Alter Command is required."
        });
      }

      console.log(
        "Applying Continue / Alter command using OpenAI..."
      );

      const result =
        await alterDraftWithOpenAI({
          sourceText,
          previousLetter,
          previousNoteFile,
          alterCommand,
          language,
          recipient,
          subject
        });

      res.json({
        success: true,

        ...result,

        model:
          OPENAI_MODEL
      });

    } catch (error) {
      console.error(
        "Continue / Alter error:",
        error
      );

      res.status(500).json({
        success: false,

        message:
          error.message ||
          "Continue / Alter failed."
      });
    }
  }
);

// =====================================================
// WORD GENERATION
// =====================================================

function textToParagraphs(text) {
  const cleaned = cleanText(text);

  if (!cleaned) {
    return [
      new Paragraph({
        children: [
          new TextRun("")
        ]
      })
    ];
  }

  return cleaned
    .split(/\n/)
    .map((line) => {
      return new Paragraph({
        spacing: {
          after: 120
        },

        children: [
          new TextRun({
            text: line
          })
        ]
      });
    });
}

async function createWordDocument({
  letter,
  noteFile
}) {
  const document =
    new Document({
      sections: [
        {
          properties: {},

          children: [
            new Paragraph({
              text: "OFFICIAL LETTER",
              heading:
                HeadingLevel.HEADING_1,

              alignment:
                AlignmentType.CENTER
            }),

            ...textToParagraphs(letter),

            new Paragraph({
              children: [
                new PageBreak()
              ]
            }),

            new Paragraph({
              text: "NOTE FILE",
              heading:
                HeadingLevel.HEADING_1,

              alignment:
                AlignmentType.CENTER
            }),

            ...textToParagraphs(noteFile)
          ]
        }
      ]
    });

  return Packer.toBuffer(document);
}

// =====================================================
// DOWNLOAD WORD
// =====================================================

app.post(
  "/api/download-word",
  async (req, res) => {
    try {
      const {
        letter,
        noteFile
      } = req.body;

      if (!cleanText(letter)) {
        return res.status(400).json({
          success: false,
          message:
            "Letter is missing."
        });
      }

      if (!cleanText(noteFile)) {
        return res.status(400).json({
          success: false,
          message:
            "Note File is missing."
        });
      }

      const buffer =
        await createWordDocument({
          letter,
          noteFile
        });

      const filename =
        `Revenue_Draft_${Date.now()}.docx`;

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      );

      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${filename}"`
      );

      res.send(buffer);

    } catch (error) {
      console.error(
        "Word generation error:",
        error
      );

      res.status(500).json({
        success: false,
        message:
          "Word file generation failed."
      });
    }
  }
);

// =====================================================
// FRONTEND
// =====================================================

app.use(
  express.static(
    path.join(__dirname, "public")
  )
);

app.use((req, res) => {
  res.sendFile(
    path.join(
      __dirname,
      "public",
      "index.html"
    )
  );
});

// =====================================================
// ERROR HANDLER
// =====================================================

app.use(
  (error, req, res, next) => {
    console.error(
      "Server error:",
      error
    );

    if (
      error instanceof multer.MulterError &&
      error.code === "LIMIT_FILE_SIZE"
    ) {
      return res.status(413).json({
        success: false,
        message:
          `PDF is too large. Maximum allowed size is ${MAX_PDF_MB} MB.`
      });
    }

    res.status(500).json({
      success: false,

      message:
        error.message ||
        "Internal server error."
    });
  }
);

// =====================================================
// START
// =====================================================

app.listen(PORT, () => {
  console.log("");
  console.log(
    "=============================================="
  );

  console.log(
    " Revenue Office Drafting Assistant"
  );

  console.log(
    "=============================================="
  );

  console.log(
    `Server running on port ${PORT}`
  );

  console.log(
    `Gemini configured: ${Boolean(GEMINI_API_KEY)}`
  );

  console.log(
    `Gemini model: ${GEMINI_MODEL}`
  );

  console.log(
    `OpenAI configured: ${Boolean(OPENAI_API_KEY)}`
  );

  console.log(
    `OpenAI model: ${OPENAI_MODEL}`
  );

  console.log(
    "PDF OCR: Gemini"
  );

  console.log(
    "Letter + Note File: OpenAI"
  );

  console.log(
    "Word generation: Node.js"
  );

  console.log(
    "=============================================="
  );
});
