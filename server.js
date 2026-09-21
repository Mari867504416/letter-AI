require("dotenv").config();

const express = require("express");
const cors = require("cors");
const helmet = require("helmet");
const multer = require("multer");

const { GoogleGenAI } = require("@google/genai");
const OpenAI = require("openai");

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  PageBreak,
  HeadingLevel
} = require("docx");

// =====================================================
// APP
// =====================================================

const app = express();

app.set("trust proxy", 1);

const PORT = Number(
  process.env.PORT || 10000
);

// =====================================================
// ENVIRONMENT
// =====================================================

const GEMINI_API_KEY =
  process.env.GEMINI_API_KEY || "";

const GEMINI_MODEL =
  process.env.GEMINI_MODEL ||
  "gemini-3.6-flash";

const GEMINI_FALLBACK_MODELS = (
  process.env.GEMINI_FALLBACK_MODELS || ""
)
  .split(",")
  .map(x => x.trim())
  .filter(Boolean);

const OPENAI_API_KEY =
  process.env.OPENAI_API_KEY || "";

const OPENAI_MODEL =
  process.env.OPENAI_MODEL ||
  "gpt-5.6-terra";

const MAX_PDF_MB =
  Number(process.env.MAX_PDF_MB || 20);

const MAX_SOURCE_CHARS =
  Number(
    process.env.MAX_SOURCE_CHARS || 120000
  );

const FRONTEND_ORIGIN =
  process.env.FRONTEND_ORIGIN || "";

// =====================================================
// AI CLIENTS
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
// CORS
// =====================================================

const allowedOrigins =
  FRONTEND_ORIGIN
    .split(",")
    .map(x => x.trim())
    .filter(Boolean);

const corsOptions = {
  origin: function (origin, callback) {

    if (!origin) {
      return callback(null, true);
    }

    if (origin === "null") {
      return callback(null, true);
    }

    if (allowedOrigins.includes("*")) {
      return callback(null, true);
    }

    if (allowedOrigins.includes(origin)) {
      return callback(null, true);
    }

    console.log(
      "CORS blocked origin:",
      origin
    );

    console.log(
      "Allowed origins:",
      allowedOrigins
    );

    return callback(
      new Error(
        `CORS blocked origin: ${origin}`
      )
    );
  },

  methods: [
    "GET",
    "POST",
    "OPTIONS"
  ],

  allowedHeaders: [
    "Content-Type",
    "Authorization"
  ],

  credentials: false
};

app.use(
  cors(corsOptions)
);

// =====================================================
// SECURITY
// =====================================================

app.use(
  helmet({
    crossOriginResourcePolicy: false
  })
);

// =====================================================
// BODY PARSERS
// =====================================================

app.use(
  express.json({
    limit: "15mb"
  })
);

app.use(
  express.urlencoded({
    extended: true,
    limit: "15mb"
  })
);

// =====================================================
// MULTER
// =====================================================

const upload =
  multer({

    storage:
      multer.memoryStorage(),

    limits: {
      fileSize:
        MAX_PDF_MB *
        1024 *
        1024
    },

    fileFilter:
      (req, file, cb) => {

        const isPDF =
          file.mimetype ===
            "application/pdf" ||
          file.originalname
            .toLowerCase()
            .endsWith(".pdf");

        if (!isPDF) {

          return cb(
            new Error(
              "PDF files only are allowed."
            )
          );
        }

        cb(null, true);
      }
  });

// =====================================================
// BASIC HELPERS
// =====================================================

function cleanText(value) {

  if (
    value === undefined ||
    value === null
  ) {
    return "";
  }

  return String(value)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .trim();
}

// -----------------------------------------------------

function limitText(
  value,
  max = MAX_SOURCE_CHARS
) {

  const text =
    cleanText(value);

  if (text.length <= max) {
    return text;
  }

  return text.substring(
    0,
    max
  );
}

// -----------------------------------------------------

function ensureConfigured(
  client,
  name
) {

  if (!client) {
    throw new Error(
      `${name} API key is not configured in environment variables.`
    );
  }
}

// -----------------------------------------------------

function getErrorMessage(error) {

  return (
    error?.message ||
    String(error) ||
    "Unknown error"
  );
}

// -----------------------------------------------------

function sleep(ms) {

  return new Promise(
    resolve =>
      setTimeout(
        resolve,
        ms
      )
  );
}

// =====================================================
// GEMINI MODEL HELPERS
// =====================================================

function getGeminiModelCandidates() {

  return [
    GEMINI_MODEL,
    ...GEMINI_FALLBACK_MODELS
  ]
    .map(x => x.trim())
    .filter(Boolean)
    .filter(
      (value, index, array) =>
        array.indexOf(value) === index
    );
}

// -----------------------------------------------------

function getGeminiStatus(error) {

  return Number(
    error?.status ||
    error?.statusCode ||
    error?.error?.code ||
    0
  );
}

// -----------------------------------------------------

function isGeminiRetryable(error) {

  const status =
    getGeminiStatus(error);

  const message =
    getErrorMessage(error)
      .toLowerCase();

  return (
    status === 429 ||
    status === 500 ||
    status === 502 ||
    status === 503 ||
    status === 504 ||
    message.includes(
      "high demand"
    ) ||
    message.includes(
      "temporarily unavailable"
    ) ||
    message.includes(
      "overloaded"
    )
  );
}

// -----------------------------------------------------

function isGeminiModelUnavailable(error) {

  const status =
    getGeminiStatus(error);

  const message =
    getErrorMessage(error)
      .toLowerCase();

  return (
    status === 404 ||
    message.includes(
      "model is not found"
    ) ||
    message.includes(
      "no longer available"
    ) ||
    message.includes(
      "not available to new users"
    )
  );
}

// =====================================================
// GEMINI GENERATION WITH RETRY
// =====================================================

async function generateWithGeminiFallback({

  requestFactory,
  operationName

}) {

  ensureConfigured(
    gemini,
    "Gemini"
  );

  const models =
    getGeminiModelCandidates();

  if (!models.length) {

    throw new Error(
      "No Gemini model configured."
    );
  }

  let lastError = null;

  for (
    const model of models
  ) {

    console.log(
      `Gemini ${operationName} model: ${model}`
    );

    for (
      let attempt = 1;
      attempt <= 3;
      attempt++
    ) {

      try {

        const response =
          await requestFactory(
            model
          );

        console.log(
          `Gemini ${operationName} successful: ${model}`
        );

        return {
          response,
          model
        };

      } catch (error) {

        lastError =
          error;

        console.error(
          `Gemini ${operationName} failed`
        );

        console.error(
          `Model: ${model}`
        );

        console.error(
          `Status: ${
            getGeminiStatus(error) ||
            "unknown"
          }`
        );

        console.error(
          `Message: ${
            getErrorMessage(error)
          }`
        );

        // -------------------------------------------------
        // MODEL NOT AVAILABLE
        // -------------------------------------------------

        if (
          isGeminiModelUnavailable(
            error
          )
        ) {

          console.log(
            `Skipping unavailable model: ${model}`
          );

          break;
        }

        // -------------------------------------------------
        // TEMPORARY ERROR
        // -------------------------------------------------

        if (
          isGeminiRetryable(
            error
          )
        ) {

          if (
            attempt < 3
          ) {

            const delay =
              attempt * 2000;

            console.log(
              `Retrying in ${delay} ms...`
            );

            await sleep(
              delay
            );

            continue;
          }

          console.log(
            `Model ${model} failed after 3 attempts.`
          );

          break;
        }

        // -------------------------------------------------
        // OTHER ERROR
        // -------------------------------------------------

        throw error;
      }
    }
  }

  throw (
    lastError ||
    new Error(
      "All Gemini models failed."
    )
  );
}

// =====================================================
// DRAFTING SCHEMA
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
// DRAFTING INSTRUCTION
// =====================================================

const DRAFTING_SYSTEM_INSTRUCTION = `

You are an expert official drafting assistant for the
Tamil Nadu Revenue Department and District Collectorate.

Prepare:

1. Official Letter
2. Official Note File

IMPORTANT:

- Use ONLY the facts supplied by the user.
- Do NOT invent facts.
- Do NOT invent Government Orders.
- Do NOT search for Government Orders.
- Do NOT invent proceedings numbers.
- Do NOT invent dates.
- Do NOT invent names.
- Do NOT invent amounts.
- Do NOT invent references.

Missing information must be shown as:

[Name]
[Designation]
[Roc.No.]
[Date]
[Amount]
[Reference]

Use formal Tamil Nadu Government / District Collectorate
official drafting style.

The Letter should be complete.

The Note File should contain:

Background
References
Facts
Action proposed
Orders requested

For the LETTER:

Preserve logical sections such as:

From
To
Roc.No.
Date
Subject
Ref
Sir/Madam
Body paragraphs
Yours faithfully
Signature
Enclosure
//True Copy//

Do not unnecessarily combine separate paragraphs.

Each distinct body paragraph should be separated by
a blank line.

For Tamil use proper official Tamil.

For English use formal official English.

Do not mention AI.

Return only JSON matching the supplied schema.

`;

// =====================================================
// CONTINUE INSTRUCTION
// =====================================================

const CONTINUE_SYSTEM_INSTRUCTION = `

You are an expert official drafting assistant for the
Tamil Nadu Revenue Department and District Collectorate.

You are revising an existing Letter and Note File.

Preserve all valid facts from the original source.

Apply the new Continue / Alter command.

Do NOT:

- invent facts
- invent Government Orders
- invent dates
- invent amounts
- invent names
- invent references
- introduce unrelated information

Return the COMPLETE revised Letter.

Return the COMPLETE revised Note File.

Do not return only changed portions.

Maintain proper logical paragraphs.

Use formal Tamil Nadu Government / District Collectorate
official drafting style.

Do not mention AI.

Return only JSON matching the supplied schema.

`;

// =====================================================
// GEMINI PDF OCR
// =====================================================

async function extractPdfTextWithGemini(
  pdfBuffer
) {

  const base64PDF =
    pdfBuffer.toString(
      "base64"
    );

  const prompt = `

Read the entire PDF.

Perform OCR if the PDF is scanned.

Extract ALL readable text from ALL pages.

Preserve:

- names
- dates
- numbers
- reference numbers
- file numbers
- Government Order numbers
- amounts
- addresses
- headings
- paragraphs
- tables

Do NOT summarize.

Do NOT translate.

Do NOT interpret.

Do NOT add information.

If genuinely unreadable, use [ILLEGIBLE].

Return ONLY the extracted text.

`;

  const result =
    await generateWithGeminiFallback({

      operationName:
        "PDF OCR",

      requestFactory:
        model =>
          gemini.models.generateContent({

            model,

            contents: [

              {
                text:
                  prompt
              },

              {
                inlineData: {

                  mimeType:
                    "application/pdf",

                  data:
                    base64PDF
                }
              }
            ],

            config: {

              temperature:
                0,

              maxOutputTokens:
                30000
            }
          })
    });

  const text =
    cleanText(
      result.response.text ||
      ""
    );

  if (!text) {

    throw new Error(
      "Gemini returned empty OCR text."
    );
  }

  return {
    text,
    model:
      result.model
  };
}

// =====================================================
// OPENAI GENERATE
// =====================================================

async function generateDraftWithOpenAI({

  sourceText,
  workCommand,
  language,
  recipient,
  subject

}) {

  ensureConfigured(
    openai,
    "OpenAI"
  );

  const prompt = `

LANGUAGE:
${language || "English"}

RECIPIENT:
${recipient || "[Recipient]"}

SUBJECT:
${subject || "[Subject]"}

SOURCE / RECEIVED TEXT:
========================
${limitText(sourceText)}
========================

WORK / COMMAND:
===============
${limitText(
  workCommand,
  30000
)}
===============

Prepare the complete official Letter and Note File.

`;

  const response =
    await openai.responses.create({

      model:
        OPENAI_MODEL,

      instructions:
        DRAFTING_SYSTEM_INSTRUCTION,

      input:
        prompt,

      text: {

        format: {

          type:
            "json_schema",

          name:
            "revenue_drafting_output",

          strict:
            true,

          schema:
            draftingSchema
        }
      }
    });

  const raw =
    cleanText(
      response.output_text ||
      ""
    );

  if (!raw) {

    throw new Error(
      "OpenAI returned empty output."
    );
  }

  try {

    return JSON.parse(
      raw
    );

  } catch {

    throw new Error(
      "OpenAI returned invalid JSON."
    );
  }
}

// =====================================================
// GEMINI GENERATE
// =====================================================

async function generateDraftWithGemini({

  sourceText,
  workCommand,
  language,
  recipient,
  subject

}) {

  const prompt = `

${DRAFTING_SYSTEM_INSTRUCTION}

LANGUAGE:
${language || "English"}

RECIPIENT:
${recipient || "[Recipient]"}

SUBJECT:
${subject || "[Subject]"}

SOURCE / RECEIVED TEXT:
========================
${limitText(sourceText)}
========================

WORK / COMMAND:
===============
${limitText(
  workCommand,
  30000
)}
===============

`;

  const result =
    await generateWithGeminiFallback({

      operationName:
        "Drafting",

      requestFactory:
        model =>
          gemini.models.generateContent({

            model,

            contents:
              prompt,

            config: {

              temperature:
                0,

              maxOutputTokens:
                30000,

              responseMimeType:
                "application/json",

              responseSchema:
                draftingSchema
            }
          })
    });

  const raw =
    cleanText(
      result.response.text ||
      ""
    );

  if (!raw) {

    throw new Error(
      "Gemini returned empty drafting output."
    );
  }

  let output;

  try {

    output =
      JSON.parse(
        raw
      );

  } catch {

    throw new Error(
      "Gemini returned invalid JSON."
    );
  }

  return {
    ...output,
    _model:
      result.model
  };
}

// =====================================================
// OPENAI CONTINUE
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

  ensureConfigured(
    openai,
    "OpenAI"
  );

  const prompt = `

LANGUAGE:
${language || "English"}

RECIPIENT:
${recipient || "[Recipient]"}

SUBJECT:
${subject || "[Subject]"}

ORIGINAL SOURCE:
================
${limitText(
  sourceText
)}
================

PREVIOUS LETTER:
================
${limitText(
  previousLetter,
  60000
)}
================

PREVIOUS NOTE FILE:
===================
${limitText(
  previousNoteFile,
  60000
)}
===================

NEW CONTINUE / ALTER COMMAND:
=============================
${limitText(
  alterCommand,
  30000
)}
=============================

Revise both documents completely.

`;

  const response =
    await openai.responses.create({

      model:
        OPENAI_MODEL,

      instructions:
        CONTINUE_SYSTEM_INSTRUCTION,

      input:
        prompt,

      text: {

        format: {

          type:
            "json_schema",

          name:
            "revenue_drafting_output",

          strict:
            true,

          schema:
            draftingSchema
        }
      }
    });

  const raw =
    cleanText(
      response.output_text ||
      ""
    );

  if (!raw) {

    throw new Error(
      "OpenAI returned empty revised output."
    );
  }

  try {

    return JSON.parse(
      raw
    );

  } catch {

    throw new Error(
      "OpenAI returned invalid revised JSON."
    );
  }
}

// =====================================================
// GEMINI CONTINUE
// =====================================================

async function alterDraftWithGemini({

  sourceText,
  previousLetter,
  previousNoteFile,
  alterCommand,
  language,
  recipient,
  subject

}) {

  const prompt = `

${CONTINUE_SYSTEM_INSTRUCTION}

LANGUAGE:
${language || "English"}

RECIPIENT:
${recipient || "[Recipient]"}

SUBJECT:
${subject || "[Subject]"}

ORIGINAL SOURCE:
================
${limitText(
  sourceText
)}
================

PREVIOUS LETTER:
================
${limitText(
  previousLetter,
  60000
)}
================

PREVIOUS NOTE FILE:
===================
${limitText(
  previousNoteFile,
  60000
)}
===================

NEW CONTINUE / ALTER COMMAND:
=============================
${limitText(
  alterCommand,
  30000
)}
=============================

`;

  const result =
    await generateWithGeminiFallback({

      operationName:
        "Continue / Alter",

      requestFactory:
        model =>
          gemini.models.generateContent({

            model,

            contents:
              prompt,

            config: {

              temperature:
                0,

              maxOutputTokens:
                30000,

              responseMimeType:
                "application/json",

              responseSchema:
                draftingSchema
            }
          })
    });

  const raw =
    cleanText(
      result.response.text ||
      ""
    );

  if (!raw) {

    throw new Error(
      "Gemini returned empty revised output."
    );
  }

  let output;

  try {

    output =
      JSON.parse(
        raw
      );

  } catch {

    throw new Error(
      "Gemini returned invalid revised JSON."
    );
  }

  return {
    ...output,
    _model:
      result.model
  };
}

// =====================================================
// OPENAI FALLBACK CHECK
// =====================================================

function shouldFallbackToGemini(
  error
) {

  if (!error) {
    return false;
  }

  const status =
    Number(
      error.status ||
      error.statusCode ||
      0
    );

  const code =
    String(
      error.code || ""
    ).toLowerCase();

  const message =
    String(
      error.message || ""
    ).toLowerCase();

  if (status === 429) {
    return true;
  }

  const keywords = [

    "insufficient_quota",
    "credit_balance_exhausted",
    "insufficient credit",
    "quota",
    "billing",
    "rate limit",
    "rate_limit",
    "exceeded your current quota",
    "too many requests"

  ];

  return keywords.some(
    keyword =>
      code.includes(keyword) ||
      message.includes(keyword)
  );
}

// =====================================================
// HEALTH
// =====================================================

app.get(
  "/api/health",
  (req, res) => {

    return res.json({

      success:
        true,

      service:
        "Revenue Office Drafting Assistant API",

      backend:
        "Render",

      frontend:
        "Catalyst",

      geminiConfigured:
        Boolean(
          GEMINI_API_KEY
        ),

      geminiModel:
        GEMINI_MODEL,

      geminiFallbackModels:
        GEMINI_FALLBACK_MODELS,

      openaiConfigured:
        Boolean(
          OPENAI_API_KEY
        ),

      openaiModel:
        OPENAI_MODEL,

      automaticFallback:
        Boolean(
          GEMINI_API_KEY &&
          OPENAI_API_KEY
        ),

      pdfOCR:
        "Gemini",

      drafting:
        "OpenAI → Gemini",

      wordGeneration:
        true,

      wordAlignment:
        true,

      frontendConfigured:
        Boolean(
          FRONTEND_ORIGIN
        ),

      timestamp:
        new Date().toISOString()
    });
  }
);

// =====================================================
// GEMINI MODEL LIST
// =====================================================

app.get(
  "/api/gemini-models",
  async (req, res) => {

    try {

      if (!gemini) {

        return res.status(500).json({

          success:
            false,

          message:
            "GEMINI_API_KEY is not configured."
        });
      }

      const models = [];

      for await (
        const model
        of gemini.models.list()
      ) {

        const actions =
          model.supportedActions ||
          [];

        if (
          actions.includes(
            "generateContent"
          )
        ) {

          models.push({

            name:
              model.name,

            displayName:
              model.displayName,

            supportedActions:
              actions
          });
        }
      }

      return res.json({

        success:
          true,

        count:
          models.length,

        models
      });

    } catch (error) {

      console.error(
        "Gemini model list error:",
        error
      );

      return res.status(500).json({

        success:
          false,

        message:
          getErrorMessage(
            error
          )
      });
    }
  }
);

// =====================================================
// PDF TO TEXT
// =====================================================

app.post(
  "/api/pdf-to-text",
  upload.single("pdf"),
  async (
    req,
    res,
    next
  ) => {

    try {

      if (!req.file) {

        return res.status(400).json({

          success:
            false,

          message:
            "Please upload a PDF file."
        });
      }

      console.log(
        `PDF received: ${req.file.originalname}`
      );

      const result =
        await extractPdfTextWithGemini(
          req.file.buffer
        );

      const text =
        limitText(
          result.text
        );

      return res.json({

        success:
          true,

        filename:
          req.file.originalname,

        text,

        characterCount:
          text.length,

        provider:
          "Gemini",

        model:
          result.model
      });

    } catch (error) {

      next(error);
    }
  }
);

// =====================================================
// GENERATE
// =====================================================

app.post(
  "/api/generate",
  async (
    req,
    res,
    next
  ) => {

    try {

      const {
        sourceText,
        workCommand,
        language,
        recipient,
        subject
      } = req.body;

      if (
        !cleanText(
          sourceText
        )
      ) {

        return res.status(400).json({

          success:
            false,

          message:
            "Source / Received Text is required."
        });
      }

      if (
        !cleanText(
          workCommand
        )
      ) {

        return res.status(400).json({

          success:
            false,

          message:
            "Work / Command is required."
        });
      }

      // -------------------------------------------------
      // OPENAI
      // -------------------------------------------------

      if (openai) {

        try {

          console.log(
            "Trying OpenAI for Letter + Note File..."
          );

          const result =
            await generateDraftWithOpenAI({

              sourceText,

              workCommand,

              language,

              recipient,

              subject
            });

          return res.json({

            success:
              true,

            ...result,

            provider:
              "OpenAI",

            model:
              OPENAI_MODEL,

            fallback:
              false
          });

        } catch (
          openAIError
        ) {

          console.error(
            "OpenAI drafting failed:",
            getErrorMessage(
              openAIError
            )
          );

          if (
            !shouldFallbackToGemini(
              openAIError
            )
          ) {

            throw openAIError;
          }

          console.log(
            "Switching automatically to Gemini..."
          );
        }
      }

      // -------------------------------------------------
      // GEMINI FALLBACK
      // -------------------------------------------------

      if (!gemini) {

        return res.status(503).json({

          success:
            false,

          message:
            "OpenAI failed and Gemini is not configured."
        });
      }

      const result =
        await generateDraftWithGemini({

          sourceText,

          workCommand,

          language,

          recipient,

          subject
        });

      return res.json({

        success:
          true,

        ...result,

        provider:
          "Gemini",

        model:
          result._model ||
          GEMINI_MODEL,

        fallback:
          Boolean(
            openai
          )
      });

    } catch (error) {

      next(error);
    }
  }
);

// =====================================================
// CONTINUE / ALTER
// =====================================================

app.post(
  "/api/continue",
  async (
    req,
    res,
    next
  ) => {

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

      if (
        !cleanText(
          sourceText
        )
      ) {

        return res.status(400).json({

          success:
            false,

          message:
            "Original source text is missing."
        });
      }

      if (
        !cleanText(
          previousLetter
        )
      ) {

        return res.status(400).json({

          success:
            false,

          message:
            "Previous Letter is missing."
        });
      }

      if (
        !cleanText(
          previousNoteFile
        )
      ) {

        return res.status(400).json({

          success:
            false,

          message:
            "Previous Note File is missing."
        });
      }

      if (
        !cleanText(
          alterCommand
        )
      ) {

        return res.status(400).json({

          success:
            false,

          message:
            "Continue / Alter Command is required."
        });
      }

      // -------------------------------------------------
      // OPENAI
      // -------------------------------------------------

      if (openai) {

        try {

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

          return res.json({

            success:
              true,

            ...result,

            provider:
              "OpenAI",

            model:
              OPENAI_MODEL,

            fallback:
              false
          });

        } catch (
          openAIError
        ) {

          console.error(
            "OpenAI Continue / Alter failed:",
            getErrorMessage(
              openAIError
            )
          );

          if (
            !shouldFallbackToGemini(
              openAIError
            )
          ) {

            throw openAIError;
          }

          console.log(
            "Switching Continue / Alter to Gemini..."
          );
        }
      }

      // -------------------------------------------------
      // GEMINI
      // -------------------------------------------------

      if (!gemini) {

        return res.status(503).json({

          success:
            false,

          message:
            "OpenAI failed and Gemini is not configured."
        });
      }

      const result =
        await alterDraftWithGemini({

          sourceText,

          previousLetter,

          previousNoteFile,

          alterCommand,

          language,

          recipient,

          subject
        });

      return res.json({

        success:
          true,

        ...result,

        provider:
          "Gemini",

        model:
          result._model ||
          GEMINI_MODEL,

        fallback:
          Boolean(
            openai
          )
      });

    } catch (error) {

      next(error);
    }
  }
);

// =====================================================
// WORD DOCUMENT HELPERS
// =====================================================

// Detect Tamil
function containsTamil(text) {

  return /[\u0B80-\u0BFF]/.test(
    text || ""
  );
}

// -----------------------------------------------------

function getFont(text) {

  return containsTamil(text)
    ? "Nirmala UI"
    : "Times New Roman";
}

// -----------------------------------------------------

function makeTextRun(
  text,
  options = {}
) {

  return new TextRun({

    text,

    font:
      options.font ||
      getFont(text),

    size:
      options.size ||
      24,

    bold:
      options.bold ||
      false,

    italics:
      options.italics ||
      false,

    break:
      options.break
  });
}

// =====================================================
// LETTER CLASSIFICATION
// =====================================================

function classifyLetterLine(
  line
) {

  const text =
    cleanText(line);

  const lower =
    text.toLowerCase();

  // Empty
  if (!text) {
    return "empty";
  }

  // From
  if (
    /^(from\s*:|அனுப்புநர்\s*:)/i.test(
      text
    )
  ) {
    return "from";
  }

  // To
  if (
    /^(to\s*:|பெறுநர்\s*:)/i.test(
      text
    )
  ) {
    return "to";
  }

  // Subject
  if (
    /^(subject\s*:|sub\s*:|பொருள்\s*:)/i.test(
      text
    )
  ) {
    return "subject";
  }

  // Reference
  if (
    /^(ref\s*:|reference\s*:|மேற்கோள்\s*:)/i.test(
      text
    )
  ) {
    return "reference";
  }

  // Sir / Madam
  if (
    /^(sir|madam|sir\/madam|மதிப்பிற்குரிய)/i.test(
      text
    )
  ) {
    return "salutation";
  }

  // Closing
  if (
    /^(yours faithfully|yours sincerely|faithfully|தங்கள் உண்மையுள்ள)/i.test(
      text
    )
  ) {
    return "closing";
  }

  // Enclosure
  if (
    /^(encl|enclosure|enclosures|இணைப்பு)/i.test(
      text
    )
  ) {
    return "enclosure";
  }

  // True copy
  if (
    lower.includes(
      "//true copy//"
    ) ||
    lower.includes(
      "true copy"
    ) ||
    text.includes(
      "//மெய்ப்பிரதி//"
    )
  ) {
    return "truecopy";
  }

  // Roc / Date
  if (
    /^(roc\.?|rc\.?|proceedings|ந\.க\.|நாள்\s*:|date\s*:)/i.test(
      text
    )
  ) {
    return "referenceHeader";
  }

  // Signature/designation
  if (
    /^(district collector|collector|personal assistant|tahsildar|revenue divisional officer|senior revenue inspector|மாவட்ட ஆட்சியர்|வட்டாட்சியர்)/i.test(
      text
    )
  ) {
    return "signature";
  }

  return "body";
}

// =====================================================
// WORD LETTER PARAGRAPHS
// =====================================================

function buildLetterParagraphs(
  letter
) {

  const normalized =
    cleanText(letter);

  const lines =
    normalized
      .split("\n")
      .map(x => x.trim());

  const paragraphs = [];

  let bodyBuffer = [];

  function flushBody() {

    if (
      bodyBuffer.length === 0
    ) {
      return;
    }

    const bodyText =
      bodyBuffer.join(" ");

    paragraphs.push(

      new Paragraph({

        alignment:
          AlignmentType.JUSTIFIED,

        spacing: {

          line:
            320,

          after:
            160
        },

        indent: {

          firstLine:
            567
        },

        children: [

          makeTextRun(
            bodyText,
            {
              size: 24
            }
          )
        ]
      })
    );

    bodyBuffer = [];
  }

  for (
    let i = 0;
    i < lines.length;
    i++
  ) {

    const line =
      lines[i];

    if (!line) {

      flushBody();

      continue;
    }

    const type =
      classifyLetterLine(
        line
      );

    // -------------------------------------------------
    // BODY
    // -------------------------------------------------

    if (
      type === "body"
    ) {

      bodyBuffer.push(
        line
      );

      continue;
    }

    flushBody();

    // -------------------------------------------------
    // FROM
    // -------------------------------------------------

    if (
      type === "from"
    ) {

      paragraphs.push(

        new Paragraph({

          alignment:
            AlignmentType.LEFT,

          spacing: {
            after: 80
          },

          children: [

            makeTextRun(
              line,
              {
                bold: true
              }
            )
          ]
        })
      );

      continue;
    }

    // -------------------------------------------------
    // TO
    // -------------------------------------------------

    if (
      type === "to"
    ) {

      paragraphs.push(

        new Paragraph({

          alignment:
            AlignmentType.LEFT,

          spacing: {
            after: 80
          },

          children: [

            makeTextRun(
              line,
              {
                bold: true
              }
            )
          ]
        })
      );

      continue;
    }

    // -------------------------------------------------
    // SUBJECT
    // -------------------------------------------------

    if (
      type === "subject"
    ) {

      paragraphs.push(

        new Paragraph({

          alignment:
            AlignmentType.LEFT,

          spacing: {

            before:
              120,

            after:
              160
          },

          children: [

            makeTextRun(
              line,
              {
                bold: true
              }
            )
          ]
        })
      );

      continue;
    }

    // -------------------------------------------------
    // REFERENCE
    // -------------------------------------------------

    if (
      type === "reference"
    ) {

      paragraphs.push(

        new Paragraph({

          alignment:
            AlignmentType.LEFT,

          spacing: {
            after: 100
          },

          children: [

            makeTextRun(
              line,
              {
                bold: false
              }
            )
          ]
        })
      );

      continue;
    }

    // -------------------------------------------------
    // ROC / DATE
    // -------------------------------------------------

    if (
      type ===
      "referenceHeader"
    ) {

      paragraphs.push(

        new Paragraph({

          alignment:
            AlignmentType.RIGHT,

          spacing: {
            after: 80
          },

          children: [

            makeTextRun(
              line
            )
          ]
        })
      );

      continue;
    }

    // -------------------------------------------------
    // SALUTATION
    // -------------------------------------------------

    if (
      type ===
      "salutation"
    ) {

      paragraphs.push(

        new Paragraph({

          alignment:
            AlignmentType.LEFT,

          spacing: {

            before:
              160,

            after:
              160
          },

          children: [

            makeTextRun(
              line
            )
          ]
        })
      );

      continue;
    }

    // -------------------------------------------------
    // CLOSING
    // -------------------------------------------------

    if (
      type === "closing"
    ) {

      paragraphs.push(

        new Paragraph({

          alignment:
            AlignmentType.RIGHT,

          spacing: {

            before:
              240,

            after:
              80
          },

          children: [

            makeTextRun(
              line
            )
          ]
        })
      );

      continue;
    }

    // -------------------------------------------------
    // SIGNATURE
    // -------------------------------------------------

    if (
      type === "signature"
    ) {

      paragraphs.push(

        new Paragraph({

          alignment:
            AlignmentType.RIGHT,

          spacing: {

            after:
              80
          },

          children: [

            makeTextRun(
              line,
              {
                bold:
                  true
              }
            )
          ]
        })
      );

      continue;
    }

    // -------------------------------------------------
    // ENCLOSURE
    // -------------------------------------------------

    if (
      type === "enclosure"
    ) {

      paragraphs.push(

        new Paragraph({

          alignment:
            AlignmentType.LEFT,

          spacing: {

            before:
              160,

            after:
              120
          },

          children: [

            makeTextRun(
              line,
              {
                bold:
                  true
              }
            )
          ]
        })
      );

      continue;
    }

    // -------------------------------------------------
    // TRUE COPY
    // -------------------------------------------------

    if (
      type === "truecopy"
    ) {

      paragraphs.push(

        new Paragraph({

          alignment:
            AlignmentType.CENTER,

          spacing: {

            before:
              240,

            after:
              120
          },

          children: [

            makeTextRun(
              line,
              {
                bold:
                  true
              }
            )
          ]
        })
      );

      continue;
    }
  }

  flushBody();

  return paragraphs;
}

// =====================================================
// NOTE FILE PARAGRAPHS
// =====================================================

function buildNoteFileParagraphs(
  noteFile
) {

  const normalized =
    cleanText(
      noteFile
    );

  const lines =
    normalized
      .split("\n")
      .map(x => x.trim());

  const paragraphs = [];

  let buffer = [];

  function flush() {

    if (
      buffer.length === 0
    ) {
      return;
    }

    const text =
      buffer.join(" ");

    paragraphs.push(

      new Paragraph({

        alignment:
          AlignmentType.JUSTIFIED,

        spacing: {

          line:
            320,

          after:
            160
        },

        indent: {

          firstLine:
            567
        },

        children: [

          makeTextRun(
            text
          )
        ]
      })
    );

    buffer = [];
  }

  for (
    const line of lines
  ) {

    if (!line) {

      flush();

      continue;
    }

    // Heading detection
    const isHeading =
      /^(subject|sub|ref|reference|submitted|background|facts|proposal|orders requested|note file|न\.க\.|பொருள்|முன்னிலை|குறிப்பு|சமர்ப்பிக்கப்படுகிறது)/i
        .test(line);

    if (isHeading) {

      flush();

      paragraphs.push(

        new Paragraph({

          alignment:
            AlignmentType.LEFT,

          spacing: {

            before:
              160,

            after:
              120
          },

          children: [

            makeTextRun(
              line,
              {
                bold:
                  true
              }
            )
          ]
        })
      );

      continue;
    }

    buffer.push(line);
  }

  flush();

  return paragraphs;
}

// =====================================================
// CREATE WORD
// =====================================================

async function createWordDocument({

  letter,
  noteFile

}) {

  const letterParagraphs =
    buildLetterParagraphs(
      letter
    );

  const noteParagraphs =
    buildNoteFileParagraphs(
      noteFile
    );

  const document =
    new Document({

      styles: {

        default: {

          document: {

            run: {

              font:
                "Nirmala UI",

              size:
                24
            },

            paragraph: {

              spacing: {

                line:
                  320,

                after:
                  120
              }
            }
          }
        }
      },

      sections: [

        {

          properties: {

            page: {

              size: {

                width:
                  11906,

                height:
                  16838
              },

              margin: {

                top:
                  1134,

                right:
                  1134,

                bottom:
                  1134,

                left:
                  1134
              }
            }
          },

          children: [

            // -------------------------------------------------
            // LETTER TITLE
            // -------------------------------------------------

            new Paragraph({

              alignment:
                AlignmentType.CENTER,

              spacing: {

                after:
                  300
              },

              children: [

                new TextRun({

                  text:
                    "OFFICIAL LETTER",

                  bold:
                    true,

                  font:
                    "Nirmala UI",

                  size:
                    28
                })
              ]
            }),

            // -------------------------------------------------
            // LETTER
            // -------------------------------------------------

            ...letterParagraphs,

            // -------------------------------------------------
            // PAGE BREAK
            // -------------------------------------------------

            new Paragraph({

              children: [

                new PageBreak()
              ]
            }),

            // -------------------------------------------------
            // NOTE FILE TITLE
            // -------------------------------------------------

            new Paragraph({

              alignment:
                AlignmentType.CENTER,

              spacing: {

                after:
                  300
              },

              children: [

                new TextRun({

                  text:
                    "NOTE FILE",

                  bold:
                    true,

                  font:
                    "Nirmala UI",

                  size:
                    28
                })
              ]
            }),

            // -------------------------------------------------
            // NOTE FILE
            // -------------------------------------------------

            ...noteParagraphs
          ]
        }
      ]
    });

  return Packer.toBuffer(
    document
  );
}

// =====================================================
// DOWNLOAD WORD
// =====================================================

app.post(
  "/api/download-word",
  async (
    req,
    res,
    next
  ) => {

    try {

      const {
        letter,
        noteFile
      } = req.body;

      if (
        !cleanText(
          letter
        )
      ) {

        return res.status(400).json({

          success:
            false,

          message:
            "Letter is missing."
        });
      }

      if (
        !cleanText(
          noteFile
        )
      ) {

        return res.status(400).json({

          success:
            false,

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

      res.setHeader(
        "Content-Length",
        buffer.length
      );

      return res.send(
        buffer
      );

    } catch (error) {

      next(error);
    }
  }
);

// =====================================================
// API 404
// =====================================================

app.use(
  "/api",
  (
    req,
    res
  ) => {

    return res.status(404).json({

      success:
        false,

      message:
        "API endpoint not found.",

      path:
        req.originalUrl
    });
  }
);

// =====================================================
// GLOBAL ERROR HANDLER
// =====================================================

app.use(
  (
    error,
    req,
    res,
    next
  ) => {

    console.error(
      "Server error:",
      error
    );

    // -------------------------------------------------
    // MULTER FILE SIZE
    // -------------------------------------------------

    if (
      error instanceof
        multer.MulterError
    ) {

      if (
        error.code ===
        "LIMIT_FILE_SIZE"
      ) {

        return res.status(413).json({

          success:
            false,

          message:
            `PDF is too large. Maximum allowed size is ${MAX_PDF_MB} MB.`
        });
      }
    }

    // -------------------------------------------------
    // CORS
    // -------------------------------------------------

    if (
      String(
        error.message || ""
      ).startsWith(
        "CORS blocked"
      )
    ) {

      return res.status(403).json({

        success:
          false,

        message:
          error.message
      });
    }

    // -------------------------------------------------
    // PDF TYPE
    // -------------------------------------------------

    if (
      error.message ===
      "PDF files only are allowed."
    ) {

      return res.status(400).json({

        success:
          false,

        message:
          error.message
      });
    }

    // -------------------------------------------------
    // GENERIC
    // -------------------------------------------------

    return res.status(
      error.status || 500
    ).json({

      success:
        false,

      message:
        getErrorMessage(
          error
        )
    });
  }
);

// =====================================================
// START SERVER
// =====================================================

const server =
  app.listen(
    PORT,
    () => {

      console.log("");
      console.log(
        "=================================================="
      );

      console.log(
        " Revenue Office Drafting Assistant API"
      );

      console.log(
        "=================================================="
      );

      console.log(
        `Backend       : Render`
      );

      console.log(
        `Port          : ${PORT}`
      );

      console.log(
        `Frontend      : Catalyst`
      );

      console.log(
        `Frontend CORS : ${
          FRONTEND_ORIGIN ||
          "NOT CONFIGURED"
        }`
      );

      console.log(
        `Gemini        : ${
          Boolean(GEMINI_API_KEY)
        }`
      );

      console.log(
        `Gemini Model  : ${GEMINI_MODEL}`
      );

      console.log(
        `Gemini Fallbacks : ${
          GEMINI_FALLBACK_MODELS.join(
            ", "
          ) || "None"
        }`
      );

      console.log(
        `OpenAI        : ${
          Boolean(OPENAI_API_KEY)
        }`
      );

      console.log(
        `OpenAI Model  : ${OPENAI_MODEL}`
      );

      console.log(
        `PDF OCR       : Gemini`
      );

      console.log(
        `Drafting      : OpenAI → Gemini`
      );

      console.log(
        `Word          : Node.js`
      );

      console.log(
        `Word Alignment: ENABLED`
      );

      console.log(
        `A4 Page       : ENABLED`
      );

      console.log(
        `Max PDF       : ${MAX_PDF_MB} MB`
      );

      console.log(
        `Max Source    : ${MAX_SOURCE_CHARS} chars`
      );

      console.log(
        "=================================================="
      );

      console.log("");
    }
  );

// =====================================================
// GRACEFUL SHUTDOWN
// =====================================================

function shutdown(signal) {

  console.log(
    `${signal} received. Shutting down server...`
  );

  server.close(
    () => {

      console.log(
        "Server closed."
      );

      process.exit(0);
    }
  );
}

process.once(
  "SIGTERM",
  () =>
    shutdown("SIGTERM")
);

process.once(
  "SIGINT",
  () =>
    shutdown("SIGINT")
);
