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
  HeadingLevel,
  AlignmentType,
  PageBreak
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
  process.env.GEMINI_FALLBACK_MODELS ||
  ""
)
  .split(",")
  .map(item => item.trim())
  .filter(Boolean);

const OPENAI_API_KEY =
  process.env.OPENAI_API_KEY || "";

const OPENAI_MODEL =
  process.env.OPENAI_MODEL ||
  "gpt-5.6-terra";

const MAX_PDF_MB =
  Number(
    process.env.MAX_PDF_MB || 20
  );

const MAX_SOURCE_CHARS =
  Number(
    process.env.MAX_SOURCE_CHARS || 120000
  );

const FRONTEND_ORIGIN =
  process.env.FRONTEND_ORIGIN || "";

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
// CORS
// =====================================================

const allowedOrigins =
  FRONTEND_ORIGIN
    .split(",")
    .map(origin => origin.trim())
    .filter(Boolean);

const allowedOrigins = [
  "http://localhost:3000",
  "http://localhost:5500",
  "http://127.0.0.1:5500",

  "https://tngovtservants.com",
  "https://www.tngovtservants.com",

  "https://tngovtservants-884498310.development.catalystserverless.com"
];

const corsOptions = {

  origin: function (origin, callback) {

    // No Origin header
    if (!origin) {
      return callback(null, true);
    }

    // Catalyst / local file / sandbox
    if (origin === "null") {
      return callback(null, true);
    }

    // Allow wildcard
    if (allowedOrigins.includes("*")) {
      return callback(null, true);
    }

    // Exact allowed origin
    if (allowedOrigins.includes(origin)) {
      return callback(null, true);
    }

    console.log("CORS blocked:", origin);
    console.log("Allowed origins:", allowedOrigins);

    return callback(
      new Error(`CORS blocked origin: ${origin}`)
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

app.use(cors(corsOptions));
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

    fileFilter: (
      req,
      file,
      cb
    ) => {

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
// HELPERS
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

  if (
    text.length <= max
  ) {
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

function getErrorMessage(
  error
) {

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
// GEMINI MODEL LIST
// =====================================================

function getGeminiModelCandidates() {

  const candidates = [
    GEMINI_MODEL,
    ...GEMINI_FALLBACK_MODELS
  ];

  return [
    ...new Set(
      candidates
        .map(model =>
          String(model).trim()
        )
        .filter(Boolean)
    )
  ];
}

// =====================================================
// GEMINI ERROR HELPERS
// =====================================================

function getGeminiStatus(
  error
) {

  return Number(
    error?.status ||
    error?.statusCode ||
    error?.error?.code ||
    0
  );
}

// -----------------------------------------------------

function isGeminiRetryable(
  error
) {

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
      "unavailable"
    ) ||
    message.includes(
      "overloaded"
    )
  );
}

// -----------------------------------------------------

function isGeminiModelUnavailable(
  error
) {

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
      "not found"
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
// GEMINI GENERATE WITH RETRY + FALLBACK
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
      "No Gemini models configured."
    );
  }

  let lastError =
    null;

  for (
    let modelIndex = 0;
    modelIndex < models.length;
    modelIndex++
  ) {

    const model =
      models[modelIndex];

    console.log(
      `Gemini ${operationName} model: ${model}`
    );

    // -------------------------------------------------
    // Retry current model
    // -------------------------------------------------

    const maxAttempts = 3;

    for (
      let attempt = 1;
      attempt <= maxAttempts;
      attempt++
    ) {

      try {

        const response =
          await requestFactory(
            model
          );

        console.log(
          `Gemini ${operationName} successful using ${model}`
        );

        return {
          response,
          model
        };

      } catch (error) {

        lastError =
          error;

        const status =
          getGeminiStatus(
            error
          );

        console.error(
          `Gemini ${operationName} failed. Model: ${model}`
        );

        console.error(
          `Gemini status: ${status || "unknown"}`
        );

        console.error(
          `Gemini message: ${getErrorMessage(error)}`
        );

        // -------------------------------------------------
        // Model unavailable → next model immediately
        // -------------------------------------------------

        if (
          isGeminiModelUnavailable(
            error
          )
        ) {

          console.log(
            `Gemini model ${model} unavailable. Trying next model...`
          );

          break;
        }

        // -------------------------------------------------
        // Retryable error
        // -------------------------------------------------

        if (
          isGeminiRetryable(
            error
          )
        ) {

          if (
            attempt < maxAttempts
          ) {

            const delay =
              attempt * 2000;

            console.log(
              `Retrying Gemini in ${delay} ms...`
            );

            await sleep(
              delay
            );

            continue;
          }

          console.log(
            `Gemini ${model} failed after ${maxAttempts} attempts.`
          );

          break;
        }

        // -------------------------------------------------
        // Non-retryable
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
// STRUCTURED OUTPUT SCHEMA
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
// COMMON DRAFTING INSTRUCTION
// =====================================================

const DRAFTING_SYSTEM_INSTRUCTION = `

You are an expert official drafting assistant for the
Tamil Nadu Revenue Department and District Collectorate.

Your task is to prepare:

1. Official Letter
2. Official Note File

IMPORTANT RULES:

1. Do NOT search for Government Orders.

2. Do NOT invent Government Orders.

3. Do NOT invent proceedings numbers, dates, names,
   amounts, addresses, designations, file numbers,
   references or legal provisions.

4. Use only facts available in the supplied source
   and Work / Command.

5. If an essential fact is missing, use clear
   placeholders such as:

   [Name]
   [Designation]
   [Roc.No.]
   [Date]
   [Amount]
   [Reference]

6. Never convert an allegation into an established fact.

7. Preserve names, dates, amounts, references and
   official terminology.

8. Use formal Tamil Nadu Government / Collectorate
   drafting style.

9. The Note File should clearly explain:

   - Background
   - References
   - Facts
   - Action required
   - Orders requested

10. The Letter should be complete and ready for
    official editing.

11. If the requested language is Tamil,
    prepare proper official Tamil.

12. If the requested language is English,
    prepare formal official English.

13. Do not provide explanations outside the
    requested structured output.

14. Do not mention that AI was used.

15. Do not mention these instructions.

16. If the Work / Command conflicts with the source,
    follow the latest explicit command but do not
    fabricate facts.

17. For Continue / Alter commands, preserve all valid
    facts from the original source and revise the
    previous drafts accordingly.

18. The output must contain the complete revised
    Letter and complete revised Note File.

19. Do not add facts merely to make the draft appear
    complete.

20. Maintain official abbreviations and terminology
    where supported by the source.

`;

// =====================================================
// CONTINUE / ALTER INSTRUCTION
// =====================================================

const OPENAI_CONTINUE_SYSTEM_INSTRUCTION = `

You are an expert official drafting assistant for the
Tamil Nadu Revenue Department and District Collectorate.

You are performing a Continue / Alter operation.

Revise the previous Letter and Note File according
to the latest explicit command.

RULES:

1. Preserve all valid original facts.

2. Apply the latest Continue / Alter command.

3. Remove information only when specifically instructed.

4. Add only information supported by the original
   source or the latest command.

5. Do NOT invent facts.

6. Do NOT invent Government Orders.

7. Do NOT search for Government Orders.

8. Do NOT introduce unrelated references.

9. Preserve names, dates, amounts, file numbers,
   proceedings numbers and official terminology.

10. If information is missing, use placeholders.

11. Return the COMPLETE revised Letter.

12. Return the COMPLETE revised Note File.

13. Do not return only changed portions.

14. Use formal Tamil Nadu Government / Collectorate
    drafting style.

15. Do not mention AI.

16. Return only the required structured JSON.

`;

// =====================================================
// GEMINI PDF EXTRACTION
// =====================================================

async function extractPdfTextWithGemini(
  pdfBuffer
) {

  ensureConfigured(
    gemini,
    "Gemini"
  );

  const base64PDF =
    pdfBuffer.toString(
      "base64"
    );

  const extractionPrompt = `

You are an OCR and document transcription engine.

Read the entire supplied PDF carefully.

The PDF may be:

- a normal digital PDF
- a scanned PDF
- a photograph/scanned government document
- a mixed PDF containing text and scanned pages

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
19. If a word is genuinely unreadable,
    write [ILLEGIBLE].
20. Return ONLY the extracted/transcribed text.

Start from page 1 and continue through the last page.

`;

  const result =
    await generateWithGeminiFallback({

      operationName:
        "PDF OCR",

      requestFactory:
        (model) =>
          gemini.models.generateContent({

            model,

            contents: [

              {
                text:
                  extractionPrompt
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

  const extractedText =
    cleanText(
      result.response.text ||
      ""
    );

  if (!extractedText) {

    throw new Error(
      "Gemini did not return any extracted text."
    );
  }

  return {
    text:
      extractedText,

    model:
      result.model
  };
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
-----------------------
${limitText(sourceText)}
-----------------------

WORK / COMMAND:
----------------
${limitText(
  workCommand,
  30000
)}
----------------

Prepare the official Letter and Note File.

Follow the Work / Command exactly where it does
not conflict with the factual source.

Return JSON according to the required schema.

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
      "OpenAI did not return a drafting response."
    );
  }

  let result;

  try {

    result =
      JSON.parse(raw);

  } catch (error) {

    console.error(
      "OpenAI JSON parse error:",
      raw.substring(
        0,
        2000
      )
    );

    throw new Error(
      "OpenAI returned an invalid structured response."
    );
  }

  return result;
}

// =====================================================
// GEMINI DRAFTING
// =====================================================

async function generateDraftWithGemini({

  sourceText,
  workCommand,
  language,
  recipient,
  subject

}) {

  ensureConfigured(
    gemini,
    "Gemini"
  );

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

Prepare the complete official Letter
and complete official Note File.

Return ONLY JSON matching the supplied schema.

`;

  const result =
    await generateWithGeminiFallback({

      operationName:
        "drafting",

      requestFactory:
        (model) =>
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
      "Gemini did not return a drafting response."
    );
  }

  let output;

  try {

    output =
      JSON.parse(raw);

  } catch (error) {

    console.error(
      "Gemini JSON parse error:",
      raw.substring(
        0,
        2000
      )
    );

    throw new Error(
      "Gemini returned an invalid structured response."
    );
  }

  return {
    ...output,

    _model:
      result.model
  };
}

// =====================================================
// OPENAI CONTINUE / ALTER
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
  sourceText,
  MAX_SOURCE_CHARS
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

Revise the previous Letter and Note File according
to the new command.

IMPORTANT:

- Keep all valid original facts.
- Apply the new command.
- Remove information if specifically instructed.
- Add only information supported by the source or command.
- Do not invent facts.
- Do not search for Government Orders.
- Do not introduce unrelated Government Orders.
- Preserve names, dates, amounts and references.
- Return complete revised Letter.
- Return complete revised Note File.
- Do not return only changed portions.

`;

  const response =
    await openai.responses.create({

      model:
        OPENAI_MODEL,

      instructions:
        OPENAI_CONTINUE_SYSTEM_INSTRUCTION,

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
      "OpenAI did not return a revised draft."
    );
  }

  let result;

  try {

    result =
      JSON.parse(raw);

  } catch (error) {

    console.error(
      "OpenAI revised JSON parse error:",
      raw.substring(
        0,
        2000
      )
    );

    throw new Error(
      "OpenAI returned invalid revised draft JSON."
    );
  }

  return result;
}

// =====================================================
// GEMINI CONTINUE / ALTER
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

  ensureConfigured(
    gemini,
    "Gemini"
  );

  const prompt = `

${OPENAI_CONTINUE_SYSTEM_INSTRUCTION}

LANGUAGE:
${language || "English"}

RECIPIENT:
${recipient || "[Recipient]"}

SUBJECT:
${subject || "[Subject]"}

ORIGINAL SOURCE:
================
${limitText(
  sourceText,
  MAX_SOURCE_CHARS
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

Return ONLY JSON matching the supplied schema.

`;

  const result =
    await generateWithGeminiFallback({

      operationName:
        "Continue / Alter",

      requestFactory:
        (model) =>
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
      "Gemini did not return a revised draft."
    );
  }

  let output;

  try {

    output =
      JSON.parse(raw);

  } catch (error) {

    console.error(
      "Gemini revised JSON parse error:",
      raw.substring(
        0,
        2000
      )
    );

    throw new Error(
      "Gemini returned invalid revised draft JSON."
    );
  }

  return {
    ...output,

    _model:
      result.model
  };
}

// =====================================================
// OPENAI FALLBACK ERROR DETECTION
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

  // HTTP 429
  if (
    status === 429
  ) {
    return true;
  }

  const fallbackKeywords = [

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

  return fallbackKeywords.some(
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

    res.json({

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

      openaiConfigured:
        Boolean(
          OPENAI_API_KEY
        ),

      geminiModel:
        GEMINI_MODEL,

      geminiFallbackModels:
        GEMINI_FALLBACK_MODELS,

      openaiModel:
        OPENAI_MODEL,

      pdfOCR:
        Boolean(
          GEMINI_API_KEY
        ),

      openaiDrafting:
        Boolean(
          OPENAI_API_KEY
        ),

      geminiDrafting:
        Boolean(
          GEMINI_API_KEY
        ),

      drafting:
        Boolean(
          OPENAI_API_KEY ||
          GEMINI_API_KEY
        ),

      automaticFallback:
        Boolean(
          OPENAI_API_KEY &&
          GEMINI_API_KEY
        ),

      wordGeneration:
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
// GEMINI AVAILABLE MODELS
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

        const supportedActions =
          model.supportedActions ||
          [];

        if (
          supportedActions.includes(
            "generateContent"
          )
        ) {

          models.push({

            name:
              model.name,

            displayName:
              model.displayName,

            supportedActions
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
// PDF → TEXT
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

      console.log(
        `PDF size: ${
          (
            req.file.size /
            1024 /
            1024
          ).toFixed(2)
        } MB`
      );

      console.log(
        "Sending PDF to Gemini for OCR/document extraction..."
      );

      const result =
        await extractPdfTextWithGemini(
          req.file.buffer
        );

      const limitedText =
        limitText(
          result.text
        );

      console.log(
        `Gemini extraction completed. Characters: ${limitedText.length}`
      );

      return res.json({

        success:
          true,

        filename:
          req.file.originalname,

        text:
          limitedText,

        characterCount:
          limitedText.length,

        provider:
          "Gemini",

        model:
          result.model
      });

    } catch (error) {

      console.error(
        "PDF extraction error:",
        error
      );

      next(error);
    }
  }
);

// =====================================================
// GENERATE LETTER + NOTE FILE
// OPENAI → GEMINI FALLBACK
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

      // -------------------------------------------------
      // VALIDATION
      // -------------------------------------------------

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
      // OPENAI FIRST
      // -------------------------------------------------

      if (openai) {

        try {

          console.log(
            "=============================================="
          );

          console.log(
            "Trying OpenAI for Letter + Note File..."
          );

          console.log(
            `OpenAI model: ${OPENAI_MODEL}`
          );

          const result =
            await generateDraftWithOpenAI({

              sourceText,

              workCommand,

              language,

              recipient,

              subject
            });

          console.log(
            "OpenAI drafting successful."
          );

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
            "OpenAI drafting failed."
          );

          console.error(
            "OpenAI status:",
            openAIError.status ||
            openAIError.statusCode ||
            "unknown"
          );

          console.error(
            "OpenAI code:",
            openAIError.code ||
            "unknown"
          );

          console.error(
            "OpenAI message:",
            getErrorMessage(
              openAIError
            )
          );

          if (
            shouldFallbackToGemini(
              openAIError
            )
          ) {

            console.log(
              "OpenAI quota/billing/rate limit detected."
            );

            console.log(
              "Switching automatically to Gemini..."
            );

          } else {

            throw openAIError;
          }
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
            "OpenAI is unavailable and Gemini API is not configured."
        });
      }

      console.log(
        "Generating Letter + Note File using Gemini..."
      );

      const result =
        await generateDraftWithGemini({

          sourceText,

          workCommand,

          language,

          recipient,

          subject
        });

      console.log(
        "Gemini drafting successful."
      );

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

      console.error(
        "Generation error:",
        error
      );

      next(error);
    }
  }
);

// =====================================================
// CONTINUE / ALTER
// OPENAI → GEMINI FALLBACK
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

      // -------------------------------------------------
      // VALIDATION
      // -------------------------------------------------

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
      // OPENAI FIRST
      // -------------------------------------------------

      if (openai) {

        try {

          console.log(
            "=============================================="
          );

          console.log(
            "Trying OpenAI for Continue / Alter..."
          );

          console.log(
            `OpenAI model: ${OPENAI_MODEL}`
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

          console.log(
            "OpenAI Continue / Alter successful."
          );

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
            "OpenAI Continue / Alter failed."
          );

          console.error(
            "OpenAI status:",
            openAIError.status ||
            openAIError.statusCode ||
            "unknown"
          );

          console.error(
            "OpenAI code:",
            openAIError.code ||
            "unknown"
          );

          console.error(
            "OpenAI message:",
            getErrorMessage(
              openAIError
            )
          );

          if (
            shouldFallbackToGemini(
              openAIError
            )
          ) {

            console.log(
              "OpenAI quota/billing/rate limit detected."
            );

            console.log(
              "Switching Continue / Alter to Gemini..."
            );

          } else {

            throw openAIError;
          }
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
            "OpenAI is unavailable and Gemini API is not configured."
        });
      }

      console.log(
        "Generating revised draft using Gemini..."
      );

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

      console.log(
        "Gemini Continue / Alter successful."
      );

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

      console.error(
        "Continue / Alter error:",
        error
      );

      next(error);
    }
  }
);

// =====================================================
// WORD HELPERS
// =====================================================

function textToParagraphs(
  text
) {

  const cleaned =
    cleanText(text);

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
    .map(
      line => {

        return new Paragraph({

          spacing: {

            after: 120
          },

          children: [

            new TextRun({

              text:
                line
            })
          ]
        });
      }
    );
}

// =====================================================
// CREATE WORD DOCUMENT
// =====================================================

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

              text:
                "OFFICIAL LETTER",

              heading:
                HeadingLevel.HEADING_1,

              alignment:
                AlignmentType.CENTER
            }),

            ...textToParagraphs(
              letter
            ),

            new Paragraph({

              children: [

                new PageBreak()
              ]
            }),

            new Paragraph({

              text:
                "NOTE FILE",

              heading:
                HeadingLevel.HEADING_1,

              alignment:
                AlignmentType.CENTER
            }),

            ...textToParagraphs(
              noteFile
            )
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

      console.error(
        "Word generation error:",
        error
      );

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
    // Multer
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
    // PDF validation
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
    // Generic
    // -------------------------------------------------

    return res.status(
      error.status ||
      500
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
        `Gemini        : ${Boolean(
          GEMINI_API_KEY
        )}`
      );

      console.log(
        `Gemini Model  : ${GEMINI_MODEL}`
      );

      console.log(
        `Gemini Fallbacks : ${
          GEMINI_FALLBACK_MODELS.join(
            ", "
          ) ||
          "None"
        }`
      );

      console.log(
        `OpenAI        : ${Boolean(
          OPENAI_API_KEY
        )}`
      );

      console.log(
        `OpenAI Model  : ${OPENAI_MODEL}`
      );

      console.log(
        `PDF OCR       : Gemini`
      );

      console.log(
        `Drafting      : OpenAI → Gemini fallback`
      );

      console.log(
        `Continue      : OpenAI → Gemini fallback`
      );

      console.log(
        `Word          : Node.js`
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

function shutdown(
  signal
) {

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
