const express  = require('express');
const cors     = require('cors');
const multer   = require('multer');
const axios    = require('axios');
const FormData = require('form-data');
const xlsx     = require('xlsx');
const fs       = require('fs');
const path     = require('path');

// ─────────────────────────────────────────────────────────────────────────────
//  Architecture
//
//  React (localhost:3000)
//    │  multipart/form-data  { file, prompt }
//    ▼
//  Express (localhost:5000)   ← this file
//    │  multipart/form-data  { file, prompt }   forwarded via axios + form-data
//    ▼
//  FastAPI on Kaggle          ← app.py  POST /extract
//    │  runs ollama.chat(qwen2.5vl:7b)
//    ▼
//  returns JSON  { ...extractedFields, model_processing_time }
//    │
//    ▼  Express adds FileName + Timestamp, saves to Excel, replies to React
// ─────────────────────────────────────────────────────────────────────────────

const app = express();
app.use(cors());
app.use(express.json());

// multer keeps uploaded images in memory (no disk writes on local machine)
const upload = multer({ storage: multer.memoryStorage() });

const EXCEL_FILE_PATH = path.join(__dirname, 'extracted_data.xlsx');

// Pinggy tunnel URL pointing to Kaggle FastAPI.
// Updated at runtime via the app header — no server restart needed.
let FASTAPI_URL = 'https://abvun-34-178-44-143.run.pinggy-free.link';

// ─────────────────────────────────────────────────────────────────────────────
//  Helper: forward one image + prompt to FastAPI /extract
//
//  FastAPI expects:
//    POST /extract
//    Content-Type: multipart/form-data
//    file   : UploadFile   (the image binary)
//    prompt : str          (Form field)
//
//  FastAPI returns raw JSON already parsed by json.loads() — so the response
//  body IS the extracted data object (no wrapper like { success, data }).
//  We handle both shapes just in case.
// ─────────────────────────────────────────────────────────────────────────────

async function callFastAPI(fileBuffer, originalName, mimeType, prompt) {
  const base = FASTAPI_URL.replace(/\/$/, '');
  const url  = `${base}/extract`;

  console.log(`[fastapi] POST ${url}  file=${originalName}`);

  // Build multipart body that matches FastAPI's UploadFile + Form signature
  const form = new FormData();
  form.append('file',   fileBuffer, {
    filename:    originalName,
    contentType: mimeType || 'image/jpeg',
  });
  form.append('prompt', prompt);

  const response = await axios.post(url, form, {
    headers: {
      ...form.getHeaders(),   // sets correct multipart boundary
    },
    timeout:          180000, // 3 min — model cold-start on Kaggle can be slow
    maxContentLength: Infinity,
    maxBodyLength:    Infinity,
  });

  return response.data; // already-parsed JSON from FastAPI
}

// ─────────────────────────────────────────────────────────────────────────────
//  Helper: save / append one row to the local Excel file
// ─────────────────────────────────────────────────────────────────────────────

function saveToExcel(rowData) {
  let workbook;
  let existingRows = [];

  if (fs.existsSync(EXCEL_FILE_PATH)) {
    workbook     = xlsx.readFile(EXCEL_FILE_PATH);
    const sheet  = workbook.Sheets[workbook.SheetNames[0]];
    existingRows = xlsx.utils.sheet_to_json(sheet);
  } else {
    workbook = xlsx.utils.book_new();
  }

  existingRows.push(rowData);
  const newSheet = xlsx.utils.json_to_sheet(existingRows);

  if (workbook.SheetNames.length === 0) {
    xlsx.utils.book_append_sheet(workbook, newSheet, 'OMR_Data');
  } else {
    workbook.Sheets[workbook.SheetNames[0]] = newSheet;
  }

  xlsx.writeFile(workbook, EXCEL_FILE_PATH);
}

// ─────────────────────────────────────────────────────────────────────────────
//  Helper: translate axios / network errors into clean JSON responses
//  Returns true when the error was handled (caller must return immediately).
// ─────────────────────────────────────────────────────────────────────────────

function handleFastAPIError(error, res) {
  // Tunnel not running / wrong URL
  if (error.code === 'ECONNREFUSED' || error.code === 'ENOTFOUND') {
    console.error('[fastapi] Cannot connect — is the Pinggy tunnel running on Kaggle?');
    return res.status(503).json({
      success: false,
      error:   'Cannot reach Kaggle FastAPI. Check the Pinggy URL in the app header and make sure the Kaggle notebook is running.',
    }), true;
  }

  // Timeout
  if (error.code === 'ECONNABORTED' || error.message?.includes('timeout')) {
    console.error('[fastapi] Request timed out');
    return res.status(504).json({
      success: false,
      error:   'Kaggle FastAPI timed out. The model may still be loading — wait 30 s and retry.',
    }), true;
  }

  // FastAPI returned a non-2xx (e.g. 422 Unprocessable Entity, 500)
  if (error.response) {
    const status = error.response.status;
    const body   = JSON.stringify(error.response.data).substring(0, 300);
    console.error(`[fastapi] HTTP ${status}:`, body);
    return res.status(502).json({
      success: false,
      error:   `Kaggle FastAPI returned HTTP ${status}. Detail: ${body}`,
    }), true;
  }

  return false;
}

// ─────────────────────────────────────────────────────────────────────────────
//  Route: GET /api/get-url
//  Frontend reads this on mount to pre-fill the Pinggy URL input.
// ─────────────────────────────────────────────────────────────────────────────

app.get('/api/get-url', (req, res) => {
  res.json({ success: true, url: FASTAPI_URL });
});

// ─────────────────────────────────────────────────────────────────────────────
//  Route: POST /api/set-url
//  Frontend calls this when the user updates the Pinggy URL in the header.
//  Body: { url: "https://xxxx.a.pinggy.link" }
// ─────────────────────────────────────────────────────────────────────────────

app.post('/api/set-url', (req, res) => {
  const { url } = req.body;
  if (!url || typeof url !== 'string') {
    return res.status(400).json({ success: false, error: 'Missing url field.' });
  }
  FASTAPI_URL = url.trim().replace(/\/$/, '');
  console.log(`[config] Pinggy URL updated → ${FASTAPI_URL}`);
  res.json({ success: true, url: FASTAPI_URL });
});

// ─────────────────────────────────────────────────────────────────────────────
//  Route: POST /api/extract-template
//  Step 01 — evaluate one sample image to detect field names.
//
//  Forwards to FastAPI /extract with a field-discovery prompt.
//  FastAPI runs the model and returns a JSON object of field→value pairs.
//  Express returns that object to the React UI which renders the checkboxes.
//
//  FileName / Timestamp are NOT added here so they don't appear as checkboxes.
// ─────────────────────────────────────────────────────────────────────────────

app.post('/api/extract-template', upload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ success: false, error: 'No file uploaded.' });
  }

  // Build a field-discovery prompt — sent to FastAPI which passes it to the model
  const hint   = req.body.prompt ? `Additional context: ${req.body.prompt}\n\n` : '';
  const prompt =
    `${hint}` +
    `Analyze this OMR / form image carefully. ` +
    `The form may be written in any language including Hindi or other regional languages. ` +
    `Identify every distinct field name like roll number ,subject code ,name and other fields which are needed to be filled or filled by human in the image. ` +
    `Also look for any barcode or QR code present anywhere on the form — ` +
    `if found, include it as a field with the key "BarcodeNumber" and its decoded/visible value. ` +
    `If no barcode or QR code is present, still include the key "BarcodeNumber" with an empty string value. ` +
    `Return ONLY a single flat JSON object where each key is the English translation of ` +
    `the field name, and the value is the English translation of what that field contains ` +
    `(or an empty string if blank). ` +
    `All keys and all values must be in English regardless of the source language. ` +
    `No markdown, no explanation — raw JSON only.`;

  console.log(`[template] Evaluating: ${req.file.originalname}`);

  try {
    const result = await callFastAPI(
      req.file.buffer,
      req.file.originalname,
      req.file.mimetype,
      prompt
    );

    // FastAPI may return { success: false, error, raw_output } if model output
    // wasn't valid JSON — surface that clearly
    if (result.success === false) {
      console.warn('[template] FastAPI reported failure:', result.error);
      return res.status(422).json({
        success: false,
        error:   result.error || 'Model did not return valid JSON.',
        raw:     result.raw_output,
      });
    }

    // Strip internal FastAPI/model metadata before sending to frontend
    const { model_processing_time, ...fields } = result;

    const fieldCount = Object.keys(fields).length;
    if (fieldCount === 0) {
      return res.status(422).json({
        success: false,
        error:   'No fields detected. Try a clearer template image.',
      });
    }

    console.log(`[template] Detected ${fieldCount} fields  (${model_processing_time}s)`);
    res.json({ success: true, data: fields });

  } catch (error) {
    if (handleFastAPIError(error, res)) return;
    console.error('[template] Unexpected error:', error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
//  Route: POST /api/extract
//  Step 03 — extract selected fields from a single bulk image.
//
//  React sends:
//    file    — the image
//    prompt  — already built by the frontend to include only checked field names
//
//  Express forwards both to FastAPI /extract.
//  FastAPI runs qwen2.5vl:7b and returns the extracted JSON.
//  Express adds FileName + Timestamp, appends a row to Excel, replies to React.
// ─────────────────────────────────────────────────────────────────────────────

app.post('/api/extract', upload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ success: false, error: 'No file uploaded.' });
  }

  const promptText = req.body.prompt ||
    'Extract all fields from this form/OMR image. Return a flat JSON object only.';

  // Reinforce JSON-only output and English translation.
  // FastAPI already uses format="json" with ollama but the explicit language
  // instruction is essential for Hindi / regional language OMR sheets.
  const fullPrompt =
    `${promptText}\n\n` +
    `IMPORTANT: The form may be in any language including Hindi or other regional languages. ` +
    `Translate everything to English. ` +
    `Make sure that you are extracting the details with high precision ,fetch the details form the bubbled parts if it has the bubbled part ` +
    `analyse the requeted part for the extraction and return the details ` +
    `Additionally, always look for any barcode or QR code on the form and include it as ` +
    `"BarcodeNumber" in the response — use the value present beside the barcode in the sheet. ` +
    `add the extracted feilds into perfect slots and return the output (for example : retun the barcode number under the barcode itself not under roll number)do not do this mistakes ` +
    `Return ONLY a flat JSON object with the requested fields as keys in English ` +
    `and their extracted values translated to English. ` +
    `No markdown, no explanation, no extra keys.`;

  console.log(`[extract] Processing: ${req.file.originalname}`);

  try {
    const result = await callFastAPI(
      req.file.buffer,
      req.file.originalname,
      req.file.mimetype,
      fullPrompt
    );

    // Surface FastAPI-level errors (model JSON parse failures etc.)
    if (result.success === false) {
      console.warn('[extract] FastAPI reported failure:', result.error);
      return res.status(422).json({
        success: false,
        error:   result.error || 'Model returned invalid output.',
        raw:     result.raw_output,
      });
    }

    // Pull out model metadata before storing / returning
    const { model_processing_time, ...extractedFields } = result;

    // Stamp file metadata
    extractedFields.FileName  = req.file.originalname;
    extractedFields.Timestamp = new Date().toISOString();

    saveToExcel(extractedFields);
    console.log(`[extract] ✓ ${req.file.originalname}  (model: ${model_processing_time}s)`);

    res.json({ success: true, data: extractedFields });

  } catch (error) {
    if (handleFastAPIError(error, res)) return;
    console.error('[extract] Unexpected error:', error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});


// ─────────────────────────────────────────────────────────────────────────────
//  Route: POST /api/update
//  Accepts the full edited table from the frontend and rewrites the Excel file
//  so that "Load to XL Sheet" always saves the user's edited values, not the
//  raw extracted data.
//
//  Body (JSON): { rows: [ { FileName, Timestamp, field1, field2, ... }, ... ] }
// ─────────────────────────────────────────────────────────────────────────────

app.post('/api/update', (req, res) => {
  const { rows } = req.body;
  if (!Array.isArray(rows) || rows.length === 0) {
    return res.status(400).json({ success: false, error: 'rows array is required.' });
  }

  try {
    let workbook;

    if (fs.existsSync(EXCEL_FILE_PATH)) {
      workbook = xlsx.readFile(EXCEL_FILE_PATH);
    } else {
      workbook = xlsx.utils.book_new();
    }

    const newSheet = xlsx.utils.json_to_sheet(rows);

    if (workbook.SheetNames.length === 0) {
      xlsx.utils.book_append_sheet(workbook, newSheet, 'OMR_Data');
    } else {
      workbook.Sheets[workbook.SheetNames[0]] = newSheet;
    }

    xlsx.writeFile(workbook, EXCEL_FILE_PATH);
    console.log(`[update] Excel rewritten with ${rows.length} edited row(s)`);
    res.json({ success: true, rows: rows.length });
  } catch (err) {
    console.error('[update] Failed to write Excel:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
//  Route: GET /api/download
//  Streams the accumulated Excel file to the browser.
// ─────────────────────────────────────────────────────────────────────────────

app.get('/api/download', (req, res) => {
  if (!fs.existsSync(EXCEL_FILE_PATH)) {
    return res.status(404).json({
      success: false,
      error:   'No data extracted yet — run an extraction first.',
    });
  }
  res.download(EXCEL_FILE_PATH, 'extracted_data.xlsx');
});

// ─────────────────────────────────────────────────────────────────────────────
//  Route: DELETE /api/clear
//  Deletes the local Excel file so the next extraction starts fresh.
// ─────────────────────────────────────────────────────────────────────────────

app.delete('/api/clear', (req, res) => {
  if (fs.existsSync(EXCEL_FILE_PATH)) {
    fs.unlinkSync(EXCEL_FILE_PATH);
    console.log('[clear] Excel file deleted');
  }
  res.json({ success: true, message: 'Data cleared.' });
});

// ─────────────────────────────────────────────────────────────────────────────
//  Route: GET /api/ping
//  Health-check — also verifies the Pinggy tunnel is reachable by hitting
//  FastAPI's GET / which returns { status: "running" }
// ─────────────────────────────────────────────────────────────────────────────

app.get('/api/ping', async (req, res) => {
  try {
    const r = await axios.get(`${FASTAPI_URL.replace(/\/$/, '')}/`, { timeout: 8000 });
    res.json({ success: true, tunnel: 'reachable', fastapi: r.data });
  } catch (err) {
    res.status(502).json({
      success: false,
      tunnel:  'unreachable',
      error:   err.message,
    });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
//  Start
// ─────────────────────────────────────────────────────────────────────────────

const PORT = 5000;
app.listen(PORT, () => {
  console.log('\n┌─────────────────────────────────────────┐');
  console.log('│         OMR Extract — Local Bridge       │');
  console.log('├─────────────────────────────────────────┤');
  console.log(`│  Express  : http://localhost:${PORT}        │`);
  console.log(`│  FastAPI  : ${FASTAPI_URL}`);
  console.log('│  Tunnel   : Pinggy → Kaggle              │');
  console.log('│  Update URL via app header (saved live)  │');
  console.log('└─────────────────────────────────────────┘\n');
});