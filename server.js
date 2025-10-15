require('dotenv').config();

const Webex = require('webex');
const axios = require('axios').default;
const { HttpsProxyAgent } = require('https-proxy-agent');
const fs = require('fs');
const fsp = fs.promises;
const path = require('path');
const os = require('os');
const { v4: uuidv4 } = require('uuid');
const { execFile } = require('child_process');

const BOT_TOKEN = process.env.WEBEX_ACCESS_TOKEN;
if (!BOT_TOKEN) {
  console.error('ERROR: Missing WEBEX_ACCESS_TOKEN in .env');
  process.exit(1);
}

const PYTHON_BIN = process.env.PYTHON_BIN || 'python3';
const PPTX_EXTRACTOR_SCRIPT =
  process.env.PPTX_EXTRACTOR_SCRIPT || path.join(__dirname, 'pptx_to_json.py');

const TMP_DIR = path.join(os.tmpdir(), 'webex-pptx-json-bot');
fs.mkdirSync(TMP_DIR, { recursive: true });

const MAX_MARKDOWN_BYTES = 7000; // API limit is ~7439 bytes; keep a safety margin. :contentReference[oaicite:1]{index=1}

const proxyUrl = process.env.HTTPS_PROXY || process.env.HTTP_PROXY || null;
const axiosClient = axios.create({
  proxy: false,
  httpsAgent: proxyUrl ? new HttpsProxyAgent(proxyUrl) : undefined,
  timeout: 30000
});

const webex = Webex.init({
  credentials: { access_token: BOT_TOKEN },
  config: {
    logger: { level: 'error' }
  }
});

let me = null;

function byteLength(str) {
  return Buffer.byteLength(str, 'utf8');
}

function chunkMarkdown(md, maxBytes) {
  const chunks = [];
  let current = '';
  for (const line of md.split('\n')) {
    const candidate = current ? current + '\n' + line : line;
    if (byteLength(candidate) > maxBytes) {
      if (current) chunks.push(current);
      // If a single line is too long, hard-split it
      if (byteLength(line) > maxBytes) {
        let start = 0;
        const buf = Buffer.from(line, 'utf8');
        while (start < buf.length) {
          const end = Math.min(start + maxBytes, buf.length);
          chunks.push(buf.slice(start, end).toString('utf8'));
          start = end;
        }
        current = '';
      } else {
        current = line;
      }
    } else {
      current = candidate;
    }
  }
  if (current) chunks.push(current);
  return chunks;
}

function markdownCodeFence(jsonStr) {
  return '```json\n' + jsonStr + '\n```';
}

function parseFilenameFromContentDisposition(cd) {
  if (!cd) return null;
  // Example: attachment; filename="example.pptx"
  const match = cd.match(/filename\*?=(?:UTF-8'')?"?([^\";]+)"?/i);
  if (match) return match[1];
  return null;
}

async function headContent(url) {
  const res = await axiosClient.head(url, {
    headers: { Authorization: `Bearer ${BOT_TOKEN}` },
    validateStatus: () => true
  });
  return {
    status: res.status,
    headers: res.headers
  };
}

// Basic retry for 423 Locked (anti-malware scanning) with Retry-After. :contentReference[oaicite:2]{index=2}
async function downloadWithRetry(url, responseType = 'arraybuffer', maxAttempts = 4) {
  let attempt = 0;
  while (attempt < maxAttempts) {
    const res = await axiosClient.get(url, {
      headers: { Authorization: `Bearer ${BOT_TOKEN}` },
      responseType,
      validateStatus: () => true
    });
    if (res.status === 200) return res;
    if (res.status === 423) {
      const ra = parseInt(res.headers['retry-after'] || '5', 10);
      await new Promise((r) => setTimeout(r, (isNaN(ra) ? 5 : ra) * 1000));
      attempt += 1;
      continue;
    }
    if (res.status === 410) {
      throw new Error('File failed malware scan (410 Gone)');
    }
    if (res.status === 428) {
      // not scannable – try with allow=unscannable per docs
      url = url.includes('?') ? `${url}&allow=unscannable` : `${url}?allow=unscannable`;
      attempt += 1;
      continue;
    }
    throw new Error(`Unexpected HTTP ${res.status} when downloading file`);
  }
  throw new Error('Timed out waiting for file to be available');
}

async function sendMarkdown(roomId, markdown) {
  return webex.messages.create({ roomId, markdown });
}

async function sendJsonInChunks(roomId, obj, titlePrefix = 'PPTX → JSON') {
  const pretty = JSON.stringify(obj, null, 2);
  const fenced = markdownCodeFence(pretty);
  const chunks = chunkMarkdown(fenced, MAX_MARKDOWN_BYTES);
  if (chunks.length === 1) {
    await sendMarkdown(roomId, chunks[0]);
    return;
  }
  await sendMarkdown(roomId, `**${titlePrefix}** (split into ${chunks.length} parts)`);
  for (const chunk of chunks) {
    await sendMarkdown(roomId, chunk);
  }
}

async function runPythonExtractor(pptxPath) {
  return new Promise((resolve, reject) => {
    execFile(PYTHON_BIN, [PPTX_EXTRACTOR_SCRIPT, pptxPath], { maxBuffer: 20 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) {
        reject(new Error(`Python extractor failed: ${stderr || err.message}`));
        return;
      }
      try {
        const obj = JSON.parse(stdout);
        resolve(obj);
      } catch (e) {
        reject(new Error(`Invalid JSON from extractor: ${e.message}\nSTDOUT: ${stdout}`));
      }
    });
  });
}

async function handleMessageCreated(evt) {
  try {
    if (!evt || !evt.data || !evt.data.id) return;

    // Get full message details (files array is only on the full message). :contentReference[oaicite:3]{index=3}
    const msg = await webex.messages.get(evt.data.id);
    // Avoid echo loops
    if (msg.personId === me.id || (msg.personEmail || '').endsWith('@webex.bot')) return;

    // Only handle 1:1 chats for this prototype
    if (msg.roomType !== 'direct') return;

    const roomId = msg.roomId;

    if (!msg.files || msg.files.length === 0) {
      // Small help text
      await sendMarkdown(
        roomId,
        'Hi! Upload a **.pptx** file here and I will reply with a JSON representation of the slide text.'
      );
      return;
    }

    // Process the **first** PPTX file found in the message
    let selectedUrl = null;
    let fileName = null;
    let contentType = null;

    for (const fUrl of msg.files) {
      const { status, headers } = await headContent(fUrl); // check type & filename without full download
      if (status !== 200) continue;
      contentType = headers['content-type'] || '';
      fileName = parseFilenameFromContentDisposition(headers['content-disposition']) || 'upload.pptx';
      const isPptx =
        fileName.toLowerCase().endsWith('.pptx') ||
        contentType.includes('application/vnd.openxmlformats-officedocument.presentationml.presentation');
      if (isPptx) {
        selectedUrl = fUrl;
        break;
      }
    }

    if (!selectedUrl) {
      await sendMarkdown(roomId, 'Sorry — I only process **.pptx** files for now.');
      return;
    }

    // Download file bytes (with anti-malware scan retry). :contentReference[oaicite:4]{index=4}
    const res = await downloadWithRetry(selectedUrl, 'arraybuffer');
    const id = uuidv4();
    const tmpPath = path.join(TMP_DIR, `${id}-${fileName || 'upload.pptx'}`);
    await fsp.writeFile(tmpPath, res.data);

    // Run Python extractor
    const jsonObj = await runPythonExtractor(tmpPath);

    // Send JSON back
    await sendJsonInChunks(roomId, jsonObj, fileName || 'PPTX → JSON');

    // Cleanup
    fsp.unlink(tmpPath).catch(() => {});
  } catch (e) {
    const roomId = evt?.data?.roomId;
    if (roomId) {
      await sendMarkdown(
        roomId,
        `⚠️ Error: ${e.message}\n\nIf this keeps happening, try a different PPTX or check logs.`
      );
    }
    console.error('handleMessageCreated error:', e);
  }
}

(async () => {
  console.log('Initializing Webex SDK...');
  webex.once('ready', async () => {
    try {
      me = await webex.people.get('me');
      console.log(`Authenticated as: ${me.displayName} (${(me.emails || [])[0] || 'no-email'})`);
      // Start websocket listener for messages. :contentReference[oaicite:5]{index=5}
      await webex.messages.listen();
      console.log('Listening for messages over websocket (message:created)...');
      webex.messages.on('created', handleMessageCreated);

      // Nice-to-have: graceful shutdown
      const shutdown = async () => {
        try {
          await webex.messages.stopListening();
        } catch (_) {}
        process.exit(0);
      };
      process.on('SIGINT', shutdown);
      process.on('SIGTERM', shutdown);
    } catch (err) {
      console.error('Startup error:', err.message);
      process.exit(1);
    }
  });
})();
