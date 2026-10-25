// PPTX Remote — present a .pptx fullscreen on the Mac, control it from a phone.
// Zero dependencies. Slides are rendered by Keynote (AppleScript export to PNG).
const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { execFile } = require('child_process');

const PORT = Number(process.env.PORT) || 3000;
const ROOT = __dirname;
const DECKS = path.join(ROOT, 'decks');
const PUBLIC = path.join(ROOT, 'public');
fs.mkdirSync(DECKS, { recursive: true });

// Shared presentation state, broadcast to every connected screen.
let state = { deckId: null, name: null, slides: [], notes: [], index: 0, status: 'idle', error: null };
const clients = new Set();

function broadcast() {
  const msg = `data: ${JSON.stringify(state)}\n\n`;
  for (const res of clients) res.write(msg);
}

function lanAddress() {
  for (const list of Object.values(os.networkInterfaces())) {
    for (const a of list || []) if (a.family === 'IPv4' && !a.internal) return a.address;
  }
  return 'localhost';
}

function exportWithKeynote(pptxPath, outDir) {
  return new Promise((resolve, reject) => {
    execFile('osascript', [path.join(ROOT, 'export.applescript'), pptxPath, outDir],
      { timeout: 200 * 1000, maxBuffer: 50 * 1024 * 1024 },
      (err, stdout, stderr) => {
        if (err) return reject(new Error(`${(stderr || err.message).trim()} (if this keeps happening, quit and reopen Keynote)`));
        resolve(stdout.replace(/\n$/, ''));
      });
  });
}

// Keynote names exports like "Deck.001.png" — sort by the number in the name.
function slideFiles(dir) {
  const files = [];
  (function walk(d) {
    for (const f of fs.readdirSync(d, { withFileTypes: true })) {
      const p = path.join(d, f.name);
      if (f.isDirectory()) walk(p);
      else if (/\.png$/i.test(f.name)) files.push(p);
    }
  })(dir);
  const num = (p) => Number((path.basename(p).match(/(\d+)\.png$/i) || [0, 0])[1]);
  return files.sort((a, b) => num(a) - num(b));
}

async function handleUpload(req, res) {
  const name = decodeURIComponent(req.headers['x-filename'] || 'deck.pptx').replace(/[/\\]/g, '_');
  if (!/\.(pptx?|key)$/i.test(name)) return send(res, 400, { error: 'Please upload a .pptx file' });

  const id = crypto.randomBytes(6).toString('hex');
  const dir = path.join(DECKS, id);
  const imgDir = path.join(dir, 'slides');
  fs.mkdirSync(imgDir, { recursive: true });
  const file = path.join(dir, name);

  await new Promise((resolve, reject) => {
    const out = fs.createWriteStream(file);
    req.pipe(out);
    out.on('finish', resolve);
    out.on('error', reject);
  });

  state = { ...state, status: 'converting', name, error: null };
  broadcast();
  send(res, 202, { ok: true });

  try {
    const notesRaw = await exportWithKeynote(file, imgDir);
    const files = slideFiles(imgDir);
    if (!files.length) throw new Error('Keynote exported no slides');
    const notes = notesRaw.split('\u001e');
    state = {
      deckId: id, name,
      slides: files.map((f) => `/decks/${id}/${path.relative(dir, f).split(path.sep).map(encodeURIComponent).join('/')}`),
      notes: files.map((_, i) => (notes[i] || '').trim()),
      index: 0, status: 'ready', error: null,
    };
  } catch (e) {
    console.error(e);
    state = { ...state, status: state.slides.length ? 'ready' : 'idle', error: `Conversion failed: ${e.message}` };
  }
  broadcast();
}

function send(res, code, obj) {
  res.writeHead(code, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify(obj));
}

function readJson(req) {
  return new Promise((resolve) => {
    let body = '';
    req.on('data', (c) => (body += c));
    req.on('end', () => { try { resolve(JSON.parse(body || '{}')); } catch { resolve({}); } });
  });
}

const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.svg': 'image/svg+xml' };

function serveFile(res, base, rel) {
  const p = path.normalize(path.join(base, rel));
  if (!p.startsWith(base)) return send(res, 403, { error: 'forbidden' });
  fs.stat(p, (err, st) => {
    if (err || !st.isFile()) return send(res, 404, { error: 'not found' });
    res.writeHead(200, { 'Content-Type': TYPES[path.extname(p).toLowerCase()] || 'application/octet-stream', 'Cache-Control': 'no-cache' });
    fs.createReadStream(p).pipe(res);
  });
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://x');
  const p = decodeURIComponent(url.pathname);

  if (p === '/events') {
    res.writeHead(200, { 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-cache', Connection: 'keep-alive' });
    res.write(`data: ${JSON.stringify(state)}\n\n`);
    clients.add(res);
    const ping = setInterval(() => res.write(': ping\n\n'), 20000);
    req.on('close', () => { clearInterval(ping); clients.delete(res); });
    return;
  }
  if (p === '/api/info') return send(res, 200, { remoteUrl: `http://${lanAddress()}:${PORT}/remote` });
  if (p === '/api/upload' && req.method === 'POST') return handleUpload(req, res);
  if (p === '/api/goto' && req.method === 'POST') {
    const { index, delta } = await readJson(req);
    if (!state.slides.length) return send(res, 409, { error: 'no deck loaded' });
    const next = typeof index === 'number' ? index : state.index + (Number(delta) || 0);
    state.index = Math.max(0, Math.min(state.slides.length - 1, next));
    broadcast();
    return send(res, 200, { index: state.index });
  }
  if (p.startsWith('/decks/')) return serveFile(res, DECKS, p.slice('/decks/'.length));
  if (p === '/' ) return serveFile(res, PUBLIC, 'present.html');
  if (p === '/remote') return serveFile(res, PUBLIC, 'remote.html');
  return serveFile(res, PUBLIC, p.slice(1));
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Presenter (open on this Mac):  http://localhost:${PORT}/`);
  console.log(`Remote (open on your phone):   http://${lanAddress()}:${PORT}/remote`);
});
