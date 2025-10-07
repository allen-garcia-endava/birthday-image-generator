const { app } = require('@azure/functions');
const {
  BlobServiceClient,
  StorageSharedKeyCredential,
  BlobSASPermissions,
  generateBlobSASQueryParameters
} = require('@azure/storage-blob');

const fs = require('fs');
const path = require('path');
const csv = require('csv-parser');
const dayjs = require('dayjs');

const PHOTOS_BASE_URL = process.env.PHOTOS_BASE_URL || '';
const WIDTH = 1366;
const HEIGHT = 768;
const PHOTO_SIZE = 150;

app.http('generateBirthdayImage', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    // -- Intention gating: only run on POST (GET behaves like health/peek)
    if (request.method !== 'POST') {
      return { status: 200, body: 'OK' };
    }

    const ua = request.headers.get('user-agent') || '';
    const ctype = (request.headers.get('content-type') || '').toLowerCase();

    // Read body ONCE (never call request.json() or request.text() later again)
    let rawBody = '';
    try { rawBody = await request.text(); } catch {}

    // -- Skip warm-ups/health checks without meaningful input
    const isHealthCheck =
      ua.includes('AlwaysOn') ||
      ua.includes('Azure-Functions') ||
      ua.includes('HealthCheck') ||
      ua === '';

    const hasQueryDayMonth = !!(request.query.get('day') && request.query.get('month'));
    const hasBody = !!rawBody;

    if (isHealthCheck && !hasQueryDayMonth && !hasBody) {
      return { status: 200, body: 'OK' };
    }

    // -- Parse minimal inputs (day/month) from query or body
    let day = null, month = null;
    const dayQuery = request.query.get('day');
    const monthQuery = request.query.get('month');

    let body = null;
    if (ctype.includes('application/json') && rawBody) {
      try { body = JSON.parse(rawBody); } catch {}
    } else if (rawBody) {
      // Allow "12,9" or "12 9"
      const m = rawBody.match(/(\d{1,2})\D+(\d{1,2})/);
      if (m) body = { day: m[1], month: m[2] };
    }

    if (dayQuery && monthQuery) {
      day = parseInt(dayQuery, 10);
      month = parseInt(monthQuery, 10);
    } else if (body?.day != null && body?.month != null) {
      day = parseInt(body.day, 10);
      month = parseInt(body.month, 10);
    }

    const today = (day && month)
      ? dayjs(`${dayjs().year()}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`)
      : dayjs();

    // -- CSV source: prefer body.csvText, then csvUrl, then local fallback
    const baseDir = __dirname;
    const csvPathLocal = path.join(baseDir, 'cumples.csv');
    const fotosDir = path.join(baseDir, 'fotos');
    const assetsDir = path.join(baseDir, 'assets');

    let csvText = null;

    if (body?.csvText && typeof body.csvText === 'string') {
      csvText = body.csvText;
    }

    const csvUrl = request.query.get('csvUrl') || body?.csvUrl || null;
    if (!csvText && csvUrl) {
      try {
        const res = await fetch(csvUrl);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        csvText = await res.text();
      } catch (e) {
        return { status: 400, jsonBody: { error: 'Failed to download csvUrl', detail: String(e?.message || e) } };
      }
    }

    if (!csvText) {
      if (!fs.existsSync(csvPathLocal)) {
        return { status: 400, jsonBody: { error: 'Missing csvText/csvUrl and local cumples.csv not found' } };
      }
      csvText = fs.readFileSync(csvPathLocal, 'utf8');
    }

    // -- Parse CSV from string (BOM-safe)
    const employees = await readCsvFromString(csvText);

    // -- Render to /tmp
    const stamp = dayjs().format('YYMMDDHHmmss');
    const outFilePath = path.join('/tmp', `birthday-week-${stamp}.png`);

    const { cumpleanerosCount } = await renderCumples({
      hoy: today,
      empleados: employees,
      assetsDir,
      fotosDir,                 // local fallback
      photosBaseUrl: PHOTOS_BASE_URL,
      outFilePath
    });

    // -- Upload to Azure Blob (and generate SAS URL)
    const blobName = `birthday-week-${stamp}.png`;
    const { urlWithSas } = await uploadPngToBlobAndGetUrl(outFilePath, blobName);
    const uploadedUrl = urlWithSas;

    // -- Notify Teams (optional)
    if (process.env.TEAMS_WEBHOOK_URL && uploadedUrl) {
      const teamsPayload = {
        uploadedUrl,
        cumpleaneros: cumpleanerosCount,
        dateRange: {
          start: today.format('YYYY-MM-DD'),
          end: today.add(6, 'day').format('YYYY-MM-DD')
        }
      };
      try {
        await fetch(process.env.TEAMS_WEBHOOK_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(teamsPayload)
        });
      } catch (e) {
        // Log only, do not fail the request on Teams error
        context.log('Teams webhook failed:', e?.message || e);
      }
    }

    return {
      status: 200,
      jsonBody: {
        message: 'Image generated successfully',
        dateRange: { start: today.format('YYYY-MM-DD'), end: today.add(6, 'day').format('YYYY-MM-DD') },
        cumpleaneros: cumpleanerosCount,
        uploadedUrl,
        source: csvUrl ? 'csvUrl' : (body?.csvText ? 'csvText' : 'local')
      }
    };
  }
});

/** Uploads the PNG file to Blob Storage and returns a read-only SAS URL. */
async function uploadPngToBlobAndGetUrl(outFilePath, blobName) {
  const accountName = process.env.AZURE_STORAGE_ACCOUNT_NAME;
  const accountKey  = process.env.AZURE_STORAGE_ACCOUNT_KEY;
  const containerName = process.env.BIRTHDAY_CONTAINER;

  if (!accountName || !accountKey || !containerName) {
    throw new Error('Missing storage settings: AZURE_STORAGE_ACCOUNT_NAME / AZURE_STORAGE_ACCOUNT_KEY / BIRTHDAY_CONTAINER');
  }

  const credential = new StorageSharedKeyCredential(accountName, accountKey);
  const service = new BlobServiceClient(`https://${accountName}.blob.core.windows.net`, credential);

  const container = service.getContainerClient(containerName);
  await container.createIfNotExists();

  const blockBlob = container.getBlockBlobClient(blobName);

  await blockBlob.uploadFile(outFilePath, {
    blobHTTPHeaders: {
      blobContentType: 'image/png',
      blobCacheControl: 'public, max-age=31536000'
    }
  });

  const ttlHours = parseInt(process.env.BIRTHDAY_SAS_TTL_HOURS || '336', 10); // default: 14 days
  const startsOn = new Date(Date.now() - 5 * 60 * 1000); // tolerate small clock skews
  const expiresOn = new Date(Date.now() + ttlHours * 3600 * 1000);

  const sas = generateBlobSASQueryParameters({
    containerName,
    blobName,
    permissions: BlobSASPermissions.parse('r'),
    startsOn,
    expiresOn
  }, credential).toString();

  return { url: blockBlob.url, urlWithSas: `${blockBlob.url}?${sas}` };
}

/** Resolves an employee photo source from URL base or local fallback. */
function resolvePhotoSrc(photoName, { photosBaseUrl, fotosDir }) {
  if (!photoName || !photoName.trim()) return null;
  const name = photoName.trim();

  // Absolute URL already?
  if (/^https?:\/\//i.test(name)) return name;

  // Build URL from base (SAS-ready)
  if (photosBaseUrl) {
    return photosBaseUrl.replace(/\/+$/, '') + '/' + name.replace(/^\/+/, '');
  }

  // Local fallback
  const p = path.join(fotosDir, name);
  return fs.existsSync(p) ? p : null;
}

const { Readable } = require('stream');
/** Parses CSV text to objects, handling BOM if present. */
function readCsvFromString(text) {
  if (text && text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  text = text.replace(/^\uFEFF/, '');
  return new Promise((resolve, reject) => {
    const rows = [];
    Readable.from(text)
      .pipe(csv())
      .on('data', (r) => rows.push(r))
      .on('end', () => resolve(rows))
      .on('error', (err) => reject(err));
  });
}

/** Renders the weekly birthday image to /tmp using @napi-rs/canvas. */
async function renderCumples({ hoy, empleados, assetsDir, fotosDir, photosBaseUrl, outFilePath }) {
  const { createCanvas, loadImage, GlobalFonts } = loadCanvasBinding();

  const weekDays = Array.from({ length: 7 }, (_, i) => hoy.add(i, 'day'));
  const cumpleaneros = empleados.filter((emp) => {
    const d = parseInt(emp.Day, 10);
    const m = parseInt(emp.Month, 10);
    return weekDays.some((wd) => wd.date() === d && wd.month() + 1 === m);
  });

  // Register packaged fonts (ensure these .ttf files exist in assets/fonts)
  const fontsDir = path.join(assetsDir, 'fonts');
  try {
    GlobalFonts.registerFromPath(path.join(fontsDir, 'Inter-Regular.ttf'));
    GlobalFonts.registerFromPath(path.join(fontsDir, 'Inter-Bold.ttf'));
  } catch {}

  const canvas = createCanvas(WIDTH, HEIGHT);
  const ctx = canvas.getContext('2d');

  // Background
  const background = await loadImage(path.join(assetsDir, 'images/background.png'));
  ctx.drawImage(background, 0, 0, WIDTH, HEIGHT);

  // Date badge
  const startLabel = hoy.format('MMMM DD');
  const endLabel = hoy.add(6, 'day').format('MMMM DD');
  const rangeText = `${startLabel} to ${endLabel}`;

  ctx.save();
  ctx.translate(1100, 130);
  ctx.rotate(-0.1);
  ctx.fillStyle = '#FF5C45';
  ctx.font = 'bold 24px Inter';
  const textWidth = ctx.measureText(rangeText).width;
  const badgeWidth = textWidth + 40;
  const badgeHeight = 40;
  const radius = 20;

  ctx.beginPath();
  ctx.moveTo(-badgeWidth / 2 + radius, -badgeHeight / 2);
  ctx.lineTo(badgeWidth / 2 - radius, -badgeHeight / 2);
  ctx.quadraticCurveTo(badgeWidth / 2, -badgeHeight / 2, badgeWidth / 2, -badgeHeight / 2 + radius);
  ctx.lineTo(badgeWidth / 2, badgeHeight / 2 - radius);
  ctx.quadraticCurveTo(badgeWidth / 2, badgeHeight / 2, badgeWidth / 2 - radius, badgeHeight / 2);
  ctx.lineTo(-badgeWidth / 2 + radius, badgeHeight / 2);
  ctx.quadraticCurveTo(-badgeWidth / 2, badgeHeight / 2, -badgeWidth / 2, badgeHeight / 2 - radius);
  ctx.lineTo(-badgeWidth / 2, -badgeHeight / 2 + radius);
  ctx.quadraticCurveTo(-badgeWidth / 2, -badgeHeight / 2, -badgeWidth / 2 + radius, -badgeHeight / 2);
  ctx.closePath();
  ctx.fill();

  ctx.fillStyle = 'white';
  ctx.fillText(rangeText, -textWidth / 2, 8);
  ctx.restore();

  // People grid
  const bubble = await loadImage(path.join(assetsDir, 'images/day-bubble.png'));
  const maxPerRow = 4;
  const baseY = 320;
  const rowHeight = PHOTO_SIZE + 100;

  for (let i = 0; i < cumpleaneros.length; i++) {
    const emp = cumpleaneros[i];
    const row = Math.floor(i / maxPerRow);
    const col = i % maxPerRow;

    const few = cumpleaneros.length <= 4;
    const rowBase = few ? 400 : baseY;
    const peopleInRow = cumpleaneros.slice(row * maxPerRow, (row + 1) * maxPerRow);
    const spacingX = WIDTH / (peopleInRow.length + 1);
    const x = spacingX * (col + 1);
    const y = rowBase + row * rowHeight;

    // Photo
    const rawName = (emp.PhotoName && String(emp.PhotoName)) || '';
    let src = resolvePhotoSrc(rawName, { photosBaseUrl, fotosDir });
    if (!src) src = resolvePhotoSrc('user.png', { photosBaseUrl, fotosDir });

    try {
      const img = await loadImage(src);
      ctx.save();
      ctx.beginPath();
      ctx.arc(x, y, PHOTO_SIZE / 2, 0, Math.PI * 2);
      ctx.closePath();
      ctx.clip();
      ctx.drawImage(img, x - PHOTO_SIZE / 2, y - PHOTO_SIZE / 2, PHOTO_SIZE, PHOTO_SIZE);
      ctx.restore();
    } catch {}

    // Day bubble
    const bubbleWidth = 80;
    const bubbleHeight = 35;
    const bubbleX = x - bubbleWidth / 2 + 80;
    const bubbleY = y - PHOTO_SIZE / 2 - 10;
    ctx.drawImage(bubble, bubbleX, bubbleY, bubbleWidth, bubbleHeight);

    ctx.font = 'bold 20px Inter';
    ctx.fillStyle = 'white';
    const dayText = String(emp.Day).padStart(2, '0');
    const dayWidth = ctx.measureText(dayText).width;
    ctx.fillText(dayText, x - dayWidth / 2 + 84, bubbleY + bubbleHeight / 2 + 2);

    // Name
    ctx.font = 'bold 28px Inter';
    ctx.fillStyle = 'white';
    const name = `${emp.FirstName?.trim() ?? ''} ${emp.LastName?.trim() ?? ''}`.trim();
    const nameWidth = ctx.measureText(name).width;
    ctx.fillText(name, x - nameWidth / 2, y + PHOTO_SIZE / 2 + 40);
  }

  // Save PNG to /tmp
  const buffer = canvas.toBuffer('image/png');
  fs.writeFileSync(outFilePath, buffer);

  return { cumpleanerosCount: cumpleaneros.length, outFilePath };
}

/** Loads canvas binding for Linux/glibc environments. */
function loadCanvasBinding() {
  try {
    return require('@napi-rs/canvas');
  } catch {
    return require('@napi-rs/canvas-linux-x64-gnu');
  }
}

// Lightweight probe endpoint (optional)
app.http('testCanvas', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async () => {
    const { createCanvas } = loadCanvasBinding();
    const c = createCanvas(10, 10);
    return { status: 200, body: `canvas OK ${c.width}x${c.height}` };
  }
});
