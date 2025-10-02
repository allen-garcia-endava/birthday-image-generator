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
const cloudinary = require('cloudinary').v2;
const PHOTOS_BASE_URL = process.env.PHOTOS_BASE_URL || '';

const WIDTH = 1366;
const HEIGHT = 768;
const PHOTO_SIZE = 150;

if (process.env.CLOUDINARY_URL) {
  cloudinary.config({ secure: true }); // CLOUDINARY_URL lo resuelve todo
} else if (
  process.env.CLOUDINARY_CLOUD_NAME &&
  process.env.CLOUDINARY_API_KEY &&
  process.env.CLOUDINARY_API_SECRET
) {
  cloudinary.config({
    cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
    api_key: process.env.CLOUDINARY_API_KEY,
    api_secret: process.env.CLOUDINARY_API_SECRET,
    secure: true
  });
}

app.http('generateBirthdayImage', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous', // si vas a llamarla desde Power Automate sin HTTP premium, considera 'function'
  handler: async (request, context) => {
    console.log(`Request1: ${request.method} ${request.url}`);

    if (request.method !== 'POST') {
      return { status: 200, body: 'OK' };
    }

    const ua = request.headers.get('user-agent') || '';
    const ctype = (request.headers.get('content-type') || '').toLowerCase();

    // Lee el body UNA sola vez
    let rawBody = '';
    try {
      rawBody = await request.text();    // <-- se lee aquí y NO se vuelve a leer más abajo
    } catch { }

    // Healthcheck / warm-up (no consumas el body aquí otra vez)
    const isHealthCheck =
      ua.includes('AlwaysOn') ||
      ua.includes('Azure-Functions') ||
      ua.includes('HealthCheck') ||
      ua === '';

    const hasQueryDayMonth = !!(request.query.get('day') && request.query.get('month'));
    const hasBody = !!rawBody;

    if (isHealthCheck && !hasQueryDayMonth && !hasBody) {
      context.log('Health check / warm-up detected, returning 200 without processing');
      return { status: 200, body: 'OK' };
    }

    // --- 1) day/month desde query o body ---
    let day = null, month = null;
    const dayQuery = request.query.get('day');
    const monthQuery = request.query.get('month');

    let photosBaseUrl = PHOTOS_BASE_URL;

    // Parseo del body usando la copia
    let payload = null;
    if (ctype.includes('application/json') && rawBody) {
      try { payload = JSON.parse(rawBody); } catch { }
    } else if (rawBody) {
      // permite "12,9" o "12 9"
      const m = rawBody.match(/(\d{1,2})\D+(\d{1,2})/);
      if (m) payload = { day: m[1], month: m[2] };
    }

    // query > body
    if (dayQuery && monthQuery) {
      day = parseInt(dayQuery, 10);
      month = parseInt(monthQuery, 10);
    } else if (payload?.day != null && payload?.month != null) {
      day = parseInt(payload.day, 10);
      month = parseInt(payload.month, 10);
    }

    context.log("Day/Month from query:", { dayQuery, monthQuery }, "from body:", payload ? { day: payload.day, month: payload.month } : null);

    const hoy = (day && month)
      ? dayjs(`${dayjs().year()}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`)
      : dayjs();

    console.log("Fecha: ", hoy.format('YYYY-MM-DD'), { day, month });

    // --- 2) Origen del CSV: csvText | csvUrl | archivo local ---
    const baseDir = __dirname;
    const csvPathLocal = path.join(baseDir, 'cumples.csv');
    const fotosDir = path.join(baseDir, 'fotos');
    const assetsDir = path.join(baseDir, 'assets');

    let csvText = null;

    // A) csvText en body
    if (payload?.csvText && typeof payload.csvText === 'string') {
      csvText = payload.csvText;
    }

    // B) csvUrl en query o body
    const csvUrl = request.query.get('csvUrl') || payload?.csvUrl || null;
    if (!csvText && csvUrl) {
      try {
        const res = await fetch(csvUrl);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        csvText = await res.text();
      } catch (e) {
        return { status: 400, jsonBody: { error: `No se pudo descargar csvUrl`, detail: String(e?.message || e) } };
      }
    }

    // C) Fallback: archivo local
    if (!csvText) {
      if (!fs.existsSync(csvPathLocal)) {
        return { status: 400, jsonBody: { error: 'Falta csvText/csvUrl y no existe cumples.csv local' } };
      }
      csvText = fs.readFileSync(csvPathLocal, 'utf8');
    }

    // --- 3) Parsear CSV desde string (soporta BOM) ---
    const empleados = await readCsvFromString(csvText);

    // --- 4) Render + guardar en /tmp ---
    const stamp = dayjs().format('YYMMDDHHmmss');
    const outFilePath = path.join('/tmp', `cumples-semana-${stamp}.png`);


    const { cumpleanerosCount } = await renderCumples({
      hoy,
      empleados,
      assetsDir,
      fotosDir: path.join(__dirname, 'fotos'), // fallback local
      photosBaseUrl,                           // NUEVO
      outFilePath
    });

    // --- 5) Subir a Cloudinary (si está configurado) ---
    /*let uploadedUrl = null;
    if (cloudinary.config().cloud_name) {
      const uploadResult = await cloudinary.uploader.upload(outFilePath, {
        folder: 'cumples',
        use_filename: true,
        unique_filename: false,
        overwrite: true,
        resource_type: 'image'
      });
      uploadedUrl = uploadResult.secure_url;
    }*/

    // 5) Subir a Azure Blob Storage (si está configurado)
    const blobName = `cumples-semana-${stamp}.png`;
    const { urlWithSas } = await uploadPngToBlobAndGetUrl(outFilePath, blobName);
    // Usa esta URL para publicar en Teams:
    const uploadedUrl = urlWithSas;

    // 6 subir a TEAMS
    if (process.env.TEAMS_WEBHOOK_URL && uploadedUrl) {
      console.log("Enviando notificación a Teams...");
      const payload = {
        uploadedUrl,
        cumpleaneros: cumpleanerosCount,
        dateRange: {
          start: hoy.format('YYYY-MM-DD'),
          end: hoy.add(6, 'day').format('YYYY-MM-DD')
        }
      };
      await fetch(process.env.TEAMS_WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
    }else{
      console.log("No se envió notificación a Teams (no hay TEAMS_WEBHOOK_URL o no se subió la imagen)");
    }

    return {
      status: 200,
      jsonBody: {
        message: 'Imagen generada correctamente',
        dateRange: { start: hoy.format('YYYY-MM-DD'), end: hoy.add(6, 'day').format('YYYY-MM-DD') },
        cumpleaneros: cumpleanerosCount,
        uploadedUrl,
        source: csvUrl ? 'csvUrl' : (payload?.csvText ? 'csvText' : 'local')
      }
    };
  }
});

async function uploadPngToBlobAndGetUrl(outFilePath, blobName) {
  const accountName = process.env.AZURE_STORAGE_ACCOUNT_NAME;
  const accountKey = process.env.AZURE_STORAGE_ACCOUNT_KEY;
  const containerName = process.env.BIRTHDAY_CONTAINER;

  // Credencial con Account Key (necesaria para firmar SAS)
  const credential = new StorageSharedKeyCredential(accountName, accountKey);
  const service = new BlobServiceClient(
    `https://${accountName}.blob.core.windows.net`,
    credential
  );

  const container = service.getContainerClient(containerName);
  await container.createIfNotExists(); // idempotente

  const blockBlob = container.getBlockBlobClient(blobName);

  // Sube el archivo desde /tmp con content-type adecuado
  await blockBlob.uploadFile(outFilePath, {
    blobHTTPHeaders: {
      blobContentType: 'image/png',
      blobCacheControl: 'public, max-age=31536000' // opcional
    }
  });

  // Genera SAS de solo lectura con expiración
  const ttlHours = parseInt(process.env.BIRTHDAY_SAS_TTL_HOURS || '336', 10); // 14 días
  const startsOn = new Date(Date.now() - 5 * 60 * 1000);     // 5 min atrás por relojes
  const expiresOn = new Date(Date.now() + ttlHours * 3600 * 1000);

  const sas = generateBlobSASQueryParameters({
    containerName,
    blobName,
    permissions: BlobSASPermissions.parse('r'),
    startsOn,
    expiresOn
  }, credential).toString();

  const urlWithSas = `${blockBlob.url}?${sas}`;
  return { url: blockBlob.url, urlWithSas };
}


// Helper para construir la URL de la foto
function resolvePhotoSrc(photoName, { photosBaseUrl, fotosDir }) {
  if (!photoName || !photoName.trim()) return null;
  const name = photoName.trim();

  console.log("Photo Name: ", name);
  // 1) Si ya es URL absoluta
  if (/^https?:\/\//i.test(name)) return name;

  // 2) Si hay base URL, construir (admite SAS si la base ya trae ?sv=... etc.)
  if (photosBaseUrl) {
    // cuida dobles barras
    return photosBaseUrl.replace(/\/+$/, '') + '/' + name.replace(/^\/+/, '');
  }

  // 3) Fallback local (archivo)
  const p = path.join(fotosDir, name);
  return fs.existsSync(p) ? p : null;
}


// === Nuevo: parsear CSV desde string ===
const { Readable } = require('stream');
function readCsvFromString(text) {
  // Quitar BOM si lo trae
  if (text && text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  // También por si viene codificado como \uFEFF
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

async function renderCumples({ hoy, empleados, assetsDir, fotosDir, photosBaseUrl, outFilePath }) {
  console.log('About to require canvas…');
  const { createCanvas, loadImage, GlobalFonts } = loadCanvasBinding();

  console.log('Canvas functions:', { createCanvas, loadImage });
  const diasSemana = Array.from({ length: 7 }, (_, i) => hoy.add(i, 'day'));

  const cumpleaneros = empleados.filter((emp) => {
    const dia = parseInt(emp.Day, 10);
    const mes = parseInt(emp.Month, 10);
    return diasSemana.some((d) => d.date() === dia && d.month() + 1 === mes);
  });

  // 1) Registrar fuentes empaquetadas
  const fontsDir = path.join(assetsDir, 'fonts'); // crea esta carpeta y agrega tus .ttf
  try {
    // Registra al menos regular y bold de la MISMA familia
    GlobalFonts.registerFromPath(path.join(fontsDir, 'Inter-Regular.ttf')); // cambia por la tuya
    GlobalFonts.registerFromPath(path.join(fontsDir, 'Inter-Bold.ttf'));
    console.log('Fonts registered from', fontsDir);
  } catch (e) {
    console.log('Font registration failed (will fallback):', e?.message || e);
  }

  console.log(`Found ${cumpleaneros.length} cumpleaneros between ${hoy.format('DD/MM')} and ${hoy.add(6, 'day').format('DD/MM')}`);
  const canvas = createCanvas(WIDTH, HEIGHT);

  const ctx = canvas.getContext('2d');

  // Fondo
  const background = await loadImage(path.join(assetsDir, 'images/background.png'));
  ctx.drawImage(background, 0, 0, WIDTH, HEIGHT);

  const maxPorFila = 4;
  const filaYBase = 320;
  const filaAltura = PHOTO_SIZE + 100;

  // Badge de fecha
  const fechaInicio = hoy.format('MMMM DD');
  const fechaFin = hoy.add(6, 'day').format('MMMM DD');
  const textoRango = `${fechaInicio} to ${fechaFin}`;

  ctx.save();
  ctx.translate(1100, 130);
  ctx.rotate(-0.1);
  ctx.fillStyle = '#FF5C45';
  ctx.font = 'bold 24px Inter';
  const textWidth = ctx.measureText(textoRango).width;
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
  ctx.fillText(textoRango, -textWidth / 2, 8);
  ctx.restore();

  const burbuja = await loadImage(path.join(assetsDir, 'images/day-bubble.png'));

  for (let i = 0; i < cumpleaneros.length; i++) {
    const emp = cumpleaneros[i];
    const fila = Math.floor(i / maxPorFila);
    const columna = i % maxPorFila;

    let filaBase = cumpleaneros.length <= 4 ? 400 : filaYBase;
    const personasEnFila = cumpleaneros.slice(fila * maxPorFila, (fila + 1) * maxPorFila);
    const spacingFila = WIDTH / (personasEnFila.length + 1);
    const x = spacingFila * (columna + 1);
    const y = filaBase + fila * filaAltura;

    // === Foto ===
    const rawName = (emp.PhotoName && String(emp.PhotoName)) || '';

    console.log('Apunto de resolver foto para', emp.FirstName, emp.LastName, '->', rawName);
    let src = resolvePhotoSrc(rawName, { photosBaseUrl, fotosDir });

    // fallback a user.png (URL o local)
    if (!src) {
      const defaultName = 'user.png';
      src = resolvePhotoSrc(defaultName, { photosBaseUrl, fotosDir });
    }

    try {
      // loadImage acepta rutas locales o HTTP(S). Para blobs privados, usa SAS.
      const img = await loadImage(src);
      ctx.save();
      ctx.beginPath();
      ctx.arc(x, y, PHOTO_SIZE / 2, 0, Math.PI * 2);
      ctx.closePath();
      ctx.clip();
      ctx.drawImage(img, x - PHOTO_SIZE / 2, y - PHOTO_SIZE / 2, PHOTO_SIZE, PHOTO_SIZE);
      ctx.restore();
    } catch (e) {
      // si una foto falla, continúa sin bloquear todo
      console.log('Foto no cargó:', src, e?.message || e);
    }


    // Burbuja día
    const bubbleWidth = 80;
    const bubbleHeight = 35;
    const bubbleX = x - bubbleWidth / 2 + 80;
    const bubbleY = y - PHOTO_SIZE / 2 - 10;
    ctx.drawImage(burbuja, bubbleX, bubbleY, bubbleWidth, bubbleHeight);

    ctx.font = 'bold 20px Inter';
    ctx.fillStyle = 'white';
    const diaTexto = String(emp.Day).padStart(2, '0');
    const anchoDia = ctx.measureText(diaTexto).width;
    ctx.fillText(diaTexto, x - anchoDia / 2 + 84, bubbleY + bubbleHeight / 2 + 2);

    // Nombre
    ctx.font = 'bold 28px Inter';
    ctx.fillStyle = 'white';
    const nombre = `${emp.FirstName?.trim() ?? ''} ${emp.LastName?.trim() ?? ''}`.trim();
    const textWidthNombre = ctx.measureText(nombre).width;
    ctx.fillText(nombre, x - textWidthNombre / 2, y + PHOTO_SIZE / 2 + 40);
  }

  console.log("Rendering done, saving to", outFilePath);

  // Guardar PNG en disco (usar /tmp en Azure)
  const buffer = canvas.toBuffer('image/png');
  fs.writeFileSync(outFilePath, buffer);

  return { cumpleanerosCount: cumpleaneros.length, outFilePath };
}

function loadCanvasBinding() {
  try {
    return require('@napi-rs/canvas'); // normalmente suficiente
  } catch {
    return require('@napi-rs/canvas-linux-x64-gnu'); // fallback explícito para glibc
  }
}

app.http('testCanvas', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async () => {
    const { createCanvas } = loadCanvasBinding();
    const c = createCanvas(10, 10);
    return { status: 200, body: `canvas OK ${c.width}x${c.height}` };
  }
});
