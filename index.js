'use strict';
// Modified from ep_image_insert 1.0.7 
const eejs = require('ep_etherpad-lite/node/eejs/');
// Compat: Etherpad 2.4 uses ESM for Settings. Support both CJS and ESM.
const settingsModule = require('ep_etherpad-lite/node/utils/Settings');
const settings = settingsModule.default || settingsModule;
const { randomUUID } = require('crypto');
const path = require('path');
const url = require('url');
const fs = require('fs');
const fsp = fs.promises;
const { JSDOM } = require('jsdom');
const { exec } = require('child_process');
const util = require('util');
const execPromise = util.promisify(exec);
const mimetypes = require('mime-db');
// AWS SDK v3 for presigned URLs
let S3Client, PutObjectCommand, getSignedUrl;
try {
  ({ S3Client, PutObjectCommand } = require('@aws-sdk/client-s3'));
  ({ getSignedUrl } = require('@aws-sdk/s3-request-presigner'));
} catch (e) {
  // AWS SDK might be optional if s3_presigned storage is not used
  console.warn('[ep_images_extended] AWS SDK not installed; s3_presigned storage will not work.');
}

// Replaced log4js with lightweight console wrapper to avoid external dependency
// while preserving same API surface used below.
const logger = {
  debug: console.debug.bind(console),
  info: console.info.bind(console),
  warn: console.warn.bind(console),
  error: console.error.bind(console),
};

// Simple in-memory IP rate limiter
const _presignRateStore = new Map();
const PRESIGN_RATE_WINDOW_MS = 60 * 1000;   // 1 minute
const PRESIGN_RATE_MAX = 30;                // max 30 presigns per IP per min

// Utility: basic per-IP sliding-window rate limit
const _rateLimitCheck = (ip) => {
  const now = Date.now();
  let stamps = _presignRateStore.get(ip) || [];
  stamps = stamps.filter((t) => t > now - PRESIGN_RATE_WINDOW_MS);
  if (stamps.length >= PRESIGN_RATE_MAX) return false;
  stamps.push(now);
  _presignRateStore.set(ip, stamps);
  return true;
};

/**
 * ClientVars hook
 *
 * Exposes plugin settings from settings.json to client code inside clientVars variable
 * to be accessed from client side hooks
 *
 * @param {string} hookName Hook name ("clientVars").
 * @param {object} args Object containing the arguments passed to hook. {pad: {object}}
 * @param {function} cb Callback
 *
 * @returns {*} callback
 *
 * @see {@link http://etherpad.org/doc/v1.5.7/#index_clientvars}
 */
exports.clientVars = (hookName, args, cb) => {
  const pluginSettings = {
    storageType: 'local',
  };
  if (!settings.ep_images_extended) {
    settings.ep_images_extended = {};
  }
  const keys = Object.keys(settings.ep_images_extended);
  keys.forEach((key) => {
    if (key !== 'storage') {
      pluginSettings[key] = settings.ep_images_extended[key];
    }
  });
  if (settings.ep_images_extended.storage)
  {
    pluginSettings.storageType = settings.ep_images_extended.storage.type;
  }

  pluginSettings.mimeTypes = mimetypes;

  return cb({ep_images_extended: pluginSettings});
};

exports.eejsBlock_styles = (hookName, args, cb) => {
  args.content += "<link href='../static/plugins/ep_images_extended/static/css/ace.css' rel='stylesheet'>";
  return cb();
};

exports.eejsBlock_timesliderStyles = (hookName, args, cb) => {
  args.content += "<link href='../../static/plugins/ep_images_extended/static/css/ace.css' rel='stylesheet'>";
  args.content += '<style>.control-container{display:none}</style>';
  return cb();
};

exports.eejsBlock_body = (hookName, args, cb) => {
  const modal = eejs.require('ep_images_extended/templates/modal.ejs');
  const imageFormatMenu = eejs.require('ep_images_extended/templates/imageFormatMenu.ejs');
  args.content += modal;
  args.content += imageFormatMenu;

  return cb();
};

exports.expressConfigure = (hookName, context) => {
  /* ------------------------------------------------------------------
   * New endpoint: GET /p/:padId/pluginfw/ep_images_extended/s3_presign
   * ------------------------------------------------------------------
   * Returns: { signedUrl: string, publicUrl: string }
   * Register the route only when storage.type === 's3_presigned'
   */
  if (settings.ep_images_extended && settings.ep_images_extended.storage && settings.ep_images_extended.storage.type === 's3_presigned') {
    context.app.get('/p/:padId/pluginfw/ep_images_extended/s3_presign', async (req, res) => {
      /* ------------------ Basic auth check ------------------ */
      const hasExpressSession = req.session && (req.session.user || req.session.authorId);
      const hasPadCookie = req.cookies && (req.cookies.sessionID || req.cookies.token);
      if (!hasExpressSession && !hasPadCookie) {
        return res.status(401).json({ error: 'Authentication required' });
      }

      /* ------------------ Rate limiting --------------------- */
      const ip = req.ip || req.headers['x-forwarded-for'] || req.connection?.remoteAddress || 'unknown';
      if (!_rateLimitCheck(ip)) {
        return res.status(429).json({ error: 'Too many presign requests' });
      }

      try {
        const storageCfg = settings.ep_images_extended && settings.ep_images_extended.storage;
        if (!storageCfg || storageCfg.type !== 's3_presigned') {
          return res.status(400).json({ error: 's3_presigned storage not enabled' });
        }

        if (!S3Client || !PutObjectCommand || !getSignedUrl) {
          return res.status(500).json({ error: 'AWS SDK not available on server' });
        }

        const { bucket, region, publicURL, expires } = storageCfg;
        if (!bucket || !region) {
          return res.status(500).json({ error: 'Invalid S3 configuration' });
        }

        const { padId } = req.params;
        const { name, type } = req.query;
        if (!name || !type) {
          return res.status(400).json({ error: 'Missing name or type' });
        }

        /* ------------- MIME / extension allow-list ------------ */
        if (settings.ep_images_extended && settings.ep_images_extended.fileTypes && Array.isArray(settings.ep_images_extended.fileTypes)) {
          const allowedExts = settings.ep_images_extended.fileTypes;
          const extName = path.extname(name).replace('.', '').toLowerCase();
          if (!allowedExts.includes(extName)) {
            return res.status(400).json({ error: 'File type not allowed' });
          }
        }

        const ext = path.extname(name);
        // Ensure ext starts with '.'; if not, prefix it
        const safeExt = ext.startsWith('.') ? ext : `.${ext}`;
        const key = `${padId}/${randomUUID()}${safeExt}`;

        const s3Client = new S3Client({ region }); // credentials from env / IAM role

        const putCommand = new PutObjectCommand({
          Bucket: bucket,
          Key: key,
          ContentType: type,
        });

        const signedUrl = await getSignedUrl(s3Client, putCommand, { expiresIn: expires || 600 });

        const basePublic = publicURL || `https://${bucket}.s3.${region}.amazonaws.com/`;
        const publicUrl = new url.URL(key, basePublic).toString();

        return res.json({ signedUrl, publicUrl });
      } catch (err) {
        logger.error('[ep_images_extended] S3 presign error', err);
        return res.status(500).json({ error: 'Failed to generate presigned URL' });
      }
    });
  }

  // ADD LOCAL DISK STORAGE UPLOAD ENDPOINT ------------------------------
  // Register the route only if storage.type === 'local'
  logger.info('[ep_images_extended] storageType at startup:',
    settings.ep_images_extended?.storage?.type);
    
  if (settings.ep_images_extended && settings.ep_images_extended.storage && settings.ep_images_extended.storage.type === 'local') {
    // Route: POST /p/:padId/pluginfw/ep_images_extended/upload
    // Accepts multipart/form-data with field "file" and saves it to the
    // configured baseFolder. Responds with the public URL of the uploaded file.
    context.app.post('/p/:padId/pluginfw/ep_images_extended/upload', async (req, res) => {
      /* ------------------ Basic auth check ------------------ */
      const hasExpressSession = req.session && (req.session.user || req.session.authorId);
      const hasPadCookie = req.cookies && (req.cookies.sessionID || req.cookies.token);
      if (!hasExpressSession && !hasPadCookie) {
        return res.status(401).json({ error: 'Authentication required' });
      }

      /* ------------------ Rate limiting --------------------- */
      const ip = req.ip || req.headers['x-forwarded-for'] || req.connection?.remoteAddress || 'unknown';
      if (!_rateLimitCheck(ip)) {
        return res.status(429).json({ error: 'Too many uploads' });
      }

      try {
        // Dynamically require formidable only when needed
        const formidableMod = require('formidable');
        const IncomingForm = formidableMod.IncomingForm || formidableMod; // support both v1 and v2+ exports
        const form = new IncomingForm({ multiples: false, maxFileSize: settings.ep_images_extended.maxFileSize || 1024 * 1024 * 20 /* 20 MB default */ });

        form.parse(req, async (err, _fields, files) => {
          if (err) {
            logger.error('[ep_images_extended] formidable parse error', err);
            return res.status(400).json({ error: 'Invalid form data' });
          }
          if (!files.file) {
            return res.status(400).json({ error: 'No file provided' });
          }
          const uploaded = Array.isArray(files.file) ? files.file[0] : files.file;
          const mimeType = uploaded.mimetype || uploaded.type || 'application/octet-stream';

          // Reject non-image MIME types
          if (!mimeType.startsWith('image/')) {
            return res.status(400).json({ error: 'Not an image MIME type' });
          }

          // Enforce fileTypes allow-list if configured
          if (settings.ep_images_extended && Array.isArray(settings.ep_images_extended.fileTypes)) {
            const allowedExts = settings.ep_images_extended.fileTypes;
            const extName = path.extname(uploaded.originalFilename || uploaded.name).replace('.', '').toLowerCase();
            if (!allowedExts.includes(extName)) {
              return res.status(400).json({ error: 'File type not allowed' });
            }
          }

          const { padId } = req.params;
          const safePad = path.basename(padId); // prevent path traversal
          const baseFolder = settings.ep_images_extended.storage.baseFolder || path.join(settings.root || process.cwd(), 'src/static/images');
          const destFolder = path.resolve(baseFolder, safePad);
          await fsp.mkdir(destFolder, { recursive: true });

          const newFilename = `${randomUUID()}${path.extname(uploaded.originalFilename || uploaded.name)}`;
          const destPath = path.join(destFolder, newFilename);

          try {
            await fsp.rename(uploaded.filepath || uploaded.path, destPath);
          } catch (errMove) {
            if (errMove.code === 'EXDEV') {
              // Cross-device move: fallback to copy & unlink
              await fsp.copyFile(uploaded.filepath || uploaded.path, destPath);
              await fsp.unlink(uploaded.filepath || uploaded.path);
            } else {
              throw errMove;
            }
          }

          // Build public URL
          let publicUrl;
          if (settings.ep_images_extended.storage.baseURL) {
            publicUrl = new url.URL(path.posix.join(safePad, newFilename), settings.ep_images_extended.storage.baseURL).toString();
          } else {
            // Default to Etherpad static path assumption
            const relStatic = path.posix.join('/static/images', safePad, newFilename);
            publicUrl = relStatic;
          }

          return res.json({ url: publicUrl });
        });
      } catch (e) {
        logger.error('[ep_images_extended] Local upload error', e);
        return res.status(500).json({ error: 'Failed to process upload' });
      }
    });
  }
  // ---------------------------------------------------------------------
};

// Import/export hooks removed - ccRegisterBlockElements was here
