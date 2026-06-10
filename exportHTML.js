'use strict';
const Changeset = require('ep_etherpad-lite/static/js/Changeset');
const Security = require('ep_etherpad-lite/static/js/security');
const http = require('http');
const https = require('https');

const fetchToBase64 = (src) => new Promise((resolve) => {
  let url;
  if (src.startsWith('//')) {
    url = `http://localhost:9001${src.slice(1)}`;
  } else if (src.startsWith('/')) {
    url = `http://localhost:9001${src}`;
  } else {
    url = src;
  }

  const proto = url.startsWith('https') ? https : http;
  proto.get(url, (res) => {
    const chunks = [];
    res.on('data', (chunk) => chunks.push(chunk));
    res.on('end', () => {
      const buffer = Buffer.concat(chunks);
      const mime = (res.headers['content-type'] || 'image/png').split(';')[0];
      resolve(`data:${mime};base64,${buffer.toString('base64')}`);
    });
    res.on('error', () => resolve(null));
  }).on('error', () => resolve(null));
});

exports.getLineHTMLForExport = async (hook, context) => {
  const attribLine = context.attribLine;
  const apool = context.apool;

  if (!attribLine) return;

  let imgsHTML = '';
  const opIter = Changeset.opIterator(attribLine);

  while (opIter.hasNext()) {
    const op = opIter.next();
    const imageSrcAttrib = Changeset.opAttributeValue(op, 'image', apool);
    if (!imageSrcAttrib) continue;

    try {
      let decodedSrc = decodeURIComponent(imageSrcAttrib);
      if (!decodedSrc) continue;

      const isBase64 = decodedSrc.startsWith('data:');
      const isHttp = decodedSrc.startsWith('http');
      const isServerPath = decodedSrc.startsWith('/');

      if (!isBase64 && !isHttp && !isServerPath) {
        console.warn(`[ep_images_extended exportHTML] Invalid image src: ${decodedSrc}`);
        continue;
      }

      if (!isBase64) {
        const b64 = await fetchToBase64(decodedSrc);
        if (!b64) {
          console.warn(`[ep_images_extended exportHTML] Could not fetch image: ${decodedSrc}`);
          continue;
        }
        decodedSrc = b64;
      }

      const imageWidthAttrib = Changeset.opAttributeValue(op, 'image-width', apool);
      const imageHeightAttrib = Changeset.opAttributeValue(op, 'image-height', apool);
      const imageIdAttrib = Changeset.opAttributeValue(op, 'image-id', apool);

      const escapedSrc = decodedSrc.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
      let tag = `<img src="${escapedSrc}"`;

      if (imageWidthAttrib) {
        tag += ` width="${Security.escapeHTMLAttribute(imageWidthAttrib.replace(/px$/, ''))}"`;
      }
      if (imageHeightAttrib) {
        tag += ` height="${Security.escapeHTMLAttribute(imageHeightAttrib.replace(/px$/, ''))}"`;
      }
      if (imageIdAttrib) {
        tag += ` data-image-id="${Security.escapeHTMLAttribute(imageIdAttrib)}"`;
      }

      tag += `>`;
      imgsHTML += tag;
    } catch (e) {
      console.error(`[ep_images_extended exportHTML] Error processing image: ${imageSrcAttrib}`, e);
    }
  }

  if (!imgsHTML) return;

  const alignMatch = context.lineContent.match(/^<p style='text-align:([^']+)'>([\s\S]*)<\/p>$/);
  if (alignMatch) {
    context.lineContent = `<p style='text-align:${alignMatch[1]}'>${imgsHTML}</p>`;
  } else {
    context.lineContent = `<p>${imgsHTML}</p>`;
  }
};

exports.stylesForExport = (hook, padId, cb) => {
  cb('img { max-width: 100%; vertical-align: middle; }');
};
