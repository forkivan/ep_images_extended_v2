# ep_images_extended_v2

**Insert images inline with text, float them, resize them — with working HTML and DOCX export.**

This is a fork of [ep_images_extended](https://github.com/dcastelone/ep_images_extended) by DCastelone, with critical bug fixes for the HTML and DOCX export pipeline.

![Demo](https://i.imgur.com/qGOBRep.png)

---

## What's fixed in v2

The original plugin stored images correctly in the pad but they were invisible in HTML and DOCX exports. Six bugs were identified and fixed:

**Bug 1 — Export hooks were never registered**
`exportHTML.js` existed but `getLineHTMLForExport` and `stylesForExport` were missing from `ep.json`, so the export code never ran.

**Bug 2 — Text lines were broken**
The export handler modified `lineContent` for every line, even lines with no images. This caused text lines to lose formatting and show `*` placeholders.

**Bug 3 — Only one image per line**
A `break` statement stopped processing after the first image on a line. Multiple side-by-side images were lost.

**Bug 4 — Text alignment was lost**
When `ep_align` was installed alongside this plugin, image lines lost their alignment because the alignment wrapper from `ep_align` was overwritten.

**Bug 5 — DOCX export crashed**
`Security.escapeHTML()` escaped `/` as `&#x2F;`, turning `data:image/png;base64,...` into `data:image&#x2F;png;base64,...`. The `html-to-docx` library could not parse the MIME type and crashed with a null reference error.

**Bug 6 — Local storage images missing from exports**
Images stored on disk (local storage mode) were exported as relative server paths (`/static/images/...`). These paths do not resolve when the HTML file is opened locally or when DOCX is generated. Fixed by fetching images from the internal server and embedding them as base64 at export time.

### Result

| Feature | v1.1.2 | v2.0.0 |
|---------|--------|--------|
| Images visible in HTML export | ✗ | ✓ |
| Images visible in DOCX export | ✗ | ✓ |
| Local storage images in exports | ✗ | ✓ (embedded as base64) |
| Multiple images per line | ✗ | ✓ |
| Text alignment preserved | ✗ | ✓ |
| Text lines unaffected | ✗ | ✓ |

---

## Known limitation — PDF export

PDF export in Etherpad uses `pdfkit` with a built-in HTML parser that does not support images, colors, or advanced formatting from plugins. This is a core Etherpad limitation with no plugin hooks available. Images will not appear in the default PDF export.

---

## Requirements

**Etherpad only** (`ep_etherpad-lite`) — no other plugin is required; this plugin
adds image support itself. For images in **PDF** export specifically, also use a
real PDF exporter such as
[`ep_pdf_export_print`](https://www.npmjs.com/package/ep_pdf_export_print) or
[`ep_pdf_export_chromium`](https://www.npmjs.com/package/ep_pdf_export_chromium),
since core PDF export ignores images (see the note above).

---

## Installation

Via the Etherpad admin panel — search for `ep_images_extended_v2`.

Or from the command line in your Etherpad directory:

```
pnpm run plugins i ep_images_extended_v2
```

---

## Configuration (settings.json)

Add an `ep_images_extended` block to your `settings.json`:

| key | type | default | description |
|-----|------|---------|-------------|
| `fileTypes` | Array\<string\> | _none_ | Allowed extensions (no dot). If omitted, any `image/` MIME is accepted. |
| `maxFileSize` | Number (bytes) | _unlimited_ | Reject files larger than this. |
| `storage` | Object | `{ "type": "base64" }` | Storage backend. See below. |

### Storage strategies

**1. Embedded Base-64** (default — zero config)
```jsonc
"ep_images_extended": {
  "storage": { "type": "base64" },
  "fileTypes": ["jpeg", "jpg", "png", "gif", "bmp", "webp"],
  "maxFileSize": 5000000
}
```

**2. Amazon S3 with presigned uploads**
```jsonc
"ep_images_extended": {
  "storage": {
    "type": "s3_presigned",
    "region": "us-east-1",
    "bucket": "my-etherpad-images",
    "publicURL": "https://cdn.example.com/",
    "expires": 900
  },
  "fileTypes": ["png", "jpg", "webp"],
  "maxFileSize": 10485760
}
```
AWS credentials are read from environment variables (`AWS_ACCESS_KEY_ID`, `AWS_SECRET_ACCESS_KEY`).

**3. Local disk**
```jsonc
"ep_images_extended": {
  "storage": {
    "type": "local",
    "baseFolder": "static/images",
    "baseURL": "https://pad.example.com/etherpad-lite/static/images/"
  },
  "fileTypes": ["jpeg", "jpg", "png", "gif"],
  "maxFileSize": 5000000
}
```

---

## Contributing

Bug reports and PRs are welcome.

---

## Credits

- Original plugin: [ep_images_extended](https://github.com/dcastelone/ep_images_extended) by DCastelone
- Based on: [ep_image_insert](https://github.com/mamylinx/ep_image_insert) by Mamy Linx, John McLear, Ilmar Türk and contributors
- Bug fixes and v2: Ivan Forkaliuk

See [NOTICE.md](NOTICE.md) for full attribution.
