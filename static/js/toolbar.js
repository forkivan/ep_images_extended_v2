'use strict';
// Modified from ep_image_insert 1.0.7 
const _isValid = (file) => {
  const mimedb = clientVars.ep_images_extended.mimeTypes;
  const mimeType = mimedb[file.type];
  let validMime = null;
  const errorTitle = html10n.get('ep_images_extended.error.title');

  if (clientVars.ep_images_extended && clientVars.ep_images_extended.fileTypes) {
    validMime = false;
    if (mimeType && mimeType.extensions) {
      for (const fileType of clientVars.ep_images_extended.fileTypes) {
        const exists = mimeType.extensions.indexOf(fileType);
        if (exists > -1) {
          validMime = true;
          break; // Found a valid type
        }
      }
    }
    if (validMime === false) {
      const errorMessage = html10n.get('ep_images_extended.error.fileType');
      $.gritter.add({ title: errorTitle, text: errorMessage, sticky: true, class_name: 'error' });
      return false;
    }
  }

  if (clientVars.ep_images_extended && file.size > clientVars.ep_images_extended.maxFileSize) {
    const allowedSize = (clientVars.ep_images_extended.maxFileSize / 1000000);
    const errorText = html10n.get('ep_images_extended.error.fileSize', { maxallowed: allowedSize });
    $.gritter.add({ title: errorTitle, text: errorText, sticky: true, class_name: 'error' });
    return false;
  }
  return true;
};

// Helper function to insert the image attribute with ZWSP boundaries
const insertInlineImage = (ace, imageData) => {
  const ZWSP = '\u200b';       // Zero-Width Space
  const placeholder = '\u00a0'; // Use NBSP as placeholder
  const textToInsert = ZWSP + placeholder + ZWSP;

  const escapedSrc = escape(imageData.src);
  const attrKey = 'image'; // Standard attribute key
  const attrValue = escapedSrc; // Escaped src is the value

  // Get current selection to replace it (or insert if no selection)
  const rep = ace.ace_getRep();
  const start = rep.selStart;
  const end = rep.selEnd;

  // Add the placeholder characters, replacing selection if any
  ace.ace_replaceRange(start, end, textToInsert);

  // Define the range for the attribute: ONLY the middle character (the placeholder NBSP)
  const attrStart = [start[0], start[1] + ZWSP.length]; // Start after the first ZWSP
  const attrEnd = [start[0], start[1] + ZWSP.length + placeholder.length]; // End after the placeholder NBSP

  // Apply the image attribute to the placeholder character
  ace.ace_performDocumentApplyAttributesToRange(attrStart, attrEnd, [[attrKey, attrValue]]);

  // Move cursor after all inserted characters
  const finalEnd = [start[0], start[1] + textToInsert.length];
  ace.ace_performSelectionChange(finalEnd, finalEnd, false);
  ace.ace_focus();
};


exports.postToolbarInit = (hook, context) => {
  const toolbar = context.toolbar;
  toolbar.registerCommand('imageUpload', () => {
    $(document).find('body').find('#imageInput').remove();
    const fileInputHtml = `<input
    style="width:1px;height:1px;z-index:-10000;"
    id="imageInput" type="file" />`;
    $(document).find('body').append(fileInputHtml);

    $(document).find('body').find('#imageInput').on('change', (e) => {
      const files = e.target.files;
      if (!files.length) {
        // No specific user message needed here, browser handles it or no file selected is not an error
        return;
      }
      const file = files[0];

      if (!_isValid(file)) {
        return; // Validation errors are handled by _isValid
      }

      if (clientVars.ep_images_extended.storageType === 'base64') {
        $('#imageUploadModalLoader').removeClass('popup-show'); // Ensure loader is hidden initially
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => {
          const data = reader.result;
          const img = new Image();
          img.onload = () => {
            const widthPx = `${img.naturalWidth}px`;
            const heightPx = `${img.naturalHeight}px`;
            context.ace.callWithAce((ace) => {
              ace.ace_doInsertImage(data, widthPx, heightPx);
            }, 'imgBase64', true);
          };
          img.onerror = () => {
             console.error('[ep_images_extended toolbar] Failed to load Base64 image data to get dimensions. Inserting without dimensions.');
             context.ace.callWithAce((ace) => {
               ace.ace_doInsertImage(data);
            }, 'imgBase64Error', true);
          };
          img.src = data;
        };
        reader.onerror = (error_evt) => { // Added error handling for FileReader
            console.error('[ep_images_extended toolbar] FileReader error:', error_evt);
            const errorTitle = html10n.get('ep_images_extended.error.title');
            const errorMessage = html10n.get('ep_images_extended.error.fileRead'); // Generic file read error
            $.gritter.add({ title: errorTitle, text: errorMessage, sticky: true, class_name: 'error' });
        };
      } else if (clientVars.ep_images_extended.storageType === 's3_presigned') {
        // -------- Direct browser -> S3 upload via presigned URL --------
        const queryParams = $.param({ name: file.name, type: file.type });
        $('#imageUploadModalLoader').addClass('popup-show');

        $.getJSON(`${clientVars.padId}/pluginfw/ep_images_extended/s3_presign?${queryParams}`)
          .then((presignData) => {
            if (!presignData || !presignData.signedUrl || !presignData.publicUrl) {
              throw new Error('Invalid presign response');
            }

            // Upload the file directly to S3
            return fetch(presignData.signedUrl, {
              method: 'PUT',
              headers: { 'Content-Type': file.type },
              body: file,
            }).then((response) => {
              if (!response.ok) {
                throw new Error(`S3 upload failed with status ${response.status}`);
              }
              return presignData.publicUrl;
            });
          })
          .then((publicUrl) => {
            // Remove loader
            $('#imageUploadModalLoader').removeClass('popup-show');

            // Obtain intrinsic dimensions by loading image
            const img = new Image();
            img.onload = () => {
              const widthPx = `${img.naturalWidth}px`;
              const heightPx = `${img.naturalHeight}px`;
              context.ace.callWithAce((ace) => {
                ace.ace_doInsertImage(publicUrl, widthPx, heightPx);
              }, 'imgUploadS3', true);
            };
            img.onerror = () => {
              console.warn('[ep_images_extended toolbar] Could not load uploaded S3 image to measure size. Inserting without dimensions.');
              context.ace.callWithAce((ace) => {
                ace.ace_doInsertImage(publicUrl);
              }, 'imgUploadS3Error', true);
            };
            img.src = publicUrl;
          })
          .catch((err) => {
            console.error('[ep_images_extended toolbar] s3_presigned upload failed', err);
            $('#imageUploadModalLoader').removeClass('popup-show');
            const errorTitle = html10n.get('ep_images_extended.error.title');
            $.gritter.add({ title: errorTitle, text: err.message, sticky: true, class_name: 'error' });
          });
      } else if (clientVars.ep_images_extended.storageType === 'local') {
        // -------- Upload to server local storage endpoint --------
        $('#imageUploadModalLoader').addClass('popup-show');
        const formData = new FormData();
        formData.append('file', file, file.name);

        fetch(`${clientVars.padId}/pluginfw/ep_images_extended/upload`, {
          method: 'POST',
          body: formData,
        })
          .then(async (response) => {
            if (!response.ok) {
              throw new Error(`Upload failed with status ${response.status}`);
            }
            const data = await response.json();
            if (!data || !data.url) throw new Error('Invalid upload response');
            return data.url;
          })
          .then((publicUrl) => {
            $('#imageUploadModalLoader').removeClass('popup-show');
            const img = new Image();
            img.onload = () => {
              const widthPx = `${img.naturalWidth}px`;
              const heightPx = `${img.naturalHeight}px`;
              context.ace.callWithAce((ace) => {
                ace.ace_doInsertImage(publicUrl, widthPx, heightPx);
              }, 'imgUploadLocal', true);
            };
            img.onerror = () => {
              context.ace.callWithAce((ace) => {
                ace.ace_doInsertImage(publicUrl);
              }, 'imgUploadLocalError', true);
            };
            img.src = publicUrl;
          })
          .catch((err) => {
            console.error('[ep_images_extended toolbar] local upload failed', err);
            $('#imageUploadModalLoader').removeClass('popup-show');
            const errorTitle = html10n.get('ep_images_extended.error.title');
            $.gritter.add({ title: errorTitle, text: err.message, sticky: true, class_name: 'error' });
          });
      } else {
        // Unsupported storage type – show error and abort
        $('#imageUploadModalLoader').removeClass('popup-show');
        const errorTitle = html10n.get('ep_images_extended.error.title');
        const errorText = `Unsupported storageType: ${clientVars.ep_images_extended.storageType}. Only "base64", "s3_presigned", and "local" are supported.`;
        $.gritter.add({ title: errorTitle, text: errorText, sticky: true, class_name: 'error' });
      }
    });
    $(document).find('body').find('#imageInput').trigger('click');
  });
};
