'use strict';
// Modified from ep_image_insert 1.0.7 
// This hook is called **before** the text of a line/segment is processed by the Changeset library.
const collectContentPre = (hook, context) => {
  const classes = context.cls ? context.cls.split(' ') : [];
  let escapedSrc = null;
  let widthValue = null;
  let heightValue = null;
  let aspectRatioValue = null;
  let floatValue = null;
  let imageIdValue = null;

  for (const cls of classes) {
      if (cls.startsWith('image:')) {
          escapedSrc = cls.substring(6);
      } else if (cls.startsWith('image-width:')) {
          const potentialWidth = cls.substring(12);
          if (potentialWidth && (potentialWidth === 'auto' || /[0-9]+(%|px|em|rem|vw|vh)?$/.test(potentialWidth) || /^[0-9.]+$/.test(potentialWidth))) {
             widthValue = potentialWidth;
          }
      } else if (cls.startsWith('image-height:')) {
          const potentialHeight = cls.substring(13);
          if (potentialHeight && (potentialHeight === 'auto' || /[0-9]+(%|px|em|rem|vw|vh)?$/.test(potentialHeight) || /^[0-9.]+$/.test(potentialHeight))) {
             heightValue = potentialHeight;
          }
      } else if (cls.startsWith('imageCssAspectRatio:')) {
          const potentialAspectRatio = cls.substring(20);
          if (!isNaN(parseFloat(potentialAspectRatio))) {
            aspectRatioValue = potentialAspectRatio;
          }
      } else if (cls.startsWith('image-float:')) {
          const potentialFloat = cls.substring(12);
          if (potentialFloat && ['none', 'left', 'right', 'inline'].includes(potentialFloat)) {
            floatValue = potentialFloat;
          }
      } else if (cls.startsWith('image-id-')) {
          const potentialId = cls.substring(9);
          if (potentialId && potentialId.length > 10) {
            imageIdValue = potentialId;
          }
      }
  }

  if (escapedSrc) {
    try {
      context.cc.doAttrib(context.state, `image::${escapedSrc}`);
    } catch (e) {
      console.error('[ep_images_extended collectContentPre] Error applying image attribute:', e);
    }
  }
  if (widthValue) {
    try {
        context.cc.doAttrib(context.state, `image-width::${widthValue}`);
    } catch (e) {
        console.error('[ep_images_extended collectContentPre] Error applying image-width attribute:', e);
    }
  }
  if (heightValue) {
    try {
        context.cc.doAttrib(context.state, `image-height::${heightValue}`);
    } catch (e) {
        console.error('[ep_images_extended collectContentPre] Error applying image-height attribute:', e);
    }
  }
  if (aspectRatioValue) {
    try {
        context.cc.doAttrib(context.state, `imageCssAspectRatio::${aspectRatioValue}`);
    } catch (e) {
        console.error('[ep_images_extended collectContentPre] Error applying imageCssAspectRatio attribute:', e);
    }
  }
  if (floatValue) {
    try {
        context.cc.doAttrib(context.state, `image-float::${floatValue}`);
    } catch (e) {
        console.error('[ep_images_extended collectContentPre] Error applying image-float attribute:', e);
    }
  }
  if (imageIdValue) {
    try {
        context.cc.doAttrib(context.state, `image-id::${imageIdValue}`);
    } catch (e) {
        console.error('[ep_images_extended collectContentPre] Error applying image-id attribute:', e);
    }
  }
};

// This hook is called **after** the text of a line/segment is processed.
// We don't need special post-processing for this attribute approach.
const collectContentPost = (hook, context) => {};

// Remove collectContentImage as it's not suitable for non-<img> elements
// const collectContentImage = ... (Removed)

exports.collectContentPre = collectContentPre;
exports.collectContentPost = collectContentPost;
