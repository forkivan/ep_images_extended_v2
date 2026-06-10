'use strict';
// Modified from ep_image_insert 1.0.7 

console.log('ep_images_extended version 1.1.1');

// Optional helper (shared with ep_docx_html_customizer) that provides a CORS fetch with
// automatic same-origin proxy fallback.  If the plugin is not present we simply fall back
// to the native fetch logic.
let fetchWithCorsProxy;
try {
  ({fetchWithCorsProxy} = require('../../../ep_docx_html_customizer/transform_common'));
} catch (_) { /* helper not available – fallback to plain fetch */ }

if (!fetchWithCorsProxy && typeof window !== 'undefined') {
  fetchWithCorsProxy = window.fetchWithCorsProxy;
}

// Simple UUID generator
function generateUUID() {
  return 'xxxx-xxxx-4xxx-yxxx-xxxx'.replace(/[xy]/g, function(c) {
    const r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

// Helper function to find image placeholder position robustly
function findImagePlaceholderPosition(lineText, imageIndex, fallbackLineElement = null) {
  // ---- Primary strategy: use canonical placeholder scan -------------------
  const canonicalRanges = getAllPlaceholderRanges(lineText);
  if (imageIndex >= 0 && imageIndex < canonicalRanges.length) {
    const r = canonicalRanges[imageIndex];
    return {
      colStart: r.colStart,
      patternLength: r.patternLength,
      pattern: lineText.substr(r.colStart, r.patternLength),
    };
  }

  // Fallback heuristics (legacy) -------------------------------------------
  // Try different placeholder patterns that might exist
  const placeholderPatterns = [
    '\u200B\u200B\u200B', // 3 × ZWSP (current)
    '\u200B\u00A0\u200B', // ZWSP NBSP ZWSP (toolbar-inserted, incl. tables)
    '\u200B\u200B',        // 2 × ZWSP (collapsed)
    '\u00A0',              // single NBSP (edge / collapsed)
    '\u200B'               // single ZWSP (legacy)
  ];
  
  for (const pattern of placeholderPatterns) {
    let searchFromIndex = 0;
    let foundCount = 0;
    
    while (foundCount <= imageIndex) {
      const foundIndex = lineText.indexOf(pattern, searchFromIndex);
      if (foundIndex === -1) {
        break; // Pattern not found anymore
      }
      
      if (foundCount === imageIndex) {
        // Found the target placeholder
        return {
          colStart: foundIndex,
          patternLength: pattern.length,
          pattern: pattern
        };
      }
      
      foundCount++;
      searchFromIndex = foundIndex + pattern.length;
    }
  }
  
  // Fallback: try to find any zero-width space near expected position
  const zwspPattern = '\u200B';
  const expectedPosition = imageIndex * 3; // Rough estimate
  const searchStart = Math.max(0, expectedPosition - 10);
  const searchEnd = Math.min(lineText.length, expectedPosition + 20);
  
  for (let i = searchStart; i < searchEnd; i++) {
    if (lineText[i] === zwspPattern) {
      // console.log(`[ep_images_extended] Fallback: Found ZWSP at position ${i} for image index ${imageIndex}`);
      return {
        colStart: i,
        patternLength: 1,
        pattern: zwspPattern
      };
    }
  }
  
  // DOM-based fallback: when text-based approach fails, use DOM positioning
  if (fallbackLineElement) {
    // console.log(`[ep_images_extended] DOM fallback: Using DOM-based positioning for image index ${imageIndex}`);
    
    // Count characters before the target image element in the DOM
    const allImagePlaceholders = Array.from(fallbackLineElement.querySelectorAll('.inline-image.image-placeholder'));
    if (imageIndex < allImagePlaceholders.length) {
      const targetImageElement = allImagePlaceholders[imageIndex];
      
      // Calculate approximate text position by walking through DOM text nodes
      let approximatePosition = 0;
      const walker = document.createTreeWalker(
        fallbackLineElement,
        NodeFilter.SHOW_TEXT,
        null,
        false
      );
      
      let currentNode = walker.nextNode();
      while (currentNode) {
        if (currentNode.parentNode && 
            (currentNode.parentNode === targetImageElement || 
             targetImageElement.contains(currentNode.parentNode))) {
          // Found a text node that belongs to our target image
          // console.log(`[ep_images_extended] DOM fallback: Found position ${approximatePosition} for image index ${imageIndex}`);
          return {
            colStart: approximatePosition,
            patternLength: 3, // Assume 3 characters for safety
            pattern: '\u200B\u200B\u200B',
            isDomFallback: true
          };
        }
        // Add the length of this text node to our position counter
        approximatePosition += currentNode.textContent.length;
        currentNode = walker.nextNode();
      }
      
      // If we couldn't find the exact position, use a reasonable estimate
      // console.log(`[ep_images_extended] DOM fallback: Using estimated position ${approximatePosition} for image index ${imageIndex}`);
      return {
        colStart: Math.max(0, approximatePosition - 1),
        patternLength: 3,
        pattern: '\u200B\u200B\u200B',
        isDomFallback: true
      };
    }
  }
  
  return null; // No placeholder found
}

// Helper function to validate ace operations and document state
function validateAceOperation(ace, operation, rangeStart, rangeEnd, context = '') {
  try {
    // Validate that ace exists and has required methods
    if (!ace) {
      console.error(`[ep_images_extended ${context}] ace object not available`);
      return false;
    }
    
    // Check for required methods based on operation type
    if (operation === 'applyAttributes' && typeof ace.ace_performDocumentApplyAttributesToRange !== 'function') {
      console.error(`[ep_images_extended ${context}] ace_performDocumentApplyAttributesToRange not available`);
      return false;
    }
    
    if (operation === 'replaceRange' && typeof ace.ace_replaceRange !== 'function') {
      console.error(`[ep_images_extended ${context}] ace_replaceRange not available`);
      return false;
    }
    
    // Get current rep to validate document state
    const rep = ace.ace_getRep();
    if (!rep || !rep.lines) {
      console.error(`[ep_images_extended ${context}] Document rep not available or invalid`);
      return false;
    }
    
    // Validate range bounds if provided
    if (rangeStart && rangeEnd) {
      const [startLine, startCol] = rangeStart;
      const [endLine, endCol] = rangeEnd;
      
      // Check if line numbers are valid
      if (startLine < 0 || endLine < 0 || startLine >= rep.lines.length() || endLine >= rep.lines.length()) {
        console.error(`[ep_images_extended ${context}] Invalid line numbers in range:`, [rangeStart, rangeEnd], 'total lines:', rep.lines.length());
        return false;
      }
      
      // Check if the lines still exist
      const startLineObj = rep.lines.atIndex(startLine);
      const endLineObj = rep.lines.atIndex(endLine);
      if (!startLineObj || !endLineObj) {
        console.error(`[ep_images_extended ${context}] One or more lines no longer exist:`, startLine, endLine);
        return false;
      }
      
      // Validate column bounds
      const startLineText = startLineObj.text;
      const endLineText = endLineObj.text;
      if (startCol < 0 || endCol < 0 || startCol > startLineText.length || endCol > endLineText.length) {
        console.error(`[ep_images_extended ${context}] Invalid column positions in range:`, [rangeStart, rangeEnd], 
                     'start line length:', startLineText.length, 'end line length:', endLineText.length);
        return false;
      }
    }
    
    return true;
  } catch (error) {
    console.error(`[ep_images_extended ${context}] Error during validation:`, error);
    return false;
  }
}

exports.aceAttribsToClasses = function(name, context) {
  if (context.key === 'image' && context.value) {
    return ['image:' + context.value];
  }
  if (context.key === 'image-width' && context.value) {
    return ['image-width:' + context.value];
  }
  if (context.key === 'image-height' && context.value) {
    return ['image-height:' + context.value];
  }
  // ADDED for imageCssAspectRatio
  if (context.key === 'imageCssAspectRatio' && context.value) {
    return ['imageCssAspectRatio:' + context.value];
  }
  // ADDED for image float style
  if (context.key === 'image-float' && context.value) {
    return ['image-float:' + context.value];
  }
  // ADDED for persistent image ID
  if (context.key === 'image-id' && context.value) {
    return ['image-id-' + context.value];
  }
  return [];
};

exports.aceInitialized = (hook, context) => {
  // Bind the new image insertion function
  context.editorInfo.ace_doInsertImage = doInsertImage.bind(context);
};

// Function to render placeholders into actual images (Currently unused due to CSS background approach)
/*
const renderImagePlaceholders = (rootElement) => {
  const placeholders = $(rootElement).find('span.image-placeholder');
  placeholders.each(function() {
    const $placeholder = $(this);
    if ($placeholder.data('processed-image')) {
        return;
    }
    const attribsData = $placeholder.data('image-attribs');
    if (typeof attribsData === 'string') {
      try {
        const imageData = JSON.parse(attribsData);
        if (imageData && imageData.src) {
          const $img = $('<img>').attr('src', imageData.src);
          $img.css({
            'display': 'inline-block',
            'max-width': '100%',
            'max-height': '20em',
            'vertical-align': 'middle'
          });
          if (imageData.width) $img.attr('width', imageData.width);
          if (imageData.height) $img.attr('height', imageData.height);
          $placeholder.empty().append($img);
          $placeholder.data('processed-image', true);
        } else {
          $placeholder.text('[Invalid Image]');
          $placeholder.data('processed-image', true);
        }
      } catch (e) {
        console.error('[ep_images_extended] Failed to parse image data:', attribsData, e);
        $placeholder.text('[Parse Error]');
        $placeholder.data('processed-image', true);
      }
    } else {
      $placeholder.text('[Missing Data]');
      $placeholder.data('processed-image', true);
    }
  });
};
*/

exports.postAceInit = function (hook, context) {
  const padOuter = $('iframe[name="ace_outer"]').contents().find('body');
  if (padOuter.length === 0) {
      console.error('[ep_images_extended postAceInit] Could not find outer pad body.');
      return;
  }

  if ($('#imageResizeOutline').length === 0) {
      const $outlineBox = $('<div id="imageResizeOutline"></div>');
      $outlineBox.css({
          position: 'absolute',
          border: '1px dashed #1a73e8',
          backgroundColor: 'rgba(26, 115, 232, 0.1)',
          'pointer-events': 'none',
          display: 'none',
          'z-index': 1000,
          'box-sizing': 'border-box'
      });
      padOuter.append($outlineBox);
  }

  // Check if image formatting menu exists (should be loaded from template)
  if ($('#imageFormatMenu').length === 0) {
      // console.log('[ep_images_extended] Image format menu not found - template may not be loaded yet');
  } else {
      // console.log('[ep_images_extended] Image format menu found from template');
  }

  const $outlineBoxRef = padOuter.find('#imageResizeOutline');
  // Look for the format menu in the main document (loaded from template)
  const $formatMenuRef = $('#imageFormatMenu');
  const _aceContext = context.ace;

  if (!$outlineBoxRef || $outlineBoxRef.length === 0) {
     console.error('[ep_images_extended postAceInit] FATAL: Could not find #imageResizeOutline OUTSIDE callWithAce.');
     return; 
  }

  if (!$formatMenuRef || $formatMenuRef.length === 0) {
     console.error('[ep_images_extended postAceInit] FATAL: Could not find #imageFormatMenu OUTSIDE callWithAce.');
     return; 
  }

  context.ace.callWithAce((ace) => {
    const $innerIframe = $('iframe[name="ace_outer"]').contents().find('iframe[name="ace_inner"]');
    if ($innerIframe.length === 0) {
        console.error('ep_images_extended: ERROR - Could not find inner iframe (ace_inner).');
        return;
    }
    const innerDocBody = $innerIframe.contents().find('body')[0];
    const $inner = $(innerDocBody);
    const innerDoc = $innerIframe.contents();

    if (!$inner || $inner.length === 0) {
        console.error('ep_images_extended: ERROR - Could not get body from inner iframe.');
        return;
    }

    let isDragging = false;
    let startX = 0;
    let startWidth = 0;
    let startHeight = 0;
    // let aspectRatio = 1; // Not directly used for height calculation anymore
    let currentVisualAspectRatioHW = 1;
    let targetOuterSpan = null;
    let targetInnerSpan = null;
    let targetLineNumber = -1;
    let outlineBoxPositioned = false;
    let mousedownClientX = 0;
    let mousedownClientY = 0;
    let clickedHandle = null;
    
    // Use a global approach with CSS classes on the body and unique image identifiers
    let selectedImageSrc = null;
    let selectedImageElement = null; // Track the specific selected element
    let selectedImageLine = -1; // Track line number for persistent selection
    let selectedImageCol = -1; // Track column for persistent selection
    
    // NEW: Store the active image's unique ID
    window.epImageInsertActiveImageId = null;
    
    // Store resize positioning data without DOM attributes to avoid triggering content collection
    let resizePositionData = null;

    // Function to position and show the format menu below the selected image
    const showFormatMenu = (imageElement) => {
        // Don't show format menu in read-only mode
        if (clientVars && clientVars.readonly) {
            return;
        }
        
        if (!imageElement || !$formatMenuRef || $formatMenuRef.length === 0) {
            return;
        }
        
        try {
            const innerSpan = imageElement.querySelector('span.image-inner');
            if (!innerSpan) {
                return;
            }
            
            // Use the same complex positioning logic as the resize outline
            const imageRect = innerSpan.getBoundingClientRect();
            
            // Get all the necessary container references
            let innerBodyRect, innerIframeRect, outerBodyRect;
            let scrollTopInner, scrollLeftInner, scrollTopOuter, scrollLeftOuter;
            try {
                innerBodyRect = innerDocBody.getBoundingClientRect();
                innerIframeRect = $innerIframe[0].getBoundingClientRect();
                outerBodyRect = padOuter[0].getBoundingClientRect();
                scrollTopInner = innerDocBody.scrollTop;
                scrollLeftInner = innerDocBody.scrollLeft;
                scrollTopOuter = padOuter.scrollTop();
                scrollLeftOuter = padOuter.scrollLeft();
            } catch (e) {
                console.error('[ep_images_extended showFormatMenu] Error getting container rects/scrolls:', e);
                return; 
            }
            
            // Calculate position using the same method as resize outline
            const imageTopRelInner = imageRect.top - innerBodyRect.top + scrollTopInner;
            const imageLeftRelInner = imageRect.left - innerBodyRect.left + scrollLeftInner;
            const imageBottomRelInner = imageRect.bottom - innerBodyRect.top + scrollTopInner;
            
            const innerFrameTopRelOuter = innerIframeRect.top - outerBodyRect.top + scrollTopOuter;
            const innerFrameLeftRelOuter = innerIframeRect.left - outerBodyRect.left + scrollLeftOuter;
            
            const menuTopOuter = innerFrameTopRelOuter + imageBottomRelInner + 8; // 8px gap below image
            // For left alignment, we want the menu's left edge to match the image's left edge
            const menuLeftOuter = innerFrameLeftRelOuter + imageLeftRelInner;
            
            // Apply the same padding adjustments as resize outline
            const outerPadding = window.getComputedStyle(padOuter[0]);
            const outerPaddingTop = parseFloat(outerPadding.paddingTop) || 0;
            const outerPaddingLeft = parseFloat(outerPadding.paddingLeft) || 0; 
            
            // Apply manual offsets similar to resize outline, but adjusted for menu
            const MENU_MANUAL_OFFSET_TOP = 9; // Same as resize outline
            const MENU_MANUAL_OFFSET_LEFT = 37; // Reduced by 5px to move menu closer to left margin
            
            const finalMenuTop = menuTopOuter + outerPaddingTop + MENU_MANUAL_OFFSET_TOP;
            const finalMenuLeft = menuLeftOuter + outerPaddingLeft + MENU_MANUAL_OFFSET_LEFT;
            
            // Position the menu using absolute positioning relative to padOuter (like resize outline)
            $formatMenuRef.css({
                position: 'absolute',
                left: Math.max(10, finalMenuLeft) + 'px',
                top: finalMenuTop + 'px',
                'z-index': '10000'
            }).addClass('visible');
            
            // Move the menu to padOuter if it's not already there
            if ($formatMenuRef.parent()[0] !== padOuter[0]) {
                padOuter.append($formatMenuRef);
            }
            
            // Update menu button states based on current image float style
            updateMenuButtonStates(imageElement);
            
        } catch (e) {
            console.error('[ep_images_extended] Error positioning format menu:', e);
        }
    };
    
    // Function to hide the format menu
    const hideFormatMenu = () => {
        if ($formatMenuRef) {
            $formatMenuRef.removeClass('visible');
        }
    };
    
    // Function to update menu button states based on current image float style
    const updateMenuButtonStates = (imageElement) => {
        if (!imageElement || !$formatMenuRef || $formatMenuRef.length === 0) {
            return;
        }
        
        // Determine current float state
        let currentFloat = 'inline'; // default
        if (imageElement.classList.contains('image-float-left')) {
            currentFloat = 'left';
        } else if (imageElement.classList.contains('image-float-right')) {
            currentFloat = 'right';
        } else if (imageElement.classList.contains('image-float-none')) {
            currentFloat = 'inline';
        }
        
        // Update button active states
        $formatMenuRef.find('.image-format-button[data-wrap]').removeClass('active');
        $formatMenuRef.find(`.image-format-button[data-wrap="${currentFloat}"]`).addClass('active');
    };

    $inner.on('mousedown', '.inline-image.image-placeholder', function(evt) {
        if (evt.button !== 0) return;
        
        // Don't allow interaction in read-only mode
        if (clientVars && clientVars.readonly) {
            return;
        }
        
        // *** ENHANCED DEBUG: Track image click within table context ***
        // console.log('[ep_images_extended] *** IMAGE MOUSEDOWN EVENT START ***');
        // console.log('[ep_images_extended] Event target:', evt.target);
        // console.log('[ep_images_extended] Event currentTarget:', evt.currentTarget);
        
        targetOuterSpan = this;
        const $targetOuterSpan = $(targetOuterSpan);
        // console.log('[ep_images_extended] Target outer span element:', targetOuterSpan);
        // console.log('[ep_images_extended] Target outer span classes:', targetOuterSpan.className);
        // console.log('[ep_images_extended] Target outer span HTML length:', targetOuterSpan.outerHTML?.length || 0);

        // *** DEBUG: Check table context ***
        const closestTable = targetOuterSpan.closest('table.dataTable');
        const closestTableCell = targetOuterSpan.closest('td, th');
        const closestAceLine = targetOuterSpan.closest('.ace-line');
        
        // console.log('[ep_images_extended] Is image within table?', !!closestTable);
        if (closestTable) {
            // console.log('[ep_images_extended] Table tblId:', closestTable.getAttribute('data-tblId'));
            // console.log('[ep_images_extended] Table row:', closestTable.getAttribute('data-row'));
            // console.log('[ep_images_extended] Table cell:', !!closestTableCell);
            if (closestTableCell) {
                // console.log('[ep_images_extended] Cell data-column:', closestTableCell.getAttribute('data-column'));
                // console.log('[ep_images_extended] Cell innerHTML length:', closestTableCell.innerHTML?.length || 0);
            }
        }
        if (closestAceLine) {
            // console.log('[ep_images_extended] Ace line ID:', closestAceLine.id);
            // console.log('[ep_images_extended] Ace line classes:', closestAceLine.className);
            // console.log('[ep_images_extended] Ace line innerHTML length:', closestAceLine.innerHTML?.length || 0);
        }

        const imageId = $targetOuterSpan.attr('data-image-id');
        // console.log('[ep_images_extended] Image ID:', imageId);
        
        if (imageId) {
            const previouslyActiveId = window.epImageInsertActiveImageId;
            // console.log('[ep_images_extended] Previously active image ID:', previouslyActiveId);
            // console.log('[ep_images_extended] Setting new active image ID:', imageId);
            
            window.epImageInsertActiveImageId = imageId; // NEW: Set active ID

            // *** DEBUG: Track selection styling changes ***
            // console.log('[ep_images_extended] *** SELECTION STYLING START ***');
            
            // Use dynamic CSS injection to avoid triggering content collector on image elements
            if (previouslyActiveId !== imageId) {
                // console.log('[ep_images_extended] Updating dynamic CSS selection');
                const innerDoc = $inner[0].ownerDocument;
                
                // Remove previous dynamic style if it exists
                let existingStyle = innerDoc.getElementById('ep-image-selection-style');
                if (existingStyle) {
                    existingStyle.remove();
                }
                
                // Create new dynamic style element for the selected image
                if (imageId) {
                    const styleElement = innerDoc.createElement('style');
                    styleElement.id = 'ep-image-selection-style';
                    styleElement.textContent = `
                        span.inline-image.image-placeholder[data-image-id="${imageId}"] span.image-inner {
                            outline: 1px solid #1a73e8 !important;
                            outline-offset: 1px !important;
                        }
                        span.inline-image.image-placeholder[data-image-id="${imageId}"] span.image-resize-handle {
                            display: block !important;
                        }
                    `;
                    innerDoc.head.appendChild(styleElement);
                    // console.log('[ep_images_extended] Added dynamic CSS for image:', imageId);
                }
            }
            
            // console.log('[ep_images_extended] *** SELECTION STYLING END ***');
            
            // *** DEBUG: Check DOM state after style changes ***
            if (closestAceLine) {
                // console.log('[ep_images_extended] *** DOM STATE AFTER STYLE CHANGES ***');
                // console.log('[ep_images_extended] Ace line innerHTML length after style changes:', closestAceLine.innerHTML?.length || 0);
                // console.log('[ep_images_extended] Ace line innerHTML (first 500 chars):', closestAceLine.innerHTML?.substring(0, 500) || '');
                
                // Check for delimiter presence in the line
                const delimiterCount = (closestAceLine.innerHTML || '').split('|').length - 1;
                // console.log('[ep_images_extended] Delimiter count in ace line after style changes:', delimiterCount);
                
                // Check if tbljson class is still present
                const hasTbljsonClass = closestAceLine.innerHTML?.includes('tbljson-') || false;
                // console.log('[ep_images_extended] Ace line still has tbljson class after changes:', hasTbljsonClass);
            }
            
            showFormatMenu(targetOuterSpan);
        } else {
             console.warn('[ep_images_extended mousedown] Image clicked has no data-image-id.');
        }

        targetInnerSpan = targetOuterSpan.querySelector('span.image-inner');
        if (!targetInnerSpan) {
            console.error('[ep_images_extended mousedown] Could not find inner span.');
            targetOuterSpan = null;
            $targetOuterSpan.removeClass('selected');
            return;
        }

        const target = $(evt.target);
        const isResizeHandle = target.hasClass('image-resize-handle');
        // console.log('[ep_images_extended] Is resize handle clicked?', isResizeHandle);

        // If clicking on a resize handle, start the resize operation
        if (isResizeHandle) {
            // console.log('[ep_images_extended] *** RESIZE HANDLE CLICKED - STARTING RESIZE OPERATION ***');
            isDragging = true;
            outlineBoxPositioned = false;
            startX = evt.clientX;
            mousedownClientX = evt.clientX;
            mousedownClientY = evt.clientY;
            startWidth = targetInnerSpan.offsetWidth || parseInt(targetInnerSpan.style.width, 10) || 0;
            startHeight = targetInnerSpan.offsetHeight || parseInt(targetInnerSpan.style.height, 10) || 0;
            currentVisualAspectRatioHW = (startWidth > 0 && startHeight > 0) ? (startHeight / startWidth) : 1;

            if (target.hasClass('br')) clickedHandle = 'br';
            else clickedHandle = null;

            const lineElement = $(targetOuterSpan).closest('.ace-line')[0];

            if (lineElement) {
                const allImagePlaceholdersInLine = Array.from(lineElement.querySelectorAll('.inline-image.image-placeholder'));
                const imageIndex = allImagePlaceholdersInLine.indexOf(targetOuterSpan);

                if (imageIndex === -1) {
                    console.error('[ep_images_extended mousedown] Clicked image placeholder not found within its line DOM elements.');
                    isDragging = false;
                    resizePositionData = null;
                    $targetOuterSpan.removeClass('selected');
                    targetOuterSpan = null;
                    return;
                }

                targetLineNumber = _getLineNumberOfElement(lineElement);
                
                // Store positioning data in JavaScript variable instead of DOM attributes to avoid content collection
                _aceContext.callWithAce((ace) => {
                    const rep = ace.ace_getRep();
                    if (!rep.lines.atIndex(targetLineNumber)) {
                        console.error(`[ep_images_extended mousedown] Line ${targetLineNumber} does not exist in rep.`);
                        resizePositionData = null;
                        return;
                    }
                    const lineText = rep.lines.atIndex(targetLineNumber).text;
                    
                    // Use helper function to find placeholder position
                    const placeholderInfo = findImagePlaceholderPosition(lineText, imageIndex, lineElement);
                    
                    if (placeholderInfo) {
                        // Store in JS variable instead of DOM attributes to avoid triggering content collection
                        resizePositionData = {
                            lineNumber: targetLineNumber,
                            colStart: placeholderInfo.colStart,
                            patternLength: placeholderInfo.patternLength
                        };
                        // console.log(`[ep_images_extended mousedown] Found placeholder at position ${placeholderInfo.colStart} with pattern length ${placeholderInfo.patternLength}`);
                    } else {
                        console.error(`[ep_images_extended mousedown] Could not find any placeholder sequence for image index ${imageIndex} in line text: "${lineText}"`);
                        resizePositionData = null;
                    }
                }, 'getImageColStart', true);
            } else {
                console.error('[ep_images_extended mousedown] Could not find parent .ace-line for the clicked image.');
                isDragging = false;
                resizePositionData = null;
                $targetOuterSpan.removeClass('selected');
                targetOuterSpan = null;
                return;
            }
            evt.preventDefault();
          } else {
            // console.log('[ep_images_extended] *** SIMPLE IMAGE CLICK - NO RESIZE HANDLE ***');
            // Position cursor next to the image instead of inside it
            _aceContext.callWithAce((ace) => {
                const lineElement = $(targetOuterSpan).closest('.ace-line')[0];
                if (lineElement) {
                    const lineNumber = _getLineNumberOfElement(lineElement);
                    const rep = ace.ace_getRep();
                    if (rep.lines.atIndex(lineNumber)) {
                        const lineText = rep.lines.atIndex(lineNumber).text;
                        const allImages = Array.from(lineElement.querySelectorAll('.inline-image.image-placeholder'));
                        const imageIndex = allImages.indexOf(targetOuterSpan);
                        
                        if (imageIndex !== -1) {
                            const placeholderInfo = findImagePlaceholderPosition(lineText, imageIndex, lineElement);
                            if (placeholderInfo) {
                                // Determine cursor position based on click location relative to image
                                const imageRect = targetInnerSpan.getBoundingClientRect();
                                const clickX = evt.clientX;
                                const imageCenterX = imageRect.left + imageRect.width / 2;
                                
                                let cursorPos;
                                if (clickX < imageCenterX) {
                                    // Clicked left side - place cursor before image
                                    cursorPos = [lineNumber, placeholderInfo.colStart];
                                } else {
                                    // Clicked right side - place cursor after image
                                    cursorPos = [lineNumber, placeholderInfo.colStart + placeholderInfo.patternLength];
                                }
                                
                                // console.log(`[ep_images_extended] Positioning cursor at [${cursorPos}] based on click position`);
                                ace.ace_performSelectionChange(cursorPos, cursorPos, false);
                            }
                        }
                    }
                }
            }, 'positionCursorNextToImage', true);
        }
        
        // console.log('[ep_images_extended] *** IMAGE MOUSEDOWN EVENT END ***');
    });

    innerDoc.on('mousemove', function(evt) {
        if (isDragging) {
            if (!outlineBoxPositioned) {
                 if (!targetInnerSpan || !padOuter || !targetOuterSpan || !innerDocBody || !$innerIframe) {
                     console.error('[ep_images_extended mousemove] Cannot position outline: Required elements missing.');
                     return;
                 }
                 const currentWidth = startWidth;
                 const currentHeight = startHeight;

                 // if (currentWidth <= 0 || currentHeight <= 0) { /* Warning for this was removed */ }

                 let innerBodyRect, innerIframeRect, outerBodyRect;
                 let scrollTopInner, scrollLeftInner, scrollTopOuter, scrollLeftOuter;
                 try {
                     innerBodyRect = innerDocBody.getBoundingClientRect();
                     innerIframeRect = $innerIframe[0].getBoundingClientRect();
                     outerBodyRect = padOuter[0].getBoundingClientRect();
                     scrollTopInner = innerDocBody.scrollTop;
                     scrollLeftInner = innerDocBody.scrollLeft;
                     scrollTopOuter = padOuter.scrollTop();
                     scrollLeftOuter = padOuter.scrollLeft();
                 } catch (e) {
                     console.error('[ep_images_extended mousemove] Error getting container rects/scrolls:', e);
                     return; 
                 }

                 const clickTopRelInner = mousedownClientY - innerBodyRect.top + scrollTopInner;
                 const clickLeftRelInner = mousedownClientX - innerBodyRect.left + scrollLeftInner;
                 const innerFrameTopRelOuter = innerIframeRect.top - outerBodyRect.top + scrollTopOuter;
                 const innerFrameLeftRelOuter = innerIframeRect.left - outerBodyRect.left + scrollLeftOuter;
                 const baseClickTopOuter = innerFrameTopRelOuter + clickTopRelInner;
                 const baseClickLeftOuter = innerFrameLeftRelOuter + clickLeftRelInner;
                 let outlineTop = baseClickTopOuter;
                 let outlineLeft = baseClickLeftOuter;

                 // For bottom-right handle, position outline at top-left of image
                 if (clickedHandle === 'br') {
                    outlineLeft -= currentWidth;
                    outlineTop -= currentHeight;
                 }

                 const outerPadding = window.getComputedStyle(padOuter[0]);
                 const outerPaddingTop = parseFloat(outerPadding.paddingTop) || 0;
                 const outerPaddingLeft = parseFloat(outerPadding.paddingLeft) || 0; 
                 const finalOutlineTop = outlineTop + outerPaddingTop; 
                 const finalOutlineLeft = outlineLeft + outerPaddingLeft; 
                 const MANUAL_OFFSET_TOP = 9;
                 const MANUAL_OFFSET_LEFT = 42;
                 const finalTopWithManualOffset = finalOutlineTop + MANUAL_OFFSET_TOP; 
                 const finalLeftWithManualOffset = finalOutlineLeft + MANUAL_OFFSET_LEFT;

                 $outlineBoxRef.css({
                     left: finalLeftWithManualOffset + 'px', 
                     top: finalTopWithManualOffset + 'px',   
                     width: currentWidth + 'px',
                     height: currentHeight + 'px',
                     display: 'block'
                 });
                 outlineBoxPositioned = true;
            }

            if ($outlineBoxRef && $outlineBoxRef.length > 0) {
                const currentX = evt.clientX;
                const deltaX = currentX - startX;
                let newPixelWidth = startWidth + deltaX;

                if (targetOuterSpan) {
                    const $tableCell = $(targetOuterSpan).closest('td, th');
                    if ($tableCell.length > 0) {
                        const parentWidth = $tableCell.width();
                        if (parentWidth > 0) {
                           newPixelWidth = Math.min(newPixelWidth, parentWidth);
                        }
                    }
                }
                newPixelWidth = Math.max(20, newPixelWidth);
                const newPixelHeight = newPixelWidth * currentVisualAspectRatioHW;
                $outlineBoxRef.css({
                    width: newPixelWidth + 'px',
                    height: newPixelHeight + 'px'
                });
            } else {
                console.error('[ep_images_extended mousemove] Outline box ref missing or invalid during size update!');
            }
            $inner.css('cursor', 'nw-resize');
        }
    });

    innerDoc.on('mouseup', function(evt) {
        if (isDragging) {
            const finalX = evt.clientX;
            const deltaX = finalX - startX;
            let finalPixelWidth = startWidth + deltaX;

            if (targetOuterSpan) {
                const $tableCell = $(targetOuterSpan).closest('td, th');
                if ($tableCell.length > 0) {
                    const parentWidth = $tableCell.width();
                     if (parentWidth > 0) {
                        finalPixelWidth = Math.min(finalPixelWidth, parentWidth);
                    }
                }
            }

            finalPixelWidth = Math.max(20, Math.round(finalPixelWidth));
            const widthToApply = `${finalPixelWidth}px`;
            const finalPixelHeight = Math.round(finalPixelWidth * currentVisualAspectRatioHW);
            const heightToApplyPx = `${finalPixelHeight}px`;
            const newCssAspectRatioForVar = (startWidth > 0 && startHeight > 0) ? (startWidth / startHeight).toFixed(4) : '1';

            // Don't apply styles directly to avoid triggering content collector
            // The visual updates will be handled by acePostWriteDomLineHTML after attributes are applied
            // console.log('[ep_images_extended mouseup] Skipping direct style application to avoid content collection triggers');

            _aceContext.callWithAce((ace) => {
                const outerSpanAlive = (targetOuterSpan && document.contains(targetOuterSpan)) ? targetOuterSpan : null;
                let workingOuterSpan = outerSpanAlive;

                // Fallback: locate by active image id if our stored element vanished
                if (!workingOuterSpan && window.epImageInsertActiveImageId) {
                    workingOuterSpan = $inner.find(`.inline-image.image-placeholder[data-image-id="${window.epImageInsertActiveImageId}"]`)[0];
                }

                const placeholderRange = getPlaceholderRangeFromOuterSpan(workingOuterSpan, ace, {wholePlaceholder: true});
                if (!placeholderRange) {
                    console.error('[ep_images_extended mouseup] Could not determine placeholder range for resize.');
                    return;
                }

                if (!validateAceOperation(ace, 'applyAttributes', placeholderRange[0], placeholderRange[1], 'mouseup resize')) {
                    return;
                }

                try {
                    ace.ace_performDocumentApplyAttributesToRange(placeholderRange[0], placeholderRange[1], [
                        ['image-width', widthToApply],
                        ['image-height', heightToApplyPx],
                        ['imageCssAspectRatio', newCssAspectRatioForVar]
                    ]);
                    // console.log('[ep_images_extended mouseup] Successfully applied resize attributes (new targeting)');
                } catch (err) {
                    console.error('[ep_images_extended mouseup] Error applying resize attributes:', err);
                }
            }, 'applyImageAttributes', true);

            // Reset dragging state and clean up
            isDragging = false;
            outlineBoxPositioned = false;
            clickedHandle = null;
            $outlineBoxRef.hide();
            $inner.css('cursor', 'auto');
            
            // Keep the image selected after resizing - don't clear selection
            // Clean up resize position data
            resizePositionData = null;
            
            // Reset target references
            targetOuterSpan = null;
            targetInnerSpan = null;
            targetLineNumber = -1;
        }
        // Note: We don't clear selection here for simple clicks - only the click-outside handler does that
    });

    $inner.on('paste', function(evt) {
        const clipboardData = evt.originalEvent.clipboardData || window.clipboardData;
        if (!clipboardData) return;
        let foundImage = false;
        for (let i = 0; i < clipboardData.items.length; i++) {
            const item = clipboardData.items[i];
            if (item.kind === 'file' && item.type.startsWith('image/')) {
                const file = item.getAsFile();
                foundImage = true; 
                let isValid = true;
                const errorTitle = html10n.get('ep_images_extended.error.title');
                if (clientVars.ep_images_extended && clientVars.ep_images_extended.fileTypes) {
                    const mimedb = clientVars.ep_images_extended.mimeTypes;
                    const mimeTypeInfo = mimedb[file.type];
                    let validMime = false;
                    if (mimeTypeInfo && mimeTypeInfo.extensions) {
                       for (const fileType of clientVars.ep_images_extended.fileTypes) {
                           if (mimeTypeInfo.extensions.includes(fileType)) {
                               validMime = true;
                               break;
                           }
                       }
                    }
                    if (!validMime) {
                       const errorMessage = html10n.get('ep_images_extended.error.fileType');
                       $.gritter.add({ title: errorTitle, text: errorMessage, sticky: true, class_name: 'error' });
                       isValid = false;
                    }
                }
                if (isValid && clientVars.ep_images_extended && file.size > clientVars.ep_images_extended.maxFileSize) {
                   const allowedSize = (clientVars.ep_images_extended.maxFileSize / 1000000);
                   const errorText = html10n.get('ep_images_extended.error.fileSize', { maxallowed: allowedSize });
                   $.gritter.add({ title: errorTitle, text: errorText, sticky: true, class_name: 'error' });
                   isValid = false;
                }
                if (isValid) {
                    evt.preventDefault();

                    // Determine storage strategy (default to s3_presigned)
                    const storageType = (clientVars && clientVars.ep_images_extended && clientVars.ep_images_extended.storageType) || 's3_presigned';

                    // Global cache to avoid re-uploading the same blob within a pad session
                    window.epImageInsertUploadCache = window.epImageInsertUploadCache || {};

                    // Helper to actually insert an <img> (via ace_doInsertImage)
                    const insertIntoPad = (src, widthPx = null, heightPx = null) => {
                        _aceContext.callWithAce((ace) => {
                            ace.ace_doInsertImage(src, widthPx, heightPx);
                        }, 'pasteImage', true);
                    };

                    // Fallback helper: convert blob to base64 and insert
                    const insertAsDataUrl = (blob) => {
                        const readerB64 = new FileReader();
                        readerB64.onload = (e_reader) => {
                            const dataUrl = e_reader.target.result;
                            const probeImg = new Image();
                            probeImg.onload = () => insertIntoPad(dataUrl, `${probeImg.naturalWidth}px`, `${probeImg.naturalHeight}px`);
                            probeImg.onerror = () => insertIntoPad(dataUrl);
                            probeImg.src = dataUrl;
                        };
                        readerB64.onerror = (e_reader) => {
                            console.error('[ep_images_extended paste] FileReader error:', e_reader);
                            $.gritter.add({ title: errorTitle, text: 'Error reading pasted image file.', sticky: true, class_name: 'error' });
                        };
                        readerB64.readAsDataURL(blob);
                    };

                    if (storageType === 'base64') {
                        // Original behaviour retained
                        insertAsDataUrl(file);
                    } else if (storageType === 's3_presigned') {
                        // Upload directly to S3 (or reuse if already uploaded)
                        (async () => {
                            try {
                                // Compute SHA-256 hash for deduplication
                                const arrayBuf = await file.arrayBuffer();
                                const hashBuf = await crypto.subtle.digest('SHA-256', arrayBuf);
                                const hashHex = Array.from(new Uint8Array(hashBuf)).map(b => b.toString(16).padStart(2, '0')).join('');

                                if (window.epImageInsertUploadCache[hashHex]) {
                                    const cachedUrl = window.epImageInsertUploadCache[hashHex];
                                    const probeImg = new Image();
                                    probeImg.onload = () => insertIntoPad(cachedUrl, `${probeImg.naturalWidth}px`, `${probeImg.naturalHeight}px`);
                                    probeImg.onerror = () => insertIntoPad(cachedUrl);
                                    probeImg.src = cachedUrl;
                                    return;
                                }

                                const queryParams = $.param({ name: file.name || `${hashHex}.png`, type: file.type });
                                const presignData = await $.getJSON(`${clientVars.padId}/pluginfw/ep_images_extended/s3_presign?${queryParams}`);
                                if (!presignData || !presignData.signedUrl || !presignData.publicUrl) {
                                    throw new Error('Invalid presign response');
                                }

                                const uploadResp = await fetch(presignData.signedUrl, {
                                    method: 'PUT',
                                    headers: { 'Content-Type': file.type },
                                    body: file,
                                });
                                if (!uploadResp.ok) throw new Error(`S3 upload failed with status ${uploadResp.status}`);

                                const publicUrl = presignData.publicUrl;
                                window.epImageInsertUploadCache[hashHex] = publicUrl;

                                const probeImg = new Image();
                                probeImg.onload = () => insertIntoPad(publicUrl, `${probeImg.naturalWidth}px`, `${probeImg.naturalHeight}px`);
                                probeImg.onerror = () => insertIntoPad(publicUrl);
                                probeImg.src = publicUrl;
                            } catch (err) {
                                console.error('[ep_images_extended paste] S3 upload failed and base64 is disabled:', err);
                                $.gritter.add({ title: errorTitle, text: 'Image upload to S3 failed. Base64 is disabled.', sticky: true, class_name: 'error' });
                            }
                        })();
                    } else {
                        // Generic server upload (local disk etc.)
                        (async () => {
                            try {
                                const formData = new FormData();
                                formData.append('file', file, file.name);
                                const uploadResp = await $.ajax({
                                    type: 'POST',
                                    url: `${clientVars.padId}/pluginfw/ep_images_extended/upload`,
                                    data: formData,
                                    cache: false,
                                    contentType: false,
                                    processData: false,
                                    timeout: 60000,
                                    dataType: 'json',
                                });

                                if (!uploadResp || !uploadResp.url) {
                                    throw new Error('Invalid upload response');
                                }

                                const publicUrl = uploadResp.url;

                                const probeImg = new Image();
                                probeImg.onload = () => insertIntoPad(publicUrl, `${probeImg.naturalWidth}px`, `${probeImg.naturalHeight}px`);
                                probeImg.onerror = () => insertIntoPad(publicUrl);
                                probeImg.src = publicUrl;
                            } catch (err) {
                                console.error('[ep_images_extended paste] Server upload failed and base64 is disabled:', err);
                                $.gritter.add({ title: errorTitle, text: 'Image upload failed. Base64 is disabled.', sticky: true, class_name: 'error' });
                            }
                        })();
                    }
                }
                break; 
            }
        }
        // if (foundImage) { /* handled by preventDefault */ } 
        // else { /* Allow default paste for non-image content */ }
    });

    // Handle clicking outside images to deselect them
    $(innerDoc).on('mousedown', function(evt) {
        if (!$(evt.target).closest('.inline-image.image-placeholder').length) {
            
            if (window.epImageInsertActiveImageId) {
                const previouslyActiveId = window.epImageInsertActiveImageId;
                window.epImageInsertActiveImageId = null;
                // Remove dynamic CSS selection
                const innerDoc = $inner[0].ownerDocument;
                const existingStyle = innerDoc.getElementById('ep-image-selection-style');
                if (existingStyle) {
                    existingStyle.remove();
                }
                // Clean up any pending resize position data
                resizePositionData = null;
                // _aceContext.callWithAce((ace) => {
                //    ace.ace_callRepaint(); // Repaint might not be enough, direct class removal is better.
                // }, 'repaintAfterDeselect', true);
            }
            hideFormatMenu();
        }
    });

    // Handle keyboard events for deselecting images (e.g., Escape key)
    $(innerDoc).on('keydown', function(evt) {
        if (evt.key === 'Escape') {
            if (window.epImageInsertActiveImageId) {
                const previouslyActiveId = window.epImageInsertActiveImageId;
                window.epImageInsertActiveImageId = null;
                // Remove dynamic CSS selection
                const innerDoc = $inner[0].ownerDocument;
                const existingStyle = innerDoc.getElementById('ep-image-selection-style');
                if (existingStyle) {
                    existingStyle.remove();
                }
                // Clean up any pending resize position data
                resizePositionData = null;
                // _aceContext.callWithAce((ace) => {
                //    ace.ace_callRepaint();
                // }, 'repaintAfterDeselectEscape', true);
            }
            hideFormatMenu();
        }
    });
    
    // Handle clicking outside the format menu to hide it
    $(document).on('mousedown', function(evt) {
        if (!$(evt.target).closest('#imageFormatMenu').length && 
            !$(evt.target).closest('.inline-image.image-placeholder').length) {
            hideFormatMenu();
        }
    });
    
    // Function to show user feedback for copy/cut operations
    const showCopyFeedback = (message) => {
        // console.log(`[ep_images_extended] ${message}`);
        // Could be enhanced with a temporary toast notification
    };
    
    // Function to handle image copy/cut operations
    const handleImageCopy = async (shouldCut) => {
        try {
            let currentElement = null;
            if (window.epImageInsertActiveImageId) {
                currentElement = $inner.find(`.inline-image.image-placeholder[data-image-id="${window.epImageInsertActiveImageId}"]`)[0];
            }

            if (!currentElement) {
                console.error('[ep_images_extended copy] No image selected or active image ID not found in DOM.');
                return;
            }
            
            // Get the image source from the selected element
            const classes = currentElement.className.split(' ');
            let imageSrc = null;
            for (const cls of classes) {
                if (cls.startsWith('image:')) {
                    imageSrc = decodeURIComponent(cls.substring(6));
                    break;
                }
            }
            
            if (!imageSrc) {
                console.error('[ep_images_extended copy] Could not find image source');
                return;
            }
            
            // Check if we have clipboard API support
            if (!navigator.clipboard || !window.ClipboardItem) {
                console.error('[ep_images_extended copy] Clipboard API not supported');
                // Fallback: copy the image src as text
                try {
                    await navigator.clipboard.writeText(imageSrc);
                    showCopyFeedback(shouldCut ? 'Image URL cut to clipboard (text fallback)' : 'Image URL copied to clipboard (text fallback)');
                } catch (e) {
                    console.error('[ep_images_extended copy] Fallback text copy failed:', e);
                }
                return;
            }
            
            // Helper function to convert image to PNG blob
            const convertImageToPngBlob = async (imageSrc) => {
                return new Promise((resolve, reject) => {
                    const img = new Image();
                    img.crossOrigin = 'anonymous'; // Handle CORS if needed
                    
                    img.onload = () => {
                        try {
                            // Create a canvas to convert the image to PNG
                            const canvas = document.createElement('canvas');
                            const ctx = canvas.getContext('2d');
                            
                            canvas.width = img.naturalWidth;
                            canvas.height = img.naturalHeight;
                            
                            // Draw the image onto the canvas
                            ctx.drawImage(img, 0, 0);
                            
                            // Convert canvas to PNG blob
                            canvas.toBlob((blob) => {
                                if (blob) {
                                    resolve(blob);
                                } else {
                                    reject(new Error('Failed to convert image to PNG blob'));
                                }
                            }, 'image/png');
                        } catch (error) {
                            reject(error);
                        }
                    };
                    
                    img.onerror = () => {
                        reject(new Error('Failed to load image for conversion'));
                    };
                    
                    img.src = imageSrc;
                });
            };
            
            try {
                let blob;
                
                // For data URLs, convert to PNG regardless of original format
                if (imageSrc.startsWith('data:')) {
                    blob = await convertImageToPngBlob(imageSrc);
                } 
                // For HTTP URLs, fetch and convert to PNG
                else if (imageSrc.startsWith('http')) {
                    // First try to fetch the image
                    let response;
                    try {
                        if (typeof fetchWithCorsProxy === 'function') {
                            response = await fetchWithCorsProxy(imageSrc);
                        } else {
                            response = await fetch(imageSrc, {mode: 'cors'});
                            if (!response.ok) throw new Error(`status ${response.status}`);
                        }
                    } catch (fetchErr) {
                        // Remote fetch failed – fall back to copying the URL as plain text
                        await navigator.clipboard.writeText(imageSrc);
                        showCopyFeedback(shouldCut ? 'Image URL cut to clipboard (fallback)' : 'Image URL copied to clipboard (fallback)');
                        return; // Abort PNG-blob path
                    }
                    
                    // Convert the fetched image to a data URL, then to PNG
                    const originalBlob = await response.blob();
                    const dataUrl = await new Promise((resolve) => {
                        const reader = new FileReader();
                        reader.onload = () => resolve(reader.result);
                        reader.readAsDataURL(originalBlob);
                    });
                    
                    blob = await convertImageToPngBlob(dataUrl);
                }
                // For relative URLs or other formats, try to convert to PNG
                else {
                    blob = await convertImageToPngBlob(imageSrc);
                }
                
                // Create ClipboardItem with the PNG blob
                const clipboardItem = new ClipboardItem({
                    'image/png': blob
                });
                
                // Write to clipboard
                await navigator.clipboard.write([clipboardItem]);
                showCopyFeedback(shouldCut ? 'Image cut to clipboard' : 'Image copied to clipboard');
                
            } catch (e) {
                console.error('[ep_images_extended copy] Failed to copy image as PNG:', e);
                // Fallback to text copy
                try {
                    await navigator.clipboard.writeText(imageSrc);
                    showCopyFeedback(shouldCut ? 'Image URL cut to clipboard (fallback)' : 'Image URL copied to clipboard (fallback)');
                } catch (textError) {
                    console.error('[ep_images_extended copy] Text fallback also failed:', textError);
                    showCopyFeedback('Copy operation failed');
                }
            }
            
            // If this is a cut operation, delete the image after successful copy
            if (shouldCut) {
                // Reuse the existing delete logic
                const outerSpan = currentElement;
                const lineElement = $(outerSpan).closest('.ace-line')[0];
                if (lineElement) {
                    const allImagePlaceholdersInLine = Array.from(lineElement.querySelectorAll('.inline-image.image-placeholder'));
                    const imageIndex = allImagePlaceholdersInLine.indexOf(outerSpan);
                    
                    if (imageIndex !== -1) {
                        const targetLineNumber = _getLineNumberOfElement(lineElement);
                        
                        _aceContext.callWithAce((ace) => {
                            const rep = ace.ace_getRep();
                            if (!rep.lines.atIndex(targetLineNumber)) {
                                console.error(`[ep_images_extended cut] Line ${targetLineNumber} does not exist in rep.`);
                                return;
                            }
                            
                            const lineText = rep.lines.atIndex(targetLineNumber).text;
                            
                            // Use consolidated helper to build placeholder range
                            const placeholderRange = getPlaceholderRangeFromOuterSpan(outerSpan, ace, {wholePlaceholder: true});

                            if (placeholderRange) {
                                const rangeStart = placeholderRange[0];
                                const rangeEnd   = placeholderRange[1];
                                
                                try {
                                    // Use helper function to validate ace operation
                                    if (!validateAceOperation(ace, 'replaceRange', rangeStart, rangeEnd, 'cut')) {
                                        return;
                                    }
                                    
                                    // Delete the image by replacing the text range with empty string
                                    ace.ace_replaceRange(rangeStart, rangeEnd, '');
                                    
                                    // console.log('Successfully cut image at line', targetLineNumber, 'column', placeholderInfo.colStart);
                                    
                                } catch (error) {
                                    console.error('[ep_images_extended cut] Error deleting image:', error);
                                    console.error('[ep_images_extended cut] Range was:', [rangeStart, rangeEnd]);
                                }
                            } else {
                                console.error('[ep_images_extended cut] Could not find placeholder sequence in line text');
                            }
                        }, 'cutImage', true);
                        
                        // Clear selection state since image will be deleted
                        if (window.epImageInsertActiveImageId && $(outerSpan).attr('data-image-id') === window.epImageInsertActiveImageId) {
                            window.epImageInsertActiveImageId = null;
                        }
                        
                        hideFormatMenu();
                    }
                }
            }
            
        } catch (error) {
            console.error('[ep_images_extended copy] Clipboard operation failed:', error);
            showCopyFeedback('Copy operation failed');
        }
    };
    
    // Handle format menu button clicks
    $formatMenuRef.on('click', '.image-format-button', function(evt) {
        evt.preventDefault();
        evt.stopPropagation();
        
        const $button = $(this);
        const wrapType = $button.data('wrap');
        const action = $button.data('action');
        
        let currentElement = null;
        if (window.epImageInsertActiveImageId) {
            currentElement = $inner.find(`.inline-image.image-placeholder[data-image-id="${window.epImageInsertActiveImageId}"]`)[0];
        }

        if (!currentElement) {
            console.warn('[ep_images_extended formatMenu] No active image found for action.');
            return;
        }
        
        if (wrapType) {
            // Handle wrap type buttons
            $formatMenuRef.find('.image-format-button[data-wrap]').removeClass('active');
            $button.addClass('active');
            
            // Use the specific selected image element (like resize logic does)
            // Check both the local reference and the global current reference (for DOM regeneration)
            if (currentElement) {
                const outerSpan = currentElement;
                
                // Determine the float value to store as attribute
                let floatValue;
                switch (wrapType) {
                    case 'inline':
                        floatValue = 'none';
                        break;
                    case 'left':
                        floatValue = 'left';
                        break;
                    case 'right':
                        floatValue = 'right';
                        break;
                    default:
                        floatValue = 'none';
                }
                
                // Apply the attribute using the same method as resize functionality
                const lineElement = $(outerSpan).closest('.ace-line')[0];
                if (lineElement) {
                    const allImagePlaceholdersInLine = Array.from(lineElement.querySelectorAll('.inline-image.image-placeholder'));
                    const imageIndex = allImagePlaceholdersInLine.indexOf(outerSpan);
                    
                    if (imageIndex !== -1) {
                        const targetLineNumber = _getLineNumberOfElement(lineElement);
                        
                        _aceContext.callWithAce((ace) => {
                            const rep = ace.ace_getRep();
                            if (!rep.lines.atIndex(targetLineNumber)) {
                                console.error(`[ep_images_extended float] Line ${targetLineNumber} does not exist in rep.`);
                                return;
                            }
                            
                            const lineText = rep.lines.atIndex(targetLineNumber).text;
                            
                            // Use consolidated helper to build placeholder range
                            const placeholderRange = getPlaceholderRangeFromOuterSpan(outerSpan, ace, {wholePlaceholder: true});

                            if (placeholderRange) {
                                const rangeStart = placeholderRange[0];
                                const rangeEnd   = placeholderRange[1];
                                
                                try {
                                    // Use helper function to validate ace operation
                                    if (!validateAceOperation(ace, 'applyAttributes', rangeStart, rangeEnd, 'float')) {
                                        return;
                                    }
                                    
                                    ace.ace_performDocumentApplyAttributesToRange(rangeStart, rangeEnd, [
                                        ['image-float', floatValue]
                                    ]);
                                    
                                    // console.log('Applied float attribute:', floatValue, 'to image');
                                    
                                } catch (error) {
                                    console.error('[ep_images_extended float] Error applying float attribute:', error);
                                    console.error('[ep_images_extended float] Range was:', [rangeStart, rangeEnd]);
                                    console.error('[ep_images_extended float] Float value was:', floatValue);
                                }
                            } else {
                                console.error('[ep_images_extended float] Could not find placeholder sequence in line text');
                            }
                        }, 'applyImageFloatAttribute', true);
                    }
                }
            }
        } else if (action) {
            // Handle action buttons
            if (action === 'copy') {
                // console.log('Copy image action triggered');
                handleImageCopy(false); // false = copy (don't delete after copy)
            } else if (action === 'cut') {
                // console.log('Cut image action triggered');
                handleImageCopy(true); // true = cut (delete after copy)
            } else if (action === 'delete') {
                // console.log('Delete image action triggered');
                
                // Use the specific selected image element (like float and resize logic does)
                if (currentElement) {
                    const outerSpan = currentElement;
                    _aceContext.callWithAce((ace) => {
                        const placeholderRange = getPlaceholderRangeFromOuterSpan(outerSpan, ace, {wholePlaceholder: true});
                        if (!placeholderRange) {
                            console.error('[ep_images_extended delete] Could not locate placeholder range for deletion.');
                            return;
                        }

                        const [rangeStart, rangeEnd] = placeholderRange;

                        if (!validateAceOperation(ace, 'replaceRange', rangeStart, rangeEnd, 'delete')) {
                            return;
                        }

                        try {
                            ace.ace_replaceRange(rangeStart, rangeEnd, '');
                            // console.log('[ep_images_extended delete] Successfully deleted image via helper');
                        } catch (err) {
                            console.error('[ep_images_extended delete] Error deleting image:', err);
                        }
                    }, 'deleteImage', true);
                }
                
                hideFormatMenu();
            }
        }
    });
  }, 'image_resize_listeners', true);
};

function _getLineNumberOfElement(element) {
  let currentElement = element;
  let count = 0;
  while (currentElement = currentElement.previousElementSibling) {
    count++;
  }
  return count;
}

exports.aceEditorCSS = (hookName, context) => {
  // console.log('[ep_images_extended] aceEditorCSS called - loading CSS file');
  return [
    'ep_images_extended/static/css/ace.css'
  ];
};

exports.aceRegisterBlockElements = () => ['img'];

exports.aceCreateDomLine = (hookName, args, cb) => {
  if (args.cls && args.cls.indexOf('image:') >= 0) {
    const clss = [];
    let imageId = null; 
    const argClss = args.cls.split(' ');
    for (let i = 0; i < argClss.length; i++) {
      const cls = argClss[i];
      clss.push(cls);
      if (cls.startsWith('image-id-')) {
        imageId = cls.substring(9);
      }
    }
    clss.push('inline-image', 'character', 'image-placeholder');
    const handleHtml = 
      '<span class="image-resize-handle br" contenteditable="false"></span>';
    
    // The 'cls' in the modifier will be applied to the *outermost* span ACE creates for the line segment.
    // We will add data-image-id to this span in acePostWriteDomLineHTML.
    const modifier = {
      extraOpenTags: `<span class="image-inner"></span>${handleHtml}`,
      extraCloseTags: '',
      cls: clss.join(' '),
    };
    return cb([modifier]);
  } else {
    return cb();
  }
};

const Changeset = require('ep_etherpad-lite/static/js/Changeset');
exports.acePostWriteDomLineHTML = (hookName, context) => {
  const lineNode = context.node; 
  if (!lineNode) return;

  const placeholders = lineNode.querySelectorAll('span.image-placeholder');
  placeholders.forEach((placeholder, index) => { 
    const outerSpan = placeholder;
    const innerSpan = outerSpan.querySelector('span.image-inner');
    if (!innerSpan) return;

    // Ensure both outer and inner spans are non-editable so cursor keys cannot land inside
    if (!innerSpan.hasAttribute('contenteditable')) {
      innerSpan.setAttribute('contenteditable', 'false');
    }
    if (!outerSpan.hasAttribute('contenteditable')) {
      outerSpan.setAttribute('contenteditable', 'false');
    }

    let escapedSrc = null;
    let imageWidth = null;
    let imageCssAspectRatioVal = null;
    let imageFloatVal = null;
    let imageIdFromClass = null; 

    const classes = outerSpan.className.split(' '); 
    for (const cls of classes) {
      if (cls.startsWith('image:')) {
        escapedSrc = cls.substring(6);
      } else if (cls.startsWith('image-width:')) {
        const widthValue = cls.substring(12);
        if (/\d+px$/.test(widthValue)) {
          imageWidth = widthValue;
        }
      } else if (cls.startsWith('imageCssAspectRatio:')) {
        imageCssAspectRatioVal = cls.substring(20);
      } else if (cls.startsWith('image-float:')) {
        imageFloatVal = cls.substring(12);
      } else if (cls.startsWith('image-id-')) { 
        imageIdFromClass = cls.substring(9);
      }
    }

    if (imageIdFromClass) { 
        outerSpan.setAttribute('data-image-id', imageIdFromClass);
    } else {
        // If it's an old image without an id class from a previous version, try to remove data-image-id if it exists
        if (outerSpan.hasAttribute('data-image-id')) {
            outerSpan.removeAttribute('data-image-id');
        }
    }
    
    const currentDataImageId = outerSpan.getAttribute('data-image-id');

    // acePostWriteDomLineHTML is part of the rendering pipeline, so inline styles should be safe here
    if (imageWidth) {
      innerSpan.style.width = imageWidth;
    } // else { /* Optional: Apply a default width? */ }

    if (imageCssAspectRatioVal) {
      innerSpan.style.setProperty('--image-css-aspect-ratio', imageCssAspectRatioVal);
    } else {
      innerSpan.style.setProperty('--image-css-aspect-ratio', '1');
    }
    innerSpan.style.removeProperty('height');

    if (escapedSrc) {
      try {
        const src = decodeURIComponent(escapedSrc);
        if (src && (src.startsWith('data:') || src.startsWith('http') || src.startsWith('/'))) {
          innerSpan.style.setProperty('--image-src', `url("${src}")`);
        } // else { /* Invalid unescaped src warning removed */ }
      } catch (e) {
        console.error(`[ep_images_extended acePostWriteDomLineHTML] Error setting CSS var for placeholder #${index}:`, e);
      }
    } // else { /* Placeholder found, but no image:* class warning removed */ }
    
    // Apply float style classes based on the attribute
    outerSpan.classList.remove('image-float-left', 'image-float-right', 'image-float-none');
    if (imageFloatVal) {
      switch (imageFloatVal) {
        case 'left':
          outerSpan.classList.add('image-float-left');
          break;
        case 'right':
          outerSpan.classList.add('image-float-right');
          break;
        case 'none':
        case 'inline':
          outerSpan.classList.add('image-float-none');
          break;
      }
    } else {
      // Default to inline/none if no float attribute
      outerSpan.classList.add('image-float-none');
    }
    
    // Selection styling is now handled purely via CSS using CSS custom property
    // No DOM modifications needed here to avoid triggering content collection

    // Remove the old data-image-unique-id if it exists from previous versions
    if (outerSpan.hasAttribute('data-image-unique-id')) {
        outerSpan.removeAttribute('data-image-unique-id');
    }
  });
};

exports.aceAttribClasses = (hook, attr) => {
  return []; 
};

exports.collectContentPost = function(name, context) {
  const node = context.node;
  const state = context.state;
  const tname = context.tname;

  if (tname === 'span' && node && node.classList && node.classList.contains('image-inner')) {
    const innerNode = node;
    // let widthPx = null; // Not needed to initialize here
    // let heightPx = null; // Not needed to initialize here

    if (innerNode.style && innerNode.style.width) {
       const widthMatch = innerNode.style.width.match(/^(\d+)(?:px)?$/);
       if (widthMatch && widthMatch[1]) {
           const widthVal = parseInt(widthMatch[1], 10);
           if (!isNaN(widthVal) && widthVal > 0) {
              let widthToAttrib = `${widthVal}px`;
              if (innerNode.offsetWidth && innerNode.offsetWidth !== widthVal) {
                  widthToAttrib = `${innerNode.offsetWidth}px`;
              }
              state.attribs = state.attribs || {};
              state.attribs['image-width'] = widthToAttrib;
           } // else { /* Parsed width not positive warning removed */ }
       } // else { /* Could not parse width warning removed */ }
    } // else { /* style.width missing warning removed; decision not to delete attribute if style missing */ }

    if (innerNode.style && innerNode.style.height) {
       const heightMatch = innerNode.style.height.match(/^(\d+)(?:px)?$/); 
       if (heightMatch && heightMatch[1]) {
           const heightVal = parseInt(heightMatch[1], 10);
           if (!isNaN(heightVal) && heightVal > 0) {
              state.attribs = state.attribs || {};
              state.attribs['image-height'] = `${heightVal}px`;
           } // else { /* Parsed height not positive warning removed */ }
       } // else { /* Could not parse height warning removed */ }
    } // else { /* style.height missing warning removed */ }

    const computedStyle = window.getComputedStyle(innerNode);
    const cssAspectRatioFromVar = computedStyle.getPropertyValue('--image-css-aspect-ratio');
    if (cssAspectRatioFromVar && cssAspectRatioFromVar.trim() !== '') {
        state.attribs = state.attribs || {};
        state.attribs['imageCssAspectRatio'] = cssAspectRatioFromVar.trim();
    } else {
        if (innerNode.offsetWidth > 0 && innerNode.offsetHeight > 0) {
            const calculatedCssAspectRatio = (innerNode.offsetWidth / innerNode.offsetHeight).toFixed(4);
            state.attribs = state.attribs || {};
            state.attribs['imageCssAspectRatio'] = calculatedCssAspectRatio;
        }
    }
    
    // Preserve float attribute by checking the outer span's classes
    const outerNode = innerNode.parentElement;
    if (outerNode && outerNode.classList) {
        let floatValue = null;
        if (outerNode.classList.contains('image-float-left')) {
            floatValue = 'left';
        } else if (outerNode.classList.contains('image-float-right')) {
            floatValue = 'right';
        } else if (outerNode.classList.contains('image-float-none')) {
            floatValue = 'none';
        }
        
        if (floatValue) {
            state.attribs = state.attribs || {};
            state.attribs['image-float'] = floatValue;
        }
    }

    // NEW: Preserve image-id attribute
    if (outerNode && outerNode.getAttribute('data-image-id')) {
        const imageId = outerNode.getAttribute('data-image-id');
        if (imageId) {
            state.attribs = state.attribs || {};
            state.attribs['image-id'] = imageId;
        }
    }
  }
};

exports.aceKeyEvent = (hookName, context, cb) => {
  const { evt, editorInfo } = context;
  if (evt.type !== 'keydown') return cb(false);
  const key = evt.key;

  // Only special-casing Backspace, Delete, printable characters
  if (!['Backspace', 'Delete'].includes(key) && key.length !== 1) return cb(false);

  const ace = editorInfo;
  const rep = ace.ace_getRep();
  if (!rep || !rep.selStart || !rep.selEnd) return cb(false);

  // Only handle collapsed selections on a single line
  if (rep.selStart[0] !== rep.selEnd[0] || rep.selStart[1] !== rep.selEnd[1]) return cb(false);
  const lineNumber = rep.selStart[0];
  let col = rep.selStart[1];

  const lineObj = rep.lines.atIndex(lineNumber);
  if (!lineObj) return cb(false);
  const lineText = lineObj.text;

  const placeholders = getAllPlaceholderRanges(lineText);
  if (placeholders.length === 0) return cb(false);

  const hit = placeholders.find(r => col >= r.colStart && col <= r.colStart + r.patternLength);
  const afterHit = placeholders.find(r => col === r.colStart + r.patternLength);
  const beforeHit = placeholders.find(r => col === r.colStart);

  // Case 1: cursor inside placeholder – move it to end of placeholder to keep typing out
  if (hit && (col > hit.colStart && col < hit.colStart + hit.patternLength)) {
    ace.ace_performSelectionChange([lineNumber, hit.colStart + hit.patternLength], [lineNumber, hit.colStart + hit.patternLength], false);
    evt.preventDefault();
    return cb(true);
  }

  // Case 2: Backspace immediately after placeholder – delete whole placeholder
  if (key === 'Backspace' && afterHit) {
    ace.ace_replaceRange([lineNumber, afterHit.colStart], [lineNumber, afterHit.colStart + afterHit.patternLength], '');
    evt.preventDefault();
    return cb(true);
  }

  // Case 3: Delete immediately before placeholder – delete whole placeholder
  if (key === 'Delete' && beforeHit) {
    ace.ace_replaceRange([lineNumber, beforeHit.colStart], [lineNumber, beforeHit.colStart + beforeHit.patternLength], '');
    evt.preventDefault();
    return cb(true);
  }

  return cb(false);
};

const doInsertImage = function (src, widthPx, heightPx) {
  const ZWSP = '\u200B';
  const PLACEHOLDER = '\u200B';
  const editorInfo = this.editorInfo;
  const rep = editorInfo.ace_getRep();
  const docMan = this.documentAttributeManager;

  if (!editorInfo || !rep || !rep.selStart || !docMan || !src) {
    console.error('[ep_images_extended doInsertImage] Missing context or src');
    return;
  }

  const cursorPos = rep.selStart;
  const insertText = ZWSP + PLACEHOLDER + ZWSP; // REMOVED trailing space ' '
  
  // Insert the image placeholder text with trailing space
  editorInfo.ace_replaceRange(cursorPos, cursorPos, insertText);

  const imageAttrStart = [cursorPos[0], cursorPos[1] + ZWSP.length];
  const imageAttrEnd = [cursorPos[0], cursorPos[1] + ZWSP.length + PLACEHOLDER.length];
  const escapedSrc = encodeURIComponent(src);
  const attributesToSet = [['image', escapedSrc]];

  if (widthPx && /^\d+px$/.test(widthPx)) {
      attributesToSet.push(['image-width', widthPx]);
  }
  if (heightPx && /^\d+px$/.test(heightPx)) {
      attributesToSet.push(['image-height', heightPx]);
  }
  if (widthPx && heightPx) {
    const naturalWidthNum = parseInt(widthPx, 10);
    const naturalHeightNum = parseInt(heightPx, 10);
    if (naturalWidthNum > 0 && naturalHeightNum > 0) {
        const cssAspectRatio = (naturalWidthNum / naturalHeightNum).toFixed(4);
        attributesToSet.push(['imageCssAspectRatio', cssAspectRatio]);
    }
  }
  
  // NEW: Add image-id attribute
  const imageId = generateUUID();
  attributesToSet.push(['image-id', imageId]);

  // Apply the image attributes
  docMan.setAttributesOnRange(imageAttrStart, imageAttrEnd, attributesToSet);
  
  // CRITICAL FIX: Move cursor after the inserted image to prevent overlapping placeholders
  // This ensures that if user inserts multiple images in sequence, they don't overlap
  const newCursorPos = [cursorPos[0], cursorPos[1] + insertText.length];
  editorInfo.ace_performSelectionChange(newCursorPos, newCursorPos, false);
};

// NEW helper: build a fresh document range for an image placeholder using the DOM element.
// Returns null if it cannot determine the range.
function getPlaceholderRangeFromOuterSpan(outerSpan, ace, opts = {wholePlaceholder: true}) {
  try {
    if (!outerSpan) return null;
    const lineElement = $(outerSpan).closest('.ace-line')[0];
    if (!lineElement) return null;

    const imagePlaceholders = Array.from(lineElement.querySelectorAll('.inline-image.image-placeholder'));
    const imageIndex = imagePlaceholders.indexOf(outerSpan);
    if (imageIndex === -1) return null;

    const lineNumber = _getLineNumberOfElement(lineElement);
    const rep = ace.ace_getRep();
    if (!rep.lines.atIndex(lineNumber)) return null;

    const lineText = rep.lines.atIndex(lineNumber).text;
    const placeholderInfo = findImagePlaceholderPosition(lineText, imageIndex, lineElement);
    if (!placeholderInfo) return null;

    if (opts.wholePlaceholder) {
      return [
        [lineNumber, placeholderInfo.colStart],
        [lineNumber, placeholderInfo.colStart + placeholderInfo.patternLength]
      ];
    }
    // middle-character fallback
    const mid = placeholderInfo.colStart + Math.floor(placeholderInfo.patternLength / 2);
    return [[lineNumber, mid], [lineNumber, mid + 1]];
  } catch (e) {
    console.error('[ep_images_extended] getPlaceholderRangeFromOuterSpan error:', e);
    return null;
  }
}

// Helper: return all placeholder ranges (startCol,length) within a line
function getAllPlaceholderRanges(lineText) {
  const placeholderPatterns = [
    '\u200B\u200B\u200B',      // 3 × ZWSP
    '\u200B\u00A0\u200B',     // ZWSP NBSP ZWSP
    '\u200B\u200B',            // 2 × ZWSP
    '\u00A0',                   // single NBSP
    '\u200B'                    // single ZWSP
  ].map(p => p.replace(/\\u200B/g, '\u200B').replace(/\\u00A0/g, '\u00A0')); // real chars

  const ranges = [];
  let idx = 0;
  while (idx < lineText.length) {
    let matched = false;
    for (const pattern of placeholderPatterns) {
      if (lineText.startsWith(pattern, idx)) {
        ranges.push({colStart: idx, patternLength: pattern.length});
        idx += pattern.length;
        matched = true;
        break;
      }
    }
    if (!matched) idx += 1;
  }
  return ranges;
}

/* ------------------------------------------------------------------
 * Image rendering support for the read-only time-slider view
 * ------------------------------------------------------------------ */

/**
 * Apply inline styles to the image placeholder `outerSpan` so that the image
 * becomes visible in read-only contexts (timeslider or export preview).
 */
function _applyImageStylesForElement(outerSpan) {
  if (!outerSpan) return;
  const innerSpan = outerSpan.querySelector('span.image-inner');
  if (!innerSpan) return;

  // Recover attribute values from the CSS-classes that ACE placed on the span.
  let escSrc = null, width = null, aspect = null, floatVal = null;
  for (const cls of outerSpan.className.split(' ')) {
    if (cls.startsWith('image:')) escSrc = cls.slice(6);
    if (cls.startsWith('image-width:')) width = cls.slice(12);
    if (cls.startsWith('imageCssAspectRatio:')) aspect = cls.slice(20);
    if (cls.startsWith('image-float:')) floatVal = cls.slice(12);
  }

  // Set CSS custom properties / inline styles exactly like acePostWriteDomLineHTML.
  if (escSrc) {
    try {
      const decoded = decodeURIComponent(escSrc);
      innerSpan.style.setProperty('--image-src', `url("${decoded}")`);
    } catch (_) { /* ignore */ }
  }
  if (width) innerSpan.style.width = width;
  if (aspect) innerSpan.style.setProperty('--image-css-aspect-ratio', aspect);

  // Float behaviour (left / right / inline)
  outerSpan.classList.remove('image-float-left', 'image-float-right', 'image-float-none');
  switch (floatVal) {
    case 'left': outerSpan.classList.add('image-float-left'); break;
    case 'right': outerSpan.classList.add('image-float-right'); break;
    default: outerSpan.classList.add('image-float-none');
  }
}

/**
 * Client-side hook that runs in the time-slider once the UI is ready.
 * It ensures all image placeholders are hydrated with the correct styles and
 * repeats that every time the slider jumps to a different revision.
 */
exports.postTimesliderInit = () => {
  // Helper that (re)applies styles to every image currently in the DOM.
  const renderAllImages = () => {
    const $placeholders = $('#innerdocbody').find('span.inline-image.image-placeholder');
    $placeholders.each((_idx, el) => _applyImageStylesForElement(el));
  };

  // Initial render for the first revision shown.
  renderAllImages();

  // Re-render after every slider movement (revision change).
  if (window.BroadcastSlider && typeof window.BroadcastSlider.onSlider === 'function') {
    window.BroadcastSlider.onSlider(() => {
      // Allow the DOM update from broadcast.js to finish first.
      setTimeout(renderAllImages, 0);
    });
  }
};
