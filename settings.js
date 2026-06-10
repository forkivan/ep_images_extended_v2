// Modified from ep_image_insert 1.0.7 

'use strict';

exports.loadSettings = (hookName, args, cb) => {
  // Sync ep_images_extended config into the runtime Settings singleton that other
  // parts of this plugin import, to avoid workspace/symlink doubletons.
  try {
    const settingsModule = require('ep_etherpad-lite/node/utils/Settings');
    const runtimeSettings = settingsModule.default || settingsModule;
    if (args && args.settings && args.settings.ep_images_extended) {
      runtimeSettings.ep_images_extended = args.settings.ep_images_extended;
    }
  } catch (e) {
    console.warn('[ep_images_extended] Failed to sync settings:', e);
  }

  if (!args.settings || !args.settings.socketIo) {
    console.warn('Please update Etherpad to >=1.8.8');
  } else {
    // Setting maxHttpBufferSize to 10 MiB :)
    args.settings.socketIo.maxHttpBufferSize = 100000000;
  }
  cb();
};
