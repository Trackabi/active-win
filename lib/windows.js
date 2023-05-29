/* eslint-disable new-cap */
const { spawn } = require('child_process')
const path = require('path')

const EXE_PATH = path.join(__dirname, './active-win.exe')

const tracker = () => spawn(EXE_PATH, [], {})

module.exports = () => tracker()

// const addon = require('./windows-binding.js');
//
// module.exports = async () => addon.getActiveWindow();
//
// module.exports.getOpenWindows = async () => addon.getOpenWindows();
//
// module.exports.sync = addon.getActiveWindow;
//
// module.exports.getOpenWindowsSync = addon.getOpenWindows;
