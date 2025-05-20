import {spawn} from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs';
import {fileURLToPath} from 'node:url';
import {createRequire} from 'node:module';
import preGyp from '@mapbox/node-pre-gyp';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const isPackaged = process.mainModule?.filename.includes('app.asar');
const EXE_PATH = isPackaged
	? path.join(process.resourcesPath, 'app.asar.unpacked', 'node_modules', 'active-win', 'lib', 'active-win.exe')
	: path.join(__dirname, './active-win.exe');

const getAddon = () => {
	const require = createRequire(import.meta.url);

	const bindingPath = preGyp.find(path.resolve(path.join(__dirname, '../package.json')));

	return (fs.existsSync(bindingPath)) ? require(bindingPath) : {
		getActiveWindow() {},
		getOpenWindows() {},
	};
};

export async function activeWindow() {
	return () => spawn(EXE_PATH, [],  {
        stdio: ['ignore', 'pipe', 'inherit'],
    });
}

export function activeWindowSync() {
	return getAddon().getActiveWindow();
}

export function openWindows() {
	return getAddon().getOpenWindows();
}

export function openWindowsSync() {
	return getAddon().getOpenWindows();
}
