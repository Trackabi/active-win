import process from 'node:process';
import {promisify} from 'node:util';
import childProcess from 'node:child_process';
import {
	buildResult,
	enrichWithProcess,
	enrichWithProcessSync,
} from './process-info.js';

/**
X11 backend (`xprop` + `xwininfo`).

On a Wayland session the app runs under XWayland, so this still answers for
X11 clients but reports "no window" (`_NET_ACTIVE_WINDOW` = 0) whenever a
native Wayland window has focus. It is therefore the last resort there.
*/

const execFile = promisify(childProcess.execFile);

const xpropBinary = 'xprop';
const xwininfoBinary = 'xwininfo';
const xpropActiveArguments = ['-root', '\t$0', '_NET_ACTIVE_WINDOW'];
const xpropOpenArguments = ['-root', '_NET_CLIENT_LIST_STACKING'];
const xpropDetailsArguments = ['-id'];
const xpropEnvironment = {...process.env, LC_ALL: 'C.utf8'};

export const name = 'x11';

const processOutput = output => {
	const result = {};

	for (const row of output.trim().split('\n')) {
		if (row.includes('=')) {
			const [key, ...valueParts] = row.split('=');
			const value = valueParts.join('=');
			result[key.trim()] = value.trim();
		} else if (row.includes(':')) {
			const [key, ...valueParts] = row.split(':');
			const value = valueParts.join(':');
			result[key.trim()] = value.trim();
		}
	}

	return result;
};

const parseQuoted = value => {
	if (value === undefined) {
		return undefined;
	}

	try {
		return JSON.parse(value);
	} catch {
		return value.replaceAll(/^"|"$/g, '');
	}
};

export const parseLinux = ({stdout, boundsStdout, activeWindowId}) => {
	const result = processOutput(stdout);
	const bounds = processOutput(boundsStdout);

	const windowIdProperty = 'WM_CLIENT_LEADER(WINDOW)';
	const resultKeys = Object.keys(result);
	const windowId = (resultKeys.indexOf(windowIdProperty) > 0
		&& Number.parseInt(result[windowIdProperty].split('#').pop(), 16)) || activeWindowId;

	// WM_CLASS is "instance", "class"; the class part is the app name.
	const wmClass = result['WM_CLASS(STRING)'];
	const className = wmClass ? parseQuoted(wmClass.split(',').pop().trim()) : '';

	return buildResult({
		backend: name,
		title: parseQuoted(result['_NET_WM_NAME(UTF8_STRING)'] ?? result['WM_NAME(STRING)']) ?? null,
		id: windowId,
		name: className,
		appId: className || undefined,
		processId: result['_NET_WM_PID(CARDINAL)'],
		bounds: {
			x: bounds['Absolute upper-left X'],
			y: bounds['Absolute upper-left Y'],
			width: bounds.Width,
			height: bounds.Height,
		},
	});
};

const getActiveWindowId = activeWindowIdStdout => Number.parseInt(activeWindowIdStdout.split('\t')[1], 16);

async function getWindowInformation(windowId) {
	const [{stdout}, {stdout: boundsStdout}] = await Promise.all([
		execFile(xpropBinary, [...xpropDetailsArguments, windowId], {env: xpropEnvironment}),
		execFile(xwininfoBinary, [...xpropDetailsArguments, windowId]),
	]);

	return enrichWithProcess(parseLinux({activeWindowId: windowId, boundsStdout, stdout}));
}

function getWindowInformationSync(windowId) {
	const stdout = childProcess.execFileSync(xpropBinary, [...xpropDetailsArguments, windowId], {encoding: 'utf8', env: xpropEnvironment});
	const boundsStdout = childProcess.execFileSync(xwininfoBinary, [...xpropDetailsArguments, windowId], {encoding: 'utf8'});

	return enrichWithProcessSync(parseLinux({activeWindowId: windowId, boundsStdout, stdout}));
}

/**
Unlike the historical implementation these throw when `xprop` itself is
unusable (missing binary, no DISPLAY), so the caller can tell "backend is
broken" apart from "no window is focused" (`undefined`).
*/
export async function activeWindow() {
	const {stdout: activeWindowIdStdout} = await execFile(xpropBinary, xpropActiveArguments);
	const activeWindowId = getActiveWindowId(activeWindowIdStdout);

	if (!activeWindowId) {
		return undefined;
	}

	try {
		return await getWindowInformation(activeWindowId);
	} catch {
		// The window may have been closed between the two calls.
		return undefined;
	}
}

export function activeWindowSync() {
	const activeWindowIdStdout = childProcess.execFileSync(xpropBinary, xpropActiveArguments, {encoding: 'utf8'});
	const activeWindowId = getActiveWindowId(activeWindowIdStdout);

	if (!activeWindowId) {
		return undefined;
	}

	try {
		return getWindowInformationSync(activeWindowId);
	} catch {
		return undefined;
	}
}

const parseWindowIds = stdout => {
	const [, list] = stdout.split('#');
	if (!list) {
		return [];
	}

	return list.trim().replace('\n', '').split(',').map(id => Number.parseInt(id, 16)).filter(Boolean);
};

export async function openWindows() {
	const {stdout} = await execFile(xpropBinary, xpropOpenArguments);
	const windowIds = parseWindowIds(stdout);
	const windows = [];

	for (const windowId of windowIds) {
		try {
			// eslint-disable-next-line no-await-in-loop
			windows.push(await getWindowInformation(windowId));
		} catch {}
	}

	return windows;
}

export function openWindowsSync() {
	const stdout = childProcess.execFileSync(xpropBinary, xpropOpenArguments, {encoding: 'utf8'});
	const windows = [];

	for (const windowId of parseWindowIds(stdout)) {
		try {
			windows.push(getWindowInformationSync(windowId));
		} catch {}
	}

	return windows;
}
