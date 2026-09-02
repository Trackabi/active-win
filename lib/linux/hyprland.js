import {promisify} from 'node:util';
import childProcess from 'node:child_process';
import {
	buildResult,
	enrichWithProcess,
	enrichWithProcessSync,
} from './process-info.js';

/**
Hyprland backend: `hyprctl activewindow -j` / `hyprctl clients -j`.
*/

const execFile = promisify(childProcess.execFile);

export const name = 'hyprland';

const fromClient = client => {
	if (!client || typeof client !== 'object' || (!client.class && !client.title && !client.pid)) {
		return undefined;
	}

	return buildResult({
		backend: name,
		title: client.title,
		id: Number.parseInt(client.address, 16),
		appId: client.class || client.initialClass,
		name: client.class || client.initialClass,
		processId: client.pid,
		bounds: {
			x: client.at?.[0],
			y: client.at?.[1],
			width: client.size?.[0],
			height: client.size?.[1],
		},
	});
};

const parse = stdout => {
	try {
		return JSON.parse(stdout);
	} catch {
		return undefined;
	}
};

export async function activeWindow() {
	const {stdout} = await execFile('hyprctl', ['activewindow', '-j']);
	const result = fromClient(parse(stdout));
	return result && enrichWithProcess(result);
}

export function activeWindowSync() {
	const stdout = childProcess.execFileSync('hyprctl', ['activewindow', '-j'], {encoding: 'utf8'});
	const result = fromClient(parse(stdout));
	return result && enrichWithProcessSync(result);
}

export async function openWindows() {
	const {stdout} = await execFile('hyprctl', ['clients', '-j']);
	const clients = parse(stdout);
	if (!Array.isArray(clients)) {
		return [];
	}

	return Promise.all(clients.map(client => fromClient(client)).filter(Boolean).map(result => enrichWithProcess(result)));
}

export function openWindowsSync() {
	const stdout = childProcess.execFileSync('hyprctl', ['clients', '-j'], {encoding: 'utf8'});
	const clients = parse(stdout);
	if (!Array.isArray(clients)) {
		return [];
	}

	return clients.map(client => fromClient(client)).filter(Boolean).map(result => enrichWithProcessSync(result));
}
