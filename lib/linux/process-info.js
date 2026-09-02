import fs from 'node:fs';
import {promisify} from 'node:util';

const readFile = promisify(fs.readFile);
const readlink = promisify(fs.readlink);

export async function getMemoryUsageByPid(pid) {
	const statm = await readFile(`/proc/${pid}/statm`, 'utf8');
	return Number.parseInt(statm.split(' ')[1], 10) * 4096;
}

export function getMemoryUsageByPidSync(pid) {
	const statm = fs.readFileSync(`/proc/${pid}/statm`, 'utf8');
	return Number.parseInt(statm.split(' ')[1], 10) * 4096;
}

export const getPathByPid = pid => readlink(`/proc/${pid}/exe`);

export function getPathByPidSync(pid) {
	try {
		return fs.readlinkSync(`/proc/${pid}/exe`);
	} catch {}
}

export async function getCommByPid(pid) {
	try {
		const comm = await readFile(`/proc/${pid}/comm`, 'utf8');
		return comm.trim();
	} catch {}
}

export function getCommByPidSync(pid) {
	try {
		return fs.readFileSync(`/proc/${pid}/comm`, 'utf8').trim();
	} catch {}
}

/**
Turn a Wayland `app_id` / `wm_class` into the short app name the X11 backend
has always reported (`WM_CLASS` class part, e.g. `firefox`, `Code`).

Wayland-native and Flatpak apps identify themselves with reverse-DNS ids
(`org.mozilla.firefox`, `org.gnome.Nautilus`); consumers key their app rules
on the short form, so map those to their last segment.
*/
export function normalizeAppName(raw) {
	if (!raw) {
		return '';
	}

	let name = String(raw).trim().replace(/\.desktop$/, '');
	if (!name.includes(' ') && /^[\w-]+(?:\.[\w-]+){2,}$/.test(name)) {
		name = name.split('.').pop();
	}

	return name;
}

const toInt = value => {
	const number = Number.parseInt(value, 10);
	return Number.isNaN(number) ? undefined : number;
};

/**
Build the shared result object; `processId`-derived fields are filled in by
`enrichWithProcess` / `enrichWithProcessSync`.
*/
export function buildResult({backend, title, id, appId, name, processId, bounds}) {
	const pid = toInt(processId);
	return {
		platform: 'linux',
		backend,
		title: title ?? '',
		id: toInt(id) ?? 0,
		owner: {
			name: normalizeAppName(name || appId),
			processId: pid,
			path: undefined,
			appId: appId || undefined,
		},
		bounds: {
			x: toInt(bounds?.x) ?? 0,
			y: toInt(bounds?.y) ?? 0,
			width: toInt(bounds?.width) ?? 0,
			height: toInt(bounds?.height) ?? 0,
		},
		memoryUsage: undefined,
	};
}

export async function enrichWithProcess(result) {
	const pid = result.owner.processId;
	if (!pid) {
		return result;
	}

	const [memoryUsage, path, comm] = await Promise.all([
		getMemoryUsageByPid(pid).catch(() => undefined),
		getPathByPid(pid).catch(() => undefined),
		result.owner.name ? undefined : getCommByPid(pid),
	]);
	result.memoryUsage = memoryUsage;
	result.owner.path = path;
	if (!result.owner.name && comm) {
		result.owner.name = comm;
	}

	return result;
}

export function enrichWithProcessSync(result) {
	const pid = result.owner.processId;
	if (!pid) {
		return result;
	}

	try {
		result.memoryUsage = getMemoryUsageByPidSync(pid);
	} catch {}

	result.owner.path = getPathByPidSync(pid);
	result.owner.name ||= getCommByPidSync(pid) || '';

	return result;
}
